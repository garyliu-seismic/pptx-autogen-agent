"""跨 pptx 整页幻灯片拼接（OPC part 级克隆），用于把 SRDP ExtContent 里的库内容
（外部/别的品牌 pptx 里的幻灯片）原样拼进 auto agent 生成的 deck。

python-pptx 没有官方的“跨 presentation 复制 slide”API，这里直接在 OPC
(part + relationship 图)层面操作，而不是手工拼 zip/XML：

  1. 从源 slide part 出发，沿它的 ``.rels`` 递归克隆所有依赖 part
     （slideLayout -> slideMaster -> theme、图片/图表/内嵌 xlsx 等媒体），
     在目标 package 里分配不冲突的新 partname，内容原样复制（对 XmlPart
     是"重新序列化再解析"，等价于深拷贝，不与源共享可变状态）。
  2. 每个 part 自己 rels 里的关系 Id 在克隆时会被重新分配（新 part 从零开始
     分配 rId），所以复制完 rels 后要把 XML 里所有 r:id/r:embed/... 属性按
     "旧 rId -> 新 rId" 的映射改写；映射里没有的（详见 _SKIP_RELTYPES）则直接
     删掉该属性，避免留下悬空引用。
  3. 新克隆出来的 slide 挂到目标 presentation 的 sldIdLst；如果这次连带克隆出了
     新的 slideMaster，还要把它注册进 sldMasterIdLst（否则该 master 虽然可以
     通过 slide->layout->master 关系链被发现并序列化，但不会出现在 PowerPoint
     的"幻灯片母版"视图里）。

已知不处理（超出本次范围，遇到时会被静默丢弃，不影响该页其余内容正常拼入）：
notesSlide/notesMaster、slide 内的"跳转到某一页"超链接（reltype=slide）、
customXml（见下）。
"""
from __future__ import annotations

import os
import re
import uuid

from pptx import Presentation
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, PartFactory
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.parts.slide import SlideMasterPart

_R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_A16_NS = "{http://schemas.microsoft.com/office/drawing/2014/main}"
_ASVG_NS = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}"
# 递归克隆时跳过的关系类型：
#   - notesSlide：备注页，不影响视觉展示
#   - slide：slide 内"跳转到某一页"超链接，指向的是源文件里别的页，不该带过来
#   - customXml：真实复现过的"需要修复"来源之一——这类 customXml 往往是设计
#     工具自己的版式目录/元数据（比如某个 layout 的 Name/UniqueId 记录），只
#     克隆到"这一页用到的 master 连带的全部 layout"时，会带出一整批这种目录条
#     目，却没有它们所属的目录/索引部件（那个是挂在源文件别的地方，我们没有
#     整份克隆源 presentation），PowerPoint 打开时检测出这批 customXml 不完整
#     就整批清掉触发"需要修复"。这些纯粹是设计工具的元数据，不影响幻灯片实际
#     视觉内容，直接跳过不克隆。
_SKIP_RELTYPES = {RT.NOTES_SLIDE, RT.SLIDE, RT.CUSTOM_XML}


def splice_external_slides(dst_prs, src_pptx_path: str,
                           slide_indices: list[int] | None = None) -> list:
    """把 src_pptx_path 里指定的若干页（默认全部）整页原样克隆拼到 dst_prs 末尾。

    返回新增的 Slide 对象列表（按传入顺序）。同一次调用共享一个克隆缓存，
    因此同一源文件里多页共用的 layout/master/theme/图片只会被克隆一份。
    """
    src_prs = Presentation(src_pptx_path)
    if slide_indices is None:
        slide_indices = list(range(len(src_prs.slides)))
    memo: dict[str, Part] = {}
    alloc = _PartnameAllocator(dst_prs.part.package)
    return [_import_slide(dst_prs, src_prs.slides[i], memo, alloc) for i in slide_indices]


def fix_app_properties_slide_count(prs) -> None:
    """把 docProps/app.xml 里的 <Slides>/<Notes> 计数改成实际页数。

    这个 part python-pptx 完全不管——用 --reference-deck 打开模板时，它是
    从模板原样带过来的（比如 MainLiveDoc 原本 67 页留下的 <Slides>67</Slides>），
    之后不管是渲染器生成新页、还是 remove_slide/splice_external_slides 改了
    页数，这个计数都不会跟着变。PowerPoint 会拿这两个数字跟实际 sldIdLst 页数
    做健全性检查，真实遇到过：数字不一致但 schema 本身完全合法（validator
    检查不出来），PowerPoint 打开仍提示"需要修复"。"""
    n = len(prs.slides)
    for part in prs.part.package.iter_parts():
        if str(part.partname) != "/docProps/app.xml":
            continue
        xml = part.blob.decode("utf-8")
        xml = re.sub(r"<Slides>\d+</Slides>", f"<Slides>{n}</Slides>", xml)
        xml = re.sub(r"<Notes>\d+</Notes>", f"<Notes>{n}</Notes>", xml)
        part.blob = xml.encode("utf-8")
        return


def remove_slide(prs, idx: int) -> None:
    """删掉 prs 的第 idx 页（python-pptx 没有官方删除 API的标准写法：
    从 sldIdLst 摘掉引用，再把 presentation part 对它的关系一并丢掉）。"""
    pres_part = prs.part
    sldIdLst = pres_part._element.get_or_add_sldIdLst()
    sldId_el = sldIdLst.sldId_lst[idx]
    rId = sldId_el.rId
    sldIdLst.remove(sldId_el)
    pres_part.drop_rel(rId)


class _PartnameAllocator:
    """给克隆出的 part 分配不冲突的新 partname。

    不能用 ``package.next_partname()``：它靠 ``iter_parts()`` 现场遍历目标包的
    关系图来判断"名字被占用了没"，而克隆过程中新建的 part 要等父 part 之后
    调用 relate_to() 才会挂进关系图——同一批递归克隆里先后创建的两个"孤儿" part
    互相看不到对方，会分到同一个名字。这里改成在整次拼接开始时只扫描一次已占用
    的 partname，之后本地维护，自己分配的立刻记下来。

    占用判定按大小写不敏感比较：OPC 规范要求 part name 唯一性按大小写不敏感
    判断（.NET 的 System.IO.Packaging 会严格执行，PowerPoint 打开时报"需要修复"）。
    真实遇到过：源 pptx 用 "customXml/item291.xml"，目标模板自己用的是
    "customXML/item291.xml"（大小写不同），大小写敏感比较会误判为两个不同名字，
    实际在 OPC 层面是同一个 part name 冲突。
    """

    def __init__(self, package):
        self._used = {str(p.partname).lower() for p in package.iter_parts()}

    def next(self, template: str) -> str:
        n = 1
        while True:
            candidate = template % n
            if candidate.lower() not in self._used:
                self._used.add(candidate.lower())
                return candidate
            n += 1


def _import_slide(dst_prs, src_slide, memo: dict[str, Part], alloc: "_PartnameAllocator"):
    dst_package = dst_prs.part.package
    new_slide_part = _clone_dependency_part(dst_package, memo, src_slide.part, alloc)

    pres_part = dst_prs.part
    rId = pres_part.relate_to(new_slide_part, RT.SLIDE)
    sldIdLst = pres_part._element.get_or_add_sldIdLst()
    sldIdLst.add_sldId(rId)
    new_sld_id = sldIdLst.sldId_lst[-1].get("id")

    for part in memo.values():
        if isinstance(part, SlideMasterPart):
            _ensure_master_registered(dst_prs, part)

    _sync_section_list(pres_part._element, new_sld_id)

    return dst_prs.slides[-1]


def _clone_dependency_part(dst_package, memo: dict[str, Part], src_part: Part,
                           alloc: "_PartnameAllocator") -> Part:
    key = str(src_part.partname)
    if key in memo:
        return memo[key]

    new_partname = PackURI(alloc.next(_partname_template(src_part.partname)))
    new_part = PartFactory(new_partname, src_part.content_type, dst_package, src_part.blob)
    memo[key] = new_part

    old_to_new: dict[str, str | None] = {}
    for old_rid, rel in src_part.rels.items():
        if rel.reltype in _SKIP_RELTYPES:
            old_to_new[old_rid] = None
            continue
        if rel.is_external:
            old_to_new[old_rid] = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            cloned_target = _clone_dependency_part(dst_package, memo, rel.target_part, alloc)
            old_to_new[old_rid] = new_part.relate_to(cloned_target, rel.reltype)

    if hasattr(new_part, "_element"):
        _regenerate_creation_ids(new_part._element)
        _remap_rids(new_part._element, old_to_new)
        _add_svg_raster_fallback(new_part, alloc)

    return new_part


def _add_svg_raster_fallback(part: Part, alloc: "_PartnameAllocator") -> None:
    """源模板里常见一种"只给 SVG、外层 a:blip 没有 r:embed/r:link"的写法
    （`<a:blip><a:extLst>...<asvg:svgBlip r:embed="..."/>` 但外层 blip 自己
    没有任何直接可解析的属性）——本身不是这次拼接引入的问题（源文件里就长
    这样），但没有 raster fallback 的 blip 被怀疑是 PowerPoint 判定"需要修复"
    的诱因之一。这里做个兜底：给这类 blip 补一份栅格化（PNG）fallback 并挂到
    外层 r:embed 上，跟 PowerPoint 自己插入 SVG 图片时生成的结构对齐。
    栅格化失败（缺依赖、svg 解析不了等）不阻断整个拼接流程，跳过即可。"""
    for blip in part._element.iter(_A_NS + "blip"):
        if blip.get(_R_NS + "embed") or blip.get(_R_NS + "link"):
            continue
        svg_blip = blip.find(f".//{_ASVG_NS}svgBlip")
        if svg_blip is None:
            continue
        svg_rid = svg_blip.get(_R_NS + "embed")
        if not svg_rid or svg_rid not in part.rels:
            continue
        svg_part = part.rels[svg_rid].target_part
        png_bytes = _rasterize_svg_to_png(svg_part.blob)
        if png_bytes is None:
            continue
        png_partname = PackURI(alloc.next("/ppt/media/image%d.png"))
        png_part = PartFactory(png_partname, CT.PNG, part.package, png_bytes)
        png_rid = part.relate_to(png_part, RT.IMAGE)
        blip.set(_R_NS + "embed", png_rid)


def _rasterize_svg_to_png(svg_bytes: bytes) -> bytes | None:
    try:
        import io
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPM
    except ImportError:
        return None
    try:
        drawing = svg2rlg(io.BytesIO(svg_bytes))
        if drawing is None:
            return None
        buf = io.BytesIO()
        renderPM.drawToFile(drawing, buf, fmt="PNG")
        return buf.getvalue()
    except Exception:
        return None


_P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_P14_NS = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"


def _sync_section_list(presentation_element, new_sld_id: str) -> None:
    """PowerPoint 的"节(Section)"功能在 ``p:extLst`` 里维护一份跟
    ``p:sldIdLst`` 平行、但独立的 ``p14:sldIdLst``（按 section 分组）。这份列表
    schema 上没有强制要求跟主列表的页集合一致，但 PowerPoint 自己打开时会做
    这个一致性检查——新拼入的页只挂到主 sldIdLst、没同步进这里，会被判定成
    "需要修复"。挂到最后一个 section（没有 section 就什么都不做，说明这份
    文件本来就没用节功能）。"""
    ext_lst = presentation_element.find(qn("p:extLst"))
    if ext_lst is None:
        return
    section_lst = ext_lst.find(f".//{_P14_NS}sectionLst")
    if section_lst is None:
        return
    sections = section_lst.findall(f"{_P14_NS}section")
    if not sections:
        return
    last_section_sld_id_lst = sections[-1].find(f"{_P14_NS}sldIdLst")
    if last_section_sld_id_lst is None:
        return
    new_el = last_section_sld_id_lst.makeelement(_P14_NS + "sldId", {})
    new_el.set("id", new_sld_id)
    last_section_sld_id_lst.append(new_el)


def _regenerate_creation_ids(element) -> None:
    """跨包克隆会把源文件的 a16:creationId 原样带过来。这些 GUID 本该在整个
    生态里全局唯一，但源文件本身可能已经重复（常见于早年在 PowerPoint 里复制
    粘贴 slide/shape 时没有触发新 GUID 生成），克隆到目标包后又可能跟目标包
    自己的内容或另一次拼接带入的内容撞上。不管撞车的根因具体是不是这个，克隆
    时统一分配全新 GUID 都是标准且无风险的做法（PowerPoint 自己复制粘贴 slide
    时也是这么做的），直接从根上消除这一整类冲突。"""
    for el in element.iter(_A16_NS + "creationId"):
        el.set("id", "{%s}" % str(uuid.uuid4()).upper())


def _remap_rids(element, mapping: dict[str, str | None]) -> None:
    to_remove = []
    for el in list(element.iter()):
        for key, val in list(el.attrib.items()):
            if key.startswith(_R_NS) and val in mapping:
                new_val = mapping[val]
                if new_val is not None:
                    el.set(key, new_val)
                elif el.tag == _P_NS + "custData":
                    # r:id 在 <p:custData> 上是必填属性——不能只删属性留一个
                    # 空壳元素（会变成 schema 非法），整个元素一起删掉。
                    to_remove.append(el)
                else:
                    del el.attrib[key]

    for el in to_remove:
        parent = el.getparent()
        if parent is None:
            continue
        parent.remove(el)
        # custDataLst 被删空了就连它一起删，不留一个没有子元素的空壳。
        if parent.tag == _P_NS + "custDataLst" and len(parent) == 0:
            grandparent = parent.getparent()
            if grandparent is not None:
                grandparent.remove(parent)


def _partname_template(partname) -> str:
    """把 "/ppt/media/image7.png" 这类 partname 变成 "/ppt/media/image%d.png"
    模板，交给 ``package.next_partname()`` 去分配下一个不冲突的编号。"""
    s = str(partname)
    m = re.match(r"^(.*?)(\d+)(\.[^./]+)$", s)
    if m:
        return f"{m.group(1)}%d{m.group(3)}"
    root, ext = os.path.splitext(s)
    return f"{root}_ext%d{ext}"


def _ensure_master_registered(dst_prs, master_part: SlideMasterPart) -> None:
    """把（可能新克隆出的）slideMaster 挂到 presentation 的 sldMasterIdLst，
    已经挂过则跳过（同一 master 可能被多页共享）。

    真实遇到过、且是长期"需要修复"调查的最终根因：新 id 只在
    ``p:sldMasterId`` 这一份列表里取 max()+1，但 PowerPoint 实际把
    ``p:sldMasterId`` 和每个 master 自己 ``p:sldLayoutIdLst`` 里的
    ``p:sldLayoutId`` id 当成同一个共享的 32 位 id 空间——只避开
    sldMasterId 自己的最大值，算出来的"下一个"号码很容易撞上某个
    完全不相关的、已存在 master 的 layout id（两者是各自独立分配的，
    纯属巧合碰撞），PowerPoint 打开时判定 id 冲突触发"需要修复"。
    这里改成先扫描 dst 包里所有 slideMaster part 自己的
    sldLayoutIdLst，把这些 id 也纳入取 max 的范围。"""
    pres_part = dst_prs.part
    for rel in pres_part.rels.values():
        if not rel.is_external and rel.reltype == RT.SLIDE_MASTER \
                and rel.target_part is master_part:
            return

    rId = pres_part.relate_to(master_part, RT.SLIDE_MASTER)
    sldMasterIdLst = pres_part._element.get_or_add_sldMasterIdLst()
    existing_ids = [int(el.get("id")) for el in sldMasterIdLst.sldMasterId_lst]
    for part in pres_part.package.iter_parts():
        if isinstance(part, SlideMasterPart):
            existing_ids.extend(
                int(v) for v in part._element.xpath(".//p:sldLayoutIdLst/p:sldLayoutId/@id")
            )
    next_id = max(existing_ids, default=2147483647) + 1
    new_el = sldMasterIdLst.makeelement(qn("p:sldMasterId"), {})
    new_el.set("id", str(next_id))
    new_el.set(qn("r:id"), rId)
    sldMasterIdLst.append(new_el)
