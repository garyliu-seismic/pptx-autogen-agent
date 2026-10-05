"""PPTX 动画注入器 —— 通过直接操作 Office Open XML 给幻灯片添加进入动画。

Root cause fix：原版用 p:animEffect filter='fly_in'/'float_in'/'zoom' 等非法值，
导致 PowerPoint 打开时弹出 "found a problem with content" 修复对话框。

正确实现：
  - fade/wipe/split  → p:animEffect + 合法 filter 值 (OOXML ST_TLAnimateEffectTransition)
  - fly_in/float_in  → p:animMotion (位移路径动画，从底部或下方飞入)
  - zoom             → p:animScale (缩放动画，从 0 放大到 100%)
  - appear           → p:set only (style.visibility: hidden→visible，无过渡)

合法 animEffect filter 值（OOXML ST_TLAnimateEffectTransition）：
  fade, blinds, box, checkerboard, circle, dissolve, randomBar,
  strips, wedge, wheel, wipe, zoom, split, plus, diamond, cross

presetID / presetClass 对应关系（PowerPoint 内部枚举）：
  1=appear, 2=flyIn, 10=fade, 26=magnify(zoom), 30=split, 36=wipe
"""
from __future__ import annotations

from lxml import etree

PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"

_DUR_ENTER  = 500   # ms 默认动画时长
_DUR_DELAY  = 150   # ms cascade 间隔

# ---------------------------------------------------------------------------
# 效果定义表
# type: 'animEffect' | 'animMotion' | 'animScale' | 'set_only'
# ---------------------------------------------------------------------------
_EFFECTS: dict[str, dict] = {
    "fade": {
        "type": "animEffect",
        "presetID": "10", "presetClass": "entr", "presetSubtype": "0",
        "filter": "fade",
    },
    "wipe": {
        "type": "animEffect",
        "presetID": "36", "presetClass": "entr", "presetSubtype": "1",
        "filter": "wipe",
    },
    "split": {
        "type": "animEffect",
        "presetID": "30", "presetClass": "entr", "presetSubtype": "4",
        "filter": "split",
    },
    "fly_in": {
        # presetID=2 fly_in from bottom → p:animMotion path M0,1 L0,0
        "type": "animMotion",
        "presetID": "2", "presetClass": "entr", "presetSubtype": "8",
        "path": "M 0 1 L 0 0 E",
    },
    "float_in": {
        # presetID=2 fly_in from slight bottom (float) → similar path, shorter
        "type": "animMotion",
        "presetID": "2", "presetClass": "entr", "presetSubtype": "0",
        "path": "M 0 0.25 L 0 0 E",
    },
    "zoom": {
        # presetID=26 magnify/zoom → p:animScale from 0→100%
        "type": "animScale",
        "presetID": "26", "presetClass": "entr", "presetSubtype": "0",
    },
    "appear": {
        # presetID=1 appear → set only, no transition
        "type": "set_only",
        "presetID": "1", "presetClass": "entr", "presetSubtype": "0",
    },
}


# ---------------------------------------------------------------------------
# 页型 → 动画配置
# ---------------------------------------------------------------------------
KIND_ANIMATION: dict[str, dict] = {
    "title":      {"effect": "fade",    "cascade": False, "dur": 800},
    "section":    {"effect": "fade",    "cascade": False, "dur": 700},
    "agenda":     {"effect": "fly_in",  "cascade": True,  "dur": 500},
    "bullets":    {"effect": "fly_in",  "cascade": True,  "dur": 500},
    "cards":      {"effect": "zoom",    "cascade": True,  "dur": 450},
    "timeline":   {"effect": "wipe",    "cascade": True,  "dur": 500},
    "kpi":        {"effect": "zoom",    "cascade": True,  "dur": 500},
    "process":    {"effect": "fly_in",  "cascade": True,  "dur": 450},
    "comparison": {"effect": "wipe",    "cascade": False, "dur": 600},
    "quote":      {"effect": "fade",    "cascade": False, "dur": 900},
    "summary":    {"effect": "fly_in",  "cascade": True,  "dur": 500},
    "chart":      {"effect": "fade",    "cascade": False, "dur": 700},
    "table":      {"effect": "fade",    "cascade": False, "dur": 600},
    "gallery":    {"effect": "zoom",    "cascade": True,  "dur": 500},
    "image":      {"effect": "fade",    "cascade": False, "dur": 700},
    "figure":     {"effect": "fade",    "cascade": False, "dur": 700},
    "thanks":     {"effect": "fade",    "cascade": False, "dur": 800},
    "default":    {"effect": "fade",    "cascade": True,  "dur": 500},
}


# ---------------------------------------------------------------------------
# XML 构建辅助
# ---------------------------------------------------------------------------
def _e(tag: str, **attrib) -> etree._Element:
    return etree.Element(f"{{{PML}}}{tag}", attrib)


def _sub(parent: etree._Element, tag: str, **attrib) -> etree._Element:
    return etree.SubElement(parent, f"{{{PML}}}{tag}", attrib)


def _make_cBhvr(shape_id: str, anim_id: str, dur: str) -> etree._Element:
    """构建 p:cBhvr（行为描述，含目标 shape）。"""
    cBhvr = _e("cBhvr")
    cTn = _sub(cBhvr, "cTn", id=anim_id, dur=dur, fill="hold")
    _sub(_sub(cTn, "stCondLst"), "cond", delay="0")
    tgtEl = _sub(cBhvr, "tgtEl")
    _sub(tgtEl, "spTgt", spid=shape_id)
    return cBhvr


def _make_set_visibility(shape_id: str, anim_id: str, dur: str = "1") -> etree._Element:
    """构建 p:set（动画前设为可见）。

    Bug fix: 'to' child must use p:strVal (PML namespace), NOT a:strVal (DML).
    """
    set_el = _e("set")
    cBhvr = _sub(set_el, "cBhvr")
    cTn = _sub(cBhvr, "cTn", id=anim_id, dur="1", fill="hold")
    _sub(_sub(cTn, "stCondLst"), "cond", delay="0")
    tgtEl = _sub(cBhvr, "tgtEl")
    _sub(tgtEl, "spTgt", spid=shape_id)
    attrLst = _sub(cBhvr, "attrNameLst")
    _sub(attrLst, "attrName").text = "style.visibility"
    to_el = _sub(set_el, "to")
    # MUST be p:strVal (PML ns), not a:strVal (DML ns) — XSD schema bug fix
    _sub(to_el, "strVal", val="visible")
    return set_el


def _make_animEffect(shape_id: str, anim_id: str, dur: str,
                     filter_val: str) -> etree._Element:
    """构建 p:animEffect（fade/wipe/split 等合法 filter）。"""
    anim = _e("animEffect", transition="in", filter=filter_val)
    anim.append(_make_cBhvr(shape_id, anim_id, dur))
    return anim


def _make_animMotion(shape_id: str, anim_id: str, dur: str,
                     path: str) -> etree._Element:
    """构建 p:animMotion（fly_in/float_in 位移动画）。

    Bug fix: pathEditMode must be 'relative' or 'fixed' — 'autoResize' is invalid.
    """
    # pathEditMode: ST_TLAnimateMotionPathEditMode only allows 'relative'|'fixed'
    anim = _e("animMotion", origin="layout", path=path, pathEditMode="relative")
    cBhvr = _sub(anim, "cBhvr")
    cTn = _sub(cBhvr, "cTn", id=anim_id, dur=dur, fill="hold")
    _sub(_sub(cTn, "stCondLst"), "cond", delay="0")
    tgtEl = _sub(cBhvr, "tgtEl")
    _sub(tgtEl, "spTgt", spid=shape_id)
    attrLst = _sub(cBhvr, "attrNameLst")
    _sub(attrLst, "attrName").text = "ppt_x"
    _sub(attrLst, "attrName").text = "ppt_y"
    return anim


def _make_animScale(shape_id: str, anim_id: str, dur: str) -> etree._Element:
    """构建 p:animScale（zoom 缩放动画，从 0% 放大到 100%）。"""
    anim = _e("animScale")
    cBhvr = _sub(anim, "cBhvr")
    cTn = _sub(cBhvr, "cTn", id=anim_id, dur=dur, fill="hold")
    _sub(_sub(cTn, "stCondLst"), "cond", delay="0")
    tgtEl = _sub(cBhvr, "tgtEl")
    _sub(tgtEl, "spTgt", spid=shape_id)
    # 从 0,0 缩放到 100000,100000（单位是 1/1000 百分比）
    _sub(anim, "by", x="100000", y="100000")
    _sub(anim, "from", x="0", y="0")
    _sub(anim, "to", x="100000", y="100000")
    return anim


# ---------------------------------------------------------------------------
# 核心：构建完整 <p:timing> 节点
# ---------------------------------------------------------------------------
def _build_timing(shapes: list, effect: str, cascade: bool,
                  dur_ms: int) -> etree._Element:
    """生成包含所有 shape 进入动画的 <p:timing> 完整 XML 树。"""
    eff = _EFFECTS.get(effect, _EFFECTS["fade"])
    dur_s = str(dur_ms)

    # 全局 ID 计数器（每个 cTn id 必须全局唯一且正整数）
    _id = [1]

    def nid() -> str:
        _id[0] += 1
        return str(_id[0])

    # ── 根节点 ──────────────────────────────────────────────────────────────
    timing = _e("timing")
    tnLst  = _sub(timing, "tnLst")
    par_root = _sub(tnLst, "par")
    cTn_root = _sub(par_root, "cTn",
                    id="1", dur="indefinite", restart="whenNotActive",
                    nodeType="tmRoot")
    _sub(_sub(cTn_root, "stCondLst"), "cond",
         evt="onBegin", delay="indefinite")
    childTnLst_root = _sub(cTn_root, "childTnLst")

    # ── 主顺序容器 seq ───────────────────────────────────────────────────────
    seq = _sub(childTnLst_root, "seq", concurrent="1", nextAc="seek")
    seq_id = nid()
    cTn_seq = _sub(seq, "cTn",
                   id=seq_id, dur="indefinite", nodeType="mainSeq")
    _sub(_sub(cTn_seq, "stCondLst"), "cond",
         evt="onBegin", delay="indefinite")
    # Bug fix: 'endSync'/'evtFilter' are not in ST_TLTriggerEvent enum
    # Use 'onEnd' on the parent par instead, which is what PowerPoint generates
    childTnLst_seq = _sub(cTn_seq, "childTnLst")

    # ── 每个形状的动画 ───────────────────────────────────────────────────────
    for i, shape in enumerate(shapes):
        sp_id  = str(shape.shape_id)
        delay  = str(i * _DUR_DELAY) if (cascade and i > 0) else "0"
        is_first = (i == 0)

        # par 容器（每个形状一个）
        par_i    = _sub(childTnLst_seq, "par")
        par_i_id = nid()
        cTn_i    = _sub(par_i, "cTn", id=par_i_id, fill="hold")
        stCond   = _sub(cTn_i, "stCondLst")
        if is_first:
            # 第一个形状由点击触发
            cond = _sub(stCond, "cond", evt="onClick", delay="0")
            _sub(cond, "tn", val=seq_id)
        else:
            _sub(stCond, "cond", delay=delay)
        childTnLst_i = _sub(cTn_i, "childTnLst")

        # par 效果容器
        par_eff    = _sub(childTnLst_i, "par")
        par_eff_id = nid()
        cTn_eff    = _sub(par_eff, "cTn",
                          id=par_eff_id,
                          presetID=eff["presetID"],
                          presetClass=eff["presetClass"],
                          presetSubtype=eff["presetSubtype"],
                          fill="hold", grpId=str(i),
                          nodeType="clickEffect")
        _sub(_sub(cTn_eff, "stCondLst"), "cond", delay="0")
        childTnLst_eff = _sub(cTn_eff, "childTnLst")

        # 1. set visibility (make visible before animation starts)
        childTnLst_eff.append(_make_set_visibility(sp_id, nid(), dur_s))

        # 2. 实际动画节点（按效果类型选择正确的 XML 节点）
        anim_id = nid()
        anim_type = eff["type"]

        if anim_type == "animEffect":
            childTnLst_eff.append(
                _make_animEffect(sp_id, anim_id, dur_s, eff["filter"]))

        elif anim_type == "animMotion":
            childTnLst_eff.append(
                _make_animMotion(sp_id, anim_id, dur_s, eff["path"]))

        elif anim_type == "animScale":
            childTnLst_eff.append(
                _make_animScale(sp_id, anim_id, dur_s))

        elif anim_type == "set_only":
            # appear: 只有 set visibility，无过渡动画
            pass

    # ── prevCondLst（点击回退）
    # Bug fix: 'onPrevClick' not in ST_TLTriggerEvent — use 'onPrev' (valid alias)
    prevCond = _sub(seq, "prevCondLst")
    _sub(prevCond, "cond", evt="onPrev", delay="0")

    return timing


# ---------------------------------------------------------------------------
# 公开接口
# ---------------------------------------------------------------------------
def apply_slide_animation(slide, kind: str = "default") -> None:
    """给一页幻灯片按页型自动选择效果并注入动画。

    使用 OOXML 合法的动画节点类型：
      - fade/wipe/split → p:animEffect（合法 filter 值）
      - fly_in/float_in → p:animMotion（路径位移）
      - zoom            → p:animScale（缩放）
      - appear          → p:set only

    重复调用安全：先移除旧 timing，再注入新的。
    """
    cfg     = KIND_ANIMATION.get(kind, KIND_ANIMATION["default"])
    effect  = cfg["effect"]
    cascade = cfg["cascade"]
    dur_ms  = cfg.get("dur", _DUR_ENTER)

    all_shapes = list(slide.shapes)
    if not all_shapes:
        return

    # 跳过第一个背景图（shape_type==13 且是第一个形状）
    start_idx = 1 if (all_shapes[0].shape_type == 13) else 0
    shapes = all_shapes[start_idx:]
    if not shapes:
        return

    # 移除已有 timing
    sp_tree = slide._element
    for old in sp_tree.findall(f"{{{PML}}}timing"):
        sp_tree.remove(old)

    timing_xml = _build_timing(shapes, effect=effect,
                               cascade=cascade, dur_ms=dur_ms)
    sp_tree.append(timing_xml)
