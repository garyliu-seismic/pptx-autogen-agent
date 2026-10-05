"""全局配置：默认路径与主题预设。"""
from __future__ import annotations

from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# 默认路径
# ---------------------------------------------------------------------------
DEFAULT_INPUT_DIR = "input"                # 素材目录（txt / csv / pdf / png）
DEFAULT_OUTPUT_PATH = "output/deck.pptx"   # 输出 pptx 路径
DEFAULT_PLAN_PATH = "output/plan.json"     # --plan-only 输出的计划 JSON 路径


# ---------------------------------------------------------------------------
# 主题预设：完整配色体系 + 字体方案 + 尺寸体系
# ---------------------------------------------------------------------------
@dataclass
class Theme:
    name: str        # 主题名（中文显示）
    primary: str     # 主色（深色，标题/强调/色块）
    secondary: str   # 辅色（中色，副标题/线条/次要元素）
    accent: str      # 强调色（高亮/图标/箭头/数字）
    bg: str          # 内容页背景色
    fg: str          # 内容页正文文字色

    # ── 扩展色板 ──────────────────────────────────────────────────────────
    bg_card: str = ""        # 卡片背景（空=自动由 bg+primary 派生浅色）
    bg_dark: str = ""        # 深色块背景（封面/章节，空=使用 primary）
    neutral: str = "6B7280"  # 中性灰（辅助文字/线条/页码）
    success: str = "10B981"  # 正向指标/✔ 颜色
    warning: str = "F59E0B"  # 警示/负向指标颜色

    # ── 字体方案 ──────────────────────────────────────────────────────────
    font_major: str = "Microsoft YaHei"  # 标题字体（大标题/封面）
    font_minor: str = "Microsoft YaHei"  # 正文字体（要点/说明）
    font_mono: str = "Consolas"          # 等宽字体（代码/纯数字指标）

    # ── 字号体系（pt）──────────────────────────────────────────────────────
    size_cover_title: int = 44   # 封面主标题
    size_cover_sub: int = 20     # 封面副标题
    size_section: int = 40       # 章节分隔页标题
    size_slide_title: int = 30   # 内容页标题栏
    size_body: int = 18          # 正文要点
    size_body_sm: int = 14       # 小号正文（双栏/多要点时）
    size_caption: int = 13       # 图注/说明
    size_footer: int = 11        # 页脚/来源
    size_kpi_value: int = 42     # KPI 大数字
    size_kpi_label: int = 14     # KPI 指标名

    # ── 视觉效果 ──────────────────────────────────────────────────────────
    card_radius: float = 0.07    # 卡片圆角（adjustments[0]，0=直角，0.5=半圆）
    use_shadow: bool = False     # 是否对卡片/图表框启用阴影
    line_height: float = 1.25    # 全局行高倍率

    # ── 背景图 ────────────────────────────────────────────────────────────
    bg_image: str = ""           # 封面/章节背景图（全出血铺底）
    bg_image_soft: str = ""      # 内容页背景图（较淡，保证文字可读）

    # ── 设计风格 ─────────────────────────────────────────────────────
    # style="editorial"时启用参考文件的精细排版风格：
    #   - 内容页米白背景+5层字体级别+红色标签+细线分割
    #   - 分隔页用真实照片（bg_image）全屏铺底+左大字+右标题
    #   - 按矩幂布局，无圆角卡片
    design_style: str = "default"   # "default" | "editorial"


# ---------------------------------------------------------------------------
# 预设主题（9 个内置 + 后续可动态扩展）
# ---------------------------------------------------------------------------
THEMES: dict[str, Theme] = {

    # ── 科技蓝：深海蓝底，电光蓝高亮，适合 AI/产品/技术发布 ────────────────
    "tech": Theme(
        name="科技蓝",
        primary="1F4E79", secondary="00B0F0", accent="FFC000",
        bg="F0F7FF", fg="0D2137",
        bg_dark="0D2137", neutral="4B7BA6",
        success="00C49A", warning="FF8C42",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        size_kpi_value=48,
        bg_image="assets/backgrounds/tech_bg.png",
        bg_image_soft="assets/backgrounds/tech_bg_soft.png",
        card_radius=0.06,
    ),

    # ── 商务金：深蓝主色，金色强调，适合季报/战略/商业提案 ─────────────────
    "business": Theme(
        name="商务金",
        primary="1F3864", secondary="C9A84C", accent="C00000",
        bg="FAFAFA", fg="1A1A1A",
        bg_dark="0F1F3C", neutral="8D9CAB",
        success="2E7D5E", warning="D4820A",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        size_cover_title=42, size_section=38,
        bg_image="assets/backgrounds/business_bg.png",
        bg_image_soft="assets/backgrounds/business_bg_soft.png",
        card_radius=0.04,
    ),

    # ── 金融墨：深海军蓝，蓝绿高亮，金色指标，适合投行/量化/财务报告 ────────
    "finance": Theme(
        name="金融墨",
        primary="0B3C5D", secondary="328CC1", accent="D9B310",
        bg="F8FAFB", fg="1C2833",
        bg_dark="061A2B", neutral="5D7E96",
        success="1A9B6A", warning="E07B00",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        font_mono="Consolas",
        size_kpi_value=46, size_kpi_label=13,
        bg_image="assets/backgrounds/finance_bg.png",
        bg_image_soft="assets/backgrounds/finance_bg_soft.png",
        card_radius=0.03, use_shadow=False,
    ),

    # ── 健康绿：清新绿主色，橙色强调，适合医疗/健康/ESG ────────────────────
    "health": Theme(
        name="健康绿",
        primary="0F6C4F", secondary="2E9E6B", accent="FF8A5C",
        bg="F2FAF6", fg="1B3A2E",
        bg_dark="083D2C", neutral="5A8A70",
        success="0F6C4F", warning="E06000",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/health_bg.png",
        bg_image_soft="assets/backgrounds/health_bg_soft.png",
        card_radius=0.08, line_height=1.3,
    ),

    # ── 教育橙：暖橙主色，紫色强调，适合培训/课程/学术 ─────────────────────
    "education": Theme(
        name="教育橙",
        primary="92400E", secondary="D97706", accent="7C3AED",
        bg="FFFDF5", fg="292524",
        bg_dark="5C2800", neutral="A07040",
        success="059669", warning="D97706",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/education_bg.png",
        bg_image_soft="assets/backgrounds/education_bg_soft.png",
        card_radius=0.09, line_height=1.3,
    ),

    # ── 自然青：青色主色，翠绿高亮，适合环保/能源/可持续 ───────────────────
    "nature": Theme(
        name="自然青",
        primary="0E7490", secondary="06B6D4", accent="84CC16",
        bg="F0FDFA", fg="134E52",
        bg_dark="073747", neutral="4D8B99",
        success="65A30D", warning="D97706",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/nature_bg.png",
        bg_image_soft="assets/backgrounds/nature_bg_soft.png",
        card_radius=0.08,
    ),

    # ── 简约灰：低饱和主色，适合工作周报/纪要/中性汇报 ─────────────────────
    "elegant": Theme(
        name="简约灰",
        primary="1F2937", secondary="6B7280", accent="3B82F6",
        bg="FFFFFF", fg="111827",
        bg_dark="111827", neutral="9CA3AF",
        success="10B981", warning="F59E0B",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/elegant_bg.png",
        bg_image_soft="assets/backgrounds/elegant_bg_soft.png",
        card_radius=0.04, line_height=1.2,
    ),

    # ── 现代靛：靛蓝主色，琥珀强调，现代感强，适合产品/品牌/互联网 ─────────
    "modern": Theme(
        name="现代靛",
        primary="312E81", secondary="6366F1", accent="F59E0B",
        bg="FAFAFE", fg="1E1B4B",
        bg_dark="1E1B4B", neutral="818CF8",
        success="10B981", warning="EF4444",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        size_kpi_value=48,
        bg_image="assets/backgrounds/modern_bg.png",
        bg_image_soft="assets/backgrounds/modern_bg_soft.png",
        card_radius=0.07,
    ),

    # ── AI 科技：深宇宙底，青紫霓虹，适合 AI/大模型/深科技 ─────────────────
    "ai": Theme(
        name="AI 科技",
        primary="6366F1", secondary="22D3EE", accent="A855F7",
        bg="0B1020", fg="E2E8F0",
        bg_dark="060D1A", neutral="94A3B8",
        success="34D399", warning="FB923C",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        size_kpi_value=50,
        card_radius=0.06, use_shadow=True,
        bg_image="assets/ai_bg.png",
        bg_image_soft="assets/ai_bg_soft.png",
    ),

    # ── 深色专业：近黑底，冰蓝强调，高端科技感 ─────────────────────────────
    "dark_pro": Theme(
        name="深色专业",
        primary="0F172A", secondary="1E293B", accent="38BDF8",
        bg="0F172A", fg="F1F5F9",
        bg_dark="020617", neutral="475569",
        success="34D399", warning="FB923C",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        size_kpi_value=50,
        bg_image="assets/backgrounds/dark_pro_bg.png",
        bg_image_soft="assets/backgrounds/dark_pro_bg_soft.png",
        card_radius=0.06, use_shadow=True,
    ),

    # ── 企业蓝灰：传统企业风，适合正式报告/政府/大型企业 ───────────────────
    "corporate": Theme(
        name="企业蓝灰",
        primary="1E3A5F", secondary="2563EB", accent="60A5FA",
        bg="F8FAFC", fg="0F172A",
        bg_dark="0C1D35", neutral="64748B",
        success="16A34A", warning="D97706",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/corporate_bg.png",
        bg_image_soft="assets/backgrounds/corporate_bg_soft.png",
        card_radius=0.03, line_height=1.2,
    ),

    # ── 极简白：超低调纯白，配色仅靠 accent 线条/数字点亮 ──────────────────
    "minimal": Theme(
        name="极简白",
        primary="111827", secondary="374151", accent="6366F1",
        bg="FFFFFF", fg="111827",
        bg_dark="111827", neutral="D1D5DB",
        success="059669", warning="D97706",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/minimal_bg.png",
        bg_image_soft="assets/backgrounds/minimal_bg_soft.png",
        card_radius=0.03, line_height=1.25,
        size_cover_title=48, size_slide_title=28,
    ),

    # ── 温暖橙红：暖色系，热情活力，适合消费品/零售/活动策划 ────────────────
    "warm": Theme(
        name="温暖橙红",
        primary="7C2D12", secondary="EA580C", accent="FCD34D",
        bg="FFFBF0", fg="1C1917",
        bg_dark="431407", neutral="A86C4A",
        success="15803D", warning="EA580C",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/warm_bg.png",
        bg_image_soft="assets/backgrounds/warm_bg_soft.png",
        card_radius=0.08,
    ),

    # ── 医疗白蓝：白底+蓝色，医疗/健康/生命科学报告 ─────────────────────
    "medical": Theme(
        name="医疗白蓝",
        primary="005B8E", secondary="0099D6", accent="00C5DC",
        bg="FFFFFF", fg="1A2B3C",
        bg_dark="003A5C", neutral="8BAAB8",
        success="00A878", warning="FF7A45",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        card_radius=0.06, line_height=1.3,
    ),

    # ── 黑金奢华：黑底金字，高端品牌/奖项宣她/年度盛典 ─────────────────
    "luxury": Theme(
        name="黑金奢华",
        primary="B8860B", secondary="DAA520", accent="FFD700",
        bg="0A0A0A", fg="F5F0E8",
        bg_dark="000000", neutral="6B6B6B",
        success="7CFC00", warning="FF8C00",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        size_kpi_value=52, bg_image="assets/backgrounds/luxury_bg.png",
        bg_image_soft="assets/backgrounds/luxury_bg_soft.png",
        card_radius=0.04, use_shadow=True,
    ),

    # ── 清新薄荷：年轻/互联网/TO-C 产品/创业展示 ──────────────────
    "fresh": Theme(
        name="清新薄荷",
        primary="00BFA5", secondary="26C6DA", accent="FF4081",
        bg="F0FFFC", fg="1A3330",
        bg_dark="00695C", neutral="80CBC4",
        success="69F0AE", warning="FFAB40",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/fresh_bg.png",
        bg_image_soft="assets/backgrounds/fresh_bg_soft.png",
        card_radius=0.10, line_height=1.35,
    ),

    # ── 学术蓝灰：典雅学术风，适合研究报告/论文展示/白皮书 ─────────────
    "academic": Theme(
        name="学术蓝灰",
        primary="2C3E50", secondary="34495E", accent="2980B9",
        bg="FDFDFD", fg="2C3E50",
        bg_dark="1A252F", neutral="95A5A6",
        success="27AE60", warning="E67E22",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        size_cover_title=40, size_slide_title=26,
        bg_image="assets/backgrounds/academic_bg.png",
        bg_image_soft="assets/backgrounds/academic_bg_soft.png",
        card_radius=0.02, line_height=1.4,
    ),

    # ── 政务红：庄重正式，适合政府/党政/大型国企汇报 ───────────────
    "government": Theme(
        name="政务红",
        primary="C0392B", secondary="E74C3C", accent="F39C12",
        bg="FAFAFA", fg="1A1A1A",
        bg_dark="7B1414", neutral="888888",
        success="27AE60", warning="E67E22",
        font_major="Microsoft YaHei", font_minor="Microsoft YaHei",
        bg_image="assets/backgrounds/government_bg.png",
        bg_image_soft="assets/backgrounds/government_bg_soft.png",
        card_radius=0.02, line_height=1.3,
    ),

    # ── 山海量柯旅行：参考文件原色系，5色极简居家旅行风 ──────────
    "travel": Theme(
        name="山海旅行",
        primary="0B4F4A",      # 深森绱（封面/章节底色、标题）
        secondary="18A999",    # 青绱（标签/编号/小标题）
        accent="F05D5E",       # 珊瑚红（强调/标签枚/数字）
        bg="F7F3EA",           # 米白（内容页背景）
        fg="183438",           # 深色刻度灰（正文）
        bg_dark="0B4F4A",      # 封面/章节页深色底
        bg_card="FFFFFF",      # 内容卡片白色底
        neutral="5D7776",      # 灰绳（副标题/说明文字）
        success="18A999",      # 正向标识复用 secondary
        warning="F05D5E",      # 警示复用 accent
        font_major="Microsoft YaHei",
        font_minor="Microsoft YaHei",
        # 字号与参考文件进行匹配
        size_cover_title=44,
        size_cover_sub=24,
        size_section=40,
        size_slide_title=26,   # 内容页大标题 26pt
        size_body=17,
        size_body_sm=15,
        size_caption=13,
        size_footer=8,         # 页码 8pt极小
        size_kpi_value=36,
        size_kpi_label=17,
        card_radius=0.0,       # 直角，编辑风无图角
        line_height=1.3,
        design_style="editorial",
    ),
}

DEFAULT_THEME = "modern"
