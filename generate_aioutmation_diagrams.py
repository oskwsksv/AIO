"""Generate 5 explanatory diagrams for AIOutmation in PowerPoint format.
1. Use Case Diagram (ユースケース図)
2. User Story Map (ユーザーストーリーマップ)
3. MVP System Architecture (MVPシステム構成図)
4. Input → Process → Output Diagram
5. AIO Pyramid × Feature Mapping
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn


# Color palette
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
BLUE = RGBColor(0x2E, 0x5C, 0x8A)
LIGHT_BLUE = RGBColor(0xE8, 0xEE, 0xF7)
SKY_BLUE = RGBColor(0xB8, 0xD4, 0xE8)
ORANGE = RGBColor(0xC9, 0x6E, 0x29)
LIGHT_ORANGE = RGBColor(0xFF, 0xF4, 0xE0)
RED = RGBColor(0xB0, 0x3A, 0x2E)
LIGHT_RED = RGBColor(0xFF, 0xE0, 0xE0)
GREEN = RGBColor(0x2E, 0x7A, 0x4E)
LIGHT_GREEN = RGBColor(0xE0, 0xF4, 0xE0)
YELLOW = RGBColor(0xFF, 0xC8, 0x4D)
LIGHT_YELLOW = RGBColor(0xFF, 0xF4, 0xC0)
GRAY = RGBColor(0x80, 0x80, 0x80)
LIGHT_GRAY = RGBColor(0xF0, 0xF0, 0xF0)
DARK_GRAY = RGBColor(0x44, 0x44, 0x44)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x22, 0x22, 0x22)
PURPLE = RGBColor(0x6B, 0x46, 0xA0)
LIGHT_PURPLE = RGBColor(0xEC, 0xE0, 0xF4)


def set_jp_font(run, size=10, bold=False, color=BLACK, font_name="Yu Gothic"):
    run.font.size = Pt(size)
    run.font.name = font_name
    run.font.bold = bold
    run.font.color.rgb = color
    # Set East Asian font
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("a:rFonts"))
    if rFonts is None:
        from lxml import etree
        rFonts = etree.SubElement(rPr, qn("a:rFonts"))
    rFonts.set("eastAsia", font_name)


def set_text(tf, text, size=10, bold=False, color=BLACK, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE):
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = va
    p = tf.paragraphs[0]
    p.alignment = align
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            r = p.add_run()
        else:
            np = tf.add_paragraph()
            np.alignment = align
            np.space_after = Pt(0)
            np.space_before = Pt(0)
            r = np.add_run()
        r.text = line
        set_jp_font(r, size=size, bold=bold, color=color)


def add_rect(slide, left, top, width, height, fill, line_color=None, line_w=1):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_color is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line_color
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s


def add_rounded_rect(slide, left, top, width, height, fill, line_color=None, line_w=1):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_color is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line_color
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s


def add_oval(slide, left, top, width, height, fill, line_color=NAVY, line_w=1):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = line_color
    s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s


def add_text_in_shape(shape, text, size=10, bold=False, color=BLACK, align=PP_ALIGN.CENTER):
    set_text(shape.text_frame, text, size=size, bold=bold, color=color, align=align)


def add_textbox(slide, left, top, width, height, text, size=10, bold=False, color=BLACK,
                align=PP_ALIGN.LEFT, fill=None, va=MSO_ANCHOR.TOP):
    if fill is not None:
        add_rect(slide, left, top, width, height, fill, line_color=GRAY)
    tb = slide.shapes.add_textbox(left, top, width, height)
    set_text(tb.text_frame, text, size=size, bold=bold, color=color, align=align, va=va)
    return tb


def add_actor_stick(slide, left, top, size_factor=1.0, label="", color=NAVY):
    """Add a UML-style stick figure actor with label."""
    base_w = Inches(0.6 * size_factor)
    base_h = Inches(1.3 * size_factor)
    # Head (circle)
    head_size = Inches(0.3 * size_factor)
    head = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + (base_w - head_size)/2, top, head_size, head_size)
    head.fill.solid()
    head.fill.fore_color.rgb = WHITE
    head.line.color.rgb = color
    head.line.width = Pt(2)
    head.shadow.inherit = False
    # Body (vertical line)
    body_top = top + head_size
    body_bottom = top + Inches(0.85 * size_factor)
    body = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                       left + base_w/2, body_top,
                                       left + base_w/2, body_bottom)
    body.line.color.rgb = color
    body.line.width = Pt(2)
    # Arms (horizontal line)
    arms_y = top + Inches(0.5 * size_factor)
    arms = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                       left, arms_y, left + base_w, arms_y)
    arms.line.color.rgb = color
    arms.line.width = Pt(2)
    # Legs (V shape)
    left_leg = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                           left + base_w/2, body_bottom,
                                           left, body_bottom + Inches(0.4 * size_factor))
    left_leg.line.color.rgb = color
    left_leg.line.width = Pt(2)
    right_leg = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                            left + base_w/2, body_bottom,
                                            left + base_w, body_bottom + Inches(0.4 * size_factor))
    right_leg.line.color.rgb = color
    right_leg.line.width = Pt(2)
    # Label
    label_top = top + Inches(1.3 * size_factor)
    label_box = slide.shapes.add_textbox(left - Inches(0.4), label_top,
                                          base_w + Inches(0.8), Inches(0.4))
    set_text(label_box.text_frame, label, size=10, bold=True, color=color, align=PP_ALIGN.CENTER)


def add_arrow(slide, x1, y1, x2, y2, color=DARK_GRAY, width=1.5, dashed=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = Pt(width)
    # Add arrowhead
    c.line._get_or_add_ln()
    from lxml import etree
    ln = c.line._get_or_add_ln()
    # Add tailEnd or headEnd
    nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    head_end = etree.SubElement(ln, qn("a:tailEnd"))
    head_end.set("type", "triangle")
    head_end.set("w", "med")
    head_end.set("h", "med")
    if dashed:
        prst_dash = etree.SubElement(ln, qn("a:prstDash"))
        prst_dash.set("val", "dash")
    return c


def add_line(slide, x1, y1, x2, y2, color=DARK_GRAY, width=1.0):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


def title_bar(slide, title, subtitle=None, color=NAVY):
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.7), color)
    add_textbox(slide, Inches(0.3), Inches(0.08), Inches(13), Inches(0.4),
                title, size=20, bold=True, color=WHITE)
    if subtitle:
        add_textbox(slide, Inches(0.3), Inches(0.42), Inches(13), Inches(0.25),
                    subtitle, size=10, color=WHITE)


def add_footer(slide, page_num, total=5):
    add_textbox(slide, Inches(0.3), Inches(7.15), Inches(8), Inches(0.3),
                f"AIOutmation 解説図シリーズ ({page_num}/{total})  ─  作成: 2026年5月",
                size=9, color=GRAY, align=PP_ALIGN.LEFT)


# ============== Build Presentation ==============
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


# ============================================================
# Slide 1: Cover
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, prs.slide_height, NAVY)

add_textbox(slide, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.8),
            "AIOutmation を5つの図で理解する", size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide, Inches(0.5), Inches(2.1), Inches(12.3), Inches(0.5),
            "ユーザー視点・開発者視点・システム視点の図解集", size=16, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

# 5 thumbnails as a row
thumb_w = Inches(2.3)
thumb_h = Inches(2.5)
thumb_y = Inches(3.0)
spacing = Inches(0.15)
total_thumb_w = thumb_w * 5 + spacing * 4
thumb_start_x = (prs.slide_width - total_thumb_w) / 2

thumbnails = [
    ("①", "ユースケース図", "ユーザー視点\n誰が何をできるか", BLUE),
    ("②", "ユーザーストーリーマップ", "ユーザー視点\n体験の時間軸", ORANGE),
    ("③", "MVPシステム構成図", "開発者視点\nコンポーネント構成", PURPLE),
    ("④", "I/Oフロー図", "システム視点\n入力→処理→出力", GREEN),
    ("⑤", "AIOピラミッド×機能対応", "概念視点\n製品の位置づけ", RED),
]

for i, (num, title, desc, c) in enumerate(thumbnails):
    x = thumb_start_x + (thumb_w + spacing) * i
    box = add_rounded_rect(slide, x, thumb_y, thumb_w, thumb_h, WHITE, line_color=c, line_w=2)
    # Number
    add_textbox(slide, x, thumb_y + Inches(0.15), thumb_w, Inches(0.6),
                num, size=42, bold=True, color=c, align=PP_ALIGN.CENTER)
    add_textbox(slide, x, thumb_y + Inches(0.95), thumb_w, Inches(0.5),
                title, size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_textbox(slide, x, thumb_y + Inches(1.5), thumb_w, Inches(0.9),
                desc, size=10, color=DARK_GRAY, align=PP_ALIGN.CENTER)

add_textbox(slide, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.4),
            "出典: AIOutmation_LeanCanvas.docx / AIOutmation_MVP_Spec.docx / AIOutmation_Story.docx",
            size=10, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


# ============================================================
# Slide 2: Use Case Diagram
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_bar(slide, "①ユースケース図 (Use Case Diagram)", "ユーザー視点: 誰が何をできるシステムか")

# System boundary (rectangle in center)
sys_left = Inches(3.5)
sys_top = Inches(1.1)
sys_w = Inches(6.3)
sys_h = Inches(5.9)
sys_boundary = add_rect(slide, sys_left, sys_top, sys_w, sys_h, WHITE,
                         line_color=NAVY, line_w=2)
# System name label
add_textbox(slide, sys_left, sys_top + Inches(0.05), sys_w, Inches(0.35),
            "AIOutmation System", size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

# Use cases (ellipses inside system)
use_cases = [
    # (label, x_offset, y_offset, width, height)
    ("プロンプト管理", 0.3, 0.6, 1.5, 0.7),
    ("プロンプト\n自動生成(AI)", 2.2, 0.6, 1.6, 0.85),
    ("引用率\n自動測定", 4.3, 0.6, 1.5, 0.85),
    ("ダッシュボード\n閲覧", 0.3, 1.7, 1.8, 0.85),
    ("アラート設定\n/受信", 2.5, 1.7, 1.7, 0.85),
    ("月次レポート\n生成", 4.5, 1.7, 1.5, 0.85),
    ("AIコメンタリー\n生成", 0.5, 2.9, 1.7, 0.85),
    ("ホワイトラベル\n設定", 2.5, 2.9, 1.7, 0.85),
    ("競合分析", 4.7, 2.9, 1.3, 0.7),
    ("マルチテナント\n管理", 0.5, 4.1, 1.7, 0.85),
    ("CMS改修PR\n起票(V2)", 2.5, 4.1, 1.7, 0.85),
    ("LLM別最適化\n(V2)", 4.7, 4.1, 1.4, 0.85),
]

uc_positions = {}  # label_key -> (cx, cy) center coordinates
for label, dx, dy, w, h in use_cases:
    x = sys_left + Inches(dx)
    y = sys_top + Inches(0.5 + dy)
    ww = Inches(w)
    hh = Inches(h)
    is_v2 = "(V2)" in label
    fill = LIGHT_PURPLE if is_v2 else LIGHT_BLUE
    line_c = PURPLE if is_v2 else BLUE
    add_oval(slide, x, y, ww, hh, fill, line_color=line_c, line_w=1.5)
    # Add text manually for ellipses
    add_textbox(slide, x, y, ww, hh, label, size=9, bold=True, color=NAVY,
                align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
    uc_positions[label] = (x + ww/2, y + hh/2)

# Primary actors (left side)
actor_x = Inches(0.5)
add_actor_stick(slide, Inches(0.5), Inches(1.7), 0.9, "上場企業\nマーケPM", color=NAVY)
add_actor_stick(slide, Inches(0.5), Inches(4.0), 0.9, "代理店PM", color=ORANGE)

# Secondary actors (right side)
actor_x2 = Inches(11.2)
add_actor_stick(slide, Inches(11.0), Inches(1.7), 0.8, "LLM\nAPIs", color=PURPLE)
add_actor_stick(slide, Inches(11.0), Inches(3.5), 0.8, "CRM\n(Salesforce)", color=GREEN)
add_actor_stick(slide, Inches(11.0), Inches(5.3), 0.8, "CMS\n(各種)", color=RED)

# Lines from actors to use cases (simplified)
# 上場企業マーケPM ↔ several use cases
pm_x = Inches(1.1)
pm_y = Inches(2.0)
for label in ["プロンプト管理", "ダッシュボード\n閲覧", "アラート設定\n/受信", "AIコメンタリー\n生成"]:
    cx, cy = uc_positions[label]
    add_line(slide, pm_x, pm_y, cx, cy, color=NAVY, width=0.75)

# 代理店PM ↔ several use cases
agpm_x = Inches(1.1)
agpm_y = Inches(4.3)
for label in ["ホワイトラベル\n設定", "月次レポート\n生成", "マルチテナント\n管理", "AIコメンタリー\n生成"]:
    cx, cy = uc_positions[label]
    add_line(slide, agpm_x, agpm_y, cx, cy, color=ORANGE, width=0.75)

# LLM APIs ↔ 引用率測定, プロンプト自動生成
llm_x = Inches(11.0)
llm_y = Inches(2.0)
for label in ["引用率\n自動測定", "プロンプト\n自動生成(AI)", "AIコメンタリー\n生成", "LLM別最適化\n(V2)"]:
    cx, cy = uc_positions[label]
    add_line(slide, llm_x, llm_y, cx, cy, color=PURPLE, width=0.75)

# CRM ↔ プロンプト自動生成
crm_x = Inches(11.0)
crm_y = Inches(3.8)
add_line(slide, crm_x, crm_y, uc_positions["プロンプト\n自動生成(AI)"][0], uc_positions["プロンプト\n自動生成(AI)"][1],
         color=GREEN, width=0.75)

# CMS ↔ CMS PR起票
cms_x = Inches(11.0)
cms_y = Inches(5.6)
add_line(slide, cms_x, cms_y, uc_positions["CMS改修PR\n起票(V2)"][0], uc_positions["CMS改修PR\n起票(V2)"][1],
         color=RED, width=0.75)

# Legend
legend_y = Inches(6.5)
legend_x = Inches(0.5)
add_textbox(slide, legend_x, legend_y, Inches(2), Inches(0.3),
            "凡例:", size=10, bold=True, color=DARK_GRAY)
add_oval(slide, Inches(2.0), legend_y + Inches(0.05), Inches(0.3), Inches(0.2), LIGHT_BLUE, line_color=BLUE)
add_textbox(slide, Inches(2.3), legend_y, Inches(1.4), Inches(0.3), "MVP機能", size=9, color=DARK_GRAY)
add_oval(slide, Inches(3.8), legend_y + Inches(0.05), Inches(0.3), Inches(0.2), LIGHT_PURPLE, line_color=PURPLE)
add_textbox(slide, Inches(4.1), legend_y, Inches(1.4), Inches(0.3), "V2拡張機能", size=9, color=DARK_GRAY)

add_footer(slide, 1)


# ============================================================
# Slide 3: User Story Map
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_bar(slide, "②ユーザーストーリーマップ", "ユーザー視点: 体験の時間軸と機能の対応関係", color=ORANGE)

# Backbone activities (top row) - 6 activities
activities = ["1. オンボード", "2. 設定", "3. 日常運用", "4. 分析・改善", "5. 経営報告", "6. 拡張"]
backbone_y = Inches(0.95)
backbone_h = Inches(0.55)
col_w = (prs.slide_width - Inches(1.0)) / 6
start_x = Inches(0.5)

for i, act in enumerate(activities):
    x = start_x + col_w * i
    add_rect(slide, x + Inches(0.05), backbone_y, col_w - Inches(0.1), backbone_h, ORANGE)
    add_textbox(slide, x + Inches(0.05), backbone_y, col_w - Inches(0.1), backbone_h,
                act, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# User tasks (2nd row)
user_tasks_y = Inches(1.6)
user_tasks_h = Inches(0.8)
user_tasks = [
    "・アカウント作成\n・チーム招待\n・ブランド設定",
    "・競合登録\n・プロンプト生成\n・モデル選定",
    "・引用率モニター\n・アラート対応\n・競合急変キャッチ",
    "・トレンド分析\n・改修案レビュー\n・ライター/開発依頼",
    "・月次レポート確認\n・経営会議提出\n・ROI証明",
    "・新ブランド追加\n・多言語展開\n・代理店招待",
]
for i, tasks in enumerate(user_tasks):
    x = start_x + col_w * i
    add_rect(slide, x + Inches(0.05), user_tasks_y, col_w - Inches(0.1), user_tasks_h, LIGHT_ORANGE,
             line_color=ORANGE)
    add_textbox(slide, x + Inches(0.05), user_tasks_y, col_w - Inches(0.1), user_tasks_h,
                tasks, size=9, color=DARK_GRAY, align=PP_ALIGN.LEFT, va=MSO_ANCHOR.MIDDLE)

# Release rows: MVP, V2, V3
release_rows = [
    ("MVP", LIGHT_BLUE, BLUE, [
        "F1: マルチテナント基盤\nF1: SSO認証",
        "F2: プロンプト自動生成\nF2: タグ・優先度",
        "F3: 引用率測定エンジン\nF4: アラート(Slack/メール)",
        "F4: ダッシュボード\nF4: ヒートマップ",
        "F5: ホワイトラベルレポート\nF6: AIコメンタリー(3行)",
        "F1: Workspace複製\nF5: 代理店向け招待",
    ]),
    ("V2", LIGHT_PURPLE, PURPLE, [
        "F12: Salesforce連携\nF12: HubSpot連携",
        "F12: Gong連携(プロンプト抽出)",
        "F8: JSON-LD自動生成\nF9: CMS PR起票",
        "F7: SEO×AIO監査\n統計的ノイズ除去",
        "F10: ROI物語AI\n業界ベンチマーク",
        "Felo/Genspark対応\n多言語完全対応",
    ]),
    ("V3", LIGHT_GREEN, GREEN, [
        "F17: SOC2 Type II",
        "F13: 文章構成最適化AI(リライト)",
        "F14: Signal/Noise Engine\n統計的信頼区間",
        "F11: 業界ベンチマークDB公開",
        "AIによる経営Q&A対応\n取締役会用デック自動",
        "API公開\nプラグイン市場",
    ]),
]

release_y = Inches(2.5)
release_h = Inches(1.4)

for ri, (label, fill, line_c, tasks) in enumerate(release_rows):
    y = release_y + (release_h + Inches(0.1)) * ri
    # Left label column
    add_rect(slide, Inches(0.05), y, Inches(0.4), release_h, line_c)
    add_textbox(slide, Inches(0.05), y, Inches(0.4), release_h,
                label, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

    for i, t in enumerate(tasks):
        x = start_x + col_w * i
        add_rect(slide, x + Inches(0.05), y, col_w - Inches(0.1), release_h, fill,
                 line_color=line_c)
        add_textbox(slide, x + Inches(0.05), y, col_w - Inches(0.1), release_h,
                    t, size=9, color=DARK_GRAY, align=PP_ALIGN.LEFT, va=MSO_ANCHOR.MIDDLE)

add_textbox(slide, Inches(0.5), Inches(6.9), Inches(12.3), Inches(0.3),
            "横軸: ユーザー体験の時間軸(オンボード→拡張)  /  縦軸: 機能のリリース時期(MVP→V3)",
            size=10, color=DARK_GRAY, align=PP_ALIGN.CENTER)

add_footer(slide, 2)


# ============================================================
# Slide 4: MVP System Architecture
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_bar(slide, "③MVPシステム構成図 (Architecture)", "開発者視点: コンポーネントと技術スタック", color=PURPLE)

# Layer structure
# Layer 1: User (top)
# Layer 2: Frontend
# Layer 3: API + Auth
# Layer 4: Core Services + Worker
# Layer 5: Data Layer
# Side: External Services

# User actors at top
user_y = Inches(0.9)
add_actor_stick(slide, Inches(2.5), user_y, 0.55, "上場企業\nマーケPM", color=NAVY)
add_actor_stick(slide, Inches(8.5), user_y, 0.55, "代理店PM", color=ORANGE)

# Frontend layer
fe_y = Inches(2.0)
fe_box = add_rounded_rect(slide, Inches(2.0), fe_y, Inches(9.3), Inches(0.7), LIGHT_BLUE,
                          line_color=BLUE, line_w=1.5)
add_textbox(slide, Inches(2.0), fe_y, Inches(9.3), Inches(0.7),
            "Frontend  ─  Next.js 14 (App Router) + TypeScript + Tailwind + shadcn/ui",
            size=12, bold=True, color=BLUE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# Arrow from users to frontend
add_arrow(slide, Inches(2.8), user_y + Inches(1.1), Inches(4.0), fe_y, color=DARK_GRAY)
add_arrow(slide, Inches(8.8), user_y + Inches(1.1), Inches(9.0), fe_y, color=DARK_GRAY)

# API Gateway layer
api_y = Inches(3.0)
api_box = add_rounded_rect(slide, Inches(2.0), api_y, Inches(7.3), Inches(0.7), LIGHT_PURPLE,
                            line_color=PURPLE, line_w=1.5)
add_textbox(slide, Inches(2.0), api_y, Inches(7.3), Inches(0.7),
            "API Gateway  ─  NestJS (Node.js / TypeScript) + REST/tRPC",
            size=11, bold=True, color=PURPLE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# Auth box (right of API)
auth_box = add_rounded_rect(slide, Inches(9.5), api_y, Inches(1.8), Inches(0.7), LIGHT_GREEN,
                             line_color=GREEN, line_w=1.5)
add_textbox(slide, Inches(9.5), api_y, Inches(1.8), Inches(0.7),
            "Auth\n(Clerk)", size=11, bold=True, color=GREEN, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# Arrow FE → API
add_arrow(slide, Inches(6.7), fe_y + Inches(0.7), Inches(6.7), api_y, color=DARK_GRAY)

# Core Services layer
svc_y = Inches(4.0)
svc_w = Inches(1.85)
svc_h = Inches(1.05)
services = [
    ("Prompt\nService", "F2", BLUE),
    ("Citation\nEngine", "F3 ★", RED),
    ("Dashboard\nAPI", "F4", BLUE),
    ("Report\nGenerator", "F5", BLUE),
    ("AI Commentary\nService", "F6", BLUE),
]
for i, (name, code, c) in enumerate(services):
    x = Inches(0.5) + (svc_w + Inches(0.1)) * i
    box = add_rounded_rect(slide, x, svc_y, svc_w, svc_h, LIGHT_BLUE, line_color=c, line_w=1.5)
    add_textbox(slide, x, svc_y, svc_w, Inches(0.65),
                name, size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, x, svc_y + Inches(0.65), svc_w, Inches(0.35),
                code, size=9, color=c, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# Worker Pool layer
worker_y = Inches(5.25)
worker_box = add_rounded_rect(slide, Inches(0.5), worker_y, Inches(10.2), Inches(0.7), LIGHT_RED,
                               line_color=RED, line_w=1.5)
add_textbox(slide, Inches(0.5), worker_y, Inches(10.2), Inches(0.7),
            "LLM Worker Pool  ─  BullMQ + Redis Queue / Playwright(Containerized) + LLM API Clients",
            size=11, bold=True, color=RED, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# Data layer
data_y = Inches(6.2)
data_box1 = add_rounded_rect(slide, Inches(0.5), data_y, Inches(4.9), Inches(0.8), LIGHT_YELLOW,
                              line_color=YELLOW, line_w=1.5)
add_textbox(slide, Inches(0.5), data_y, Inches(4.9), Inches(0.8),
            "PostgreSQL 16\n(Prisma ORM、Row-Level Security、月次パーティション)",
            size=10, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

data_box2 = add_rounded_rect(slide, Inches(5.6), data_y, Inches(2.5), Inches(0.8), LIGHT_YELLOW,
                              line_color=YELLOW, line_w=1.5)
add_textbox(slide, Inches(5.6), data_y, Inches(2.5), Inches(0.8),
            "Redis\n(Cache + Queue)", size=10, bold=True, color=DARK_GRAY,
            align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

data_box3 = add_rounded_rect(slide, Inches(8.3), data_y, Inches(2.4), Inches(0.8), LIGHT_YELLOW,
                              line_color=YELLOW, line_w=1.5)
add_textbox(slide, Inches(8.3), data_y, Inches(2.4), Inches(0.8),
            "S3\n(レポート/ログ/アーカイブ)", size=10, bold=True, color=DARK_GRAY,
            align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# External Services panel (right side)
ext_x = Inches(11.4)
ext_y = Inches(2.0)
ext_w = Inches(1.7)
ext_box = add_rounded_rect(slide, ext_x, ext_y, ext_w, Inches(4.0), LIGHT_GRAY,
                            line_color=DARK_GRAY, line_w=1.5)
add_textbox(slide, ext_x, ext_y, ext_w, Inches(0.3),
            "External Services", size=11, bold=True, color=DARK_GRAY,
            align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

ext_items = [
    ("Perplexity", "(UIスクレイピング)"),
    ("ChatGPT API", "(OpenAI)"),
    ("Gemini", "(UI/API)"),
    ("AI Overviews", "(SerpAPI)"),
    ("Slack/Email", "(通知)"),
    ("Datadog/Sentry", "(監視)"),
]
for i, (name, desc) in enumerate(ext_items):
    iy = ext_y + Inches(0.4) + Inches(0.55) * i
    add_rect(slide, ext_x + Inches(0.1), iy, ext_w - Inches(0.2), Inches(0.5), WHITE,
             line_color=GRAY, line_w=1)
    add_textbox(slide, ext_x + Inches(0.1), iy, ext_w - Inches(0.2), Inches(0.25),
                name, size=9, bold=True, color=DARK_GRAY,
                align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, ext_x + Inches(0.1), iy + Inches(0.25), ext_w - Inches(0.2), Inches(0.25),
                desc, size=7, color=GRAY, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# Arrows API ↔ Auth, API ↔ Services
add_arrow(slide, Inches(9.5), api_y + Inches(0.35), Inches(9.3), api_y + Inches(0.35),
          color=DARK_GRAY, width=1)

# Arrow Services ↔ Worker (compound)
for i in range(5):
    sx = Inches(0.5) + (svc_w + Inches(0.1)) * i + svc_w/2
    add_arrow(slide, sx, svc_y + svc_h, sx, worker_y, color=DARK_GRAY, width=0.75)

# Arrow Worker ↔ External (one bidirectional concept arrow)
add_arrow(slide, Inches(10.7), worker_y + Inches(0.35), ext_x, worker_y + Inches(0.35),
          color=RED, width=1.5, dashed=True)

# Arrow API/Services ↔ Data Layer (compound)
add_arrow(slide, Inches(2.9), svc_y + svc_h, Inches(2.9), data_y, color=DARK_GRAY, width=0.75)
add_arrow(slide, Inches(7.0), worker_y + Inches(0.7), Inches(7.0), data_y, color=DARK_GRAY, width=0.75)

# Legend / star annotation
add_textbox(slide, Inches(0.5), Inches(7.05), Inches(12.0), Inches(0.3),
            "★ Citation Engine (F3) はMVPの中核。LLM Worker Poolで毎日4LLM×全プロンプトを自動実行し、引用率を算出",
            size=10, bold=True, color=RED, align=PP_ALIGN.CENTER)

add_footer(slide, 3)


# ============================================================
# Slide 5: Input → Process → Output
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_bar(slide, "④インプット → 処理 → アウトプット図", "システム視点: 何を入れたら何が出るか", color=GREEN)

# 3 columns layout
col_w = Inches(3.8)
col_h = Inches(5.3)
col_y = Inches(1.0)
gap = Inches(0.4)
total_w = col_w * 3 + gap * 2
col_start_x = (prs.slide_width - total_w) / 2

# Input column
input_x = col_start_x
add_rounded_rect(slide, input_x, col_y, col_w, col_h, LIGHT_BLUE, line_color=BLUE, line_w=2)
add_textbox(slide, input_x, col_y, col_w, Inches(0.5),
            "📥 INPUT", size=18, bold=True, color=BLUE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
add_textbox(slide, input_x, col_y + Inches(0.5), col_w, Inches(0.4),
            "システムが受け取るもの", size=10, color=DARK_GRAY, align=PP_ALIGN.CENTER)

input_items = [
    "🏢 自社情報",
    "・自社URL / 主要ドメイン\n・サブブランド・買収ブランド\n・対象市場(言語/地域)",
    "🎯 競合・業界情報",
    "・競合URL(最大10社)\n・業界キーワード",
    "❓ プロンプト",
    "・Money Prompt 30-1,000本\n・ファネル/ペルソナタグ\n(AI自動生成 or 手動)",
    "🔗 連携データ",
    "・CRM(Salesforce/HubSpot)\n・営業会話(Gong)\n・GA4 / Search Console",
    "⚙️ 設定",
    "・閾値(60%/70%等)\n・通知先(Slack/Email)\n・レポート頻度",
]
item_y = col_y + Inches(1.0)
for item in input_items:
    is_header = item.startswith(("🏢", "🎯", "❓", "🔗", "⚙️"))
    h = Inches(0.32) if is_header else Inches(0.6)
    add_textbox(slide, input_x + Inches(0.15), item_y, col_w - Inches(0.3), h,
                item, size=11 if is_header else 9, bold=is_header,
                color=NAVY if is_header else DARK_GRAY,
                align=PP_ALIGN.LEFT, va=MSO_ANCHOR.TOP)
    item_y += h + Inches(0.02)

# Process column (center)
proc_x = col_start_x + col_w + gap
add_rounded_rect(slide, proc_x, col_y, col_w, col_h, LIGHT_PURPLE, line_color=PURPLE, line_w=2)
add_textbox(slide, proc_x, col_y, col_w, Inches(0.5),
            "⚙️ PROCESS", size=18, bold=True, color=PURPLE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
add_textbox(slide, proc_x, col_y + Inches(0.5), col_w, Inches(0.4),
            "AIOutmation 内部の処理", size=10, color=DARK_GRAY, align=PP_ALIGN.CENTER)

proc_items = [
    "🤖 プロンプト自動生成",
    "Money Promptをファネル別に\nAIが300本自動提案",
    "🔄 LLM並列実行 (毎日)",
    "Perplexity/ChatGPT/Gemini/\nGoogle AI Overviews を24h自動",
    "🔍 引用判定",
    "出典URL抽出+ドメイン照合\n→ True Positive判定95%精度",
    "📊 Citation Rate算出",
    "(引用回数÷チェック数)×100\n+ノイズ除去(複数回×移動平均)",
    "🔔 異常検知・アラート",
    "閾値超え/急変/競合急伸を\n自動検知 → Slack/メール",
    "📝 AIコメンタリー生成",
    "前月比/カテゴリ/競合差分を\nClaude/GPTで3行サマリ化",
]
item_y = col_y + Inches(1.0)
for item in proc_items:
    is_header = item.startswith(("🤖", "🔄", "🔍", "📊", "🔔", "📝"))
    h = Inches(0.32) if is_header else Inches(0.5)
    add_textbox(slide, proc_x + Inches(0.15), item_y, col_w - Inches(0.3), h,
                item, size=11 if is_header else 9, bold=is_header,
                color=PURPLE if is_header else DARK_GRAY,
                align=PP_ALIGN.LEFT, va=MSO_ANCHOR.TOP)
    item_y += h + Inches(0.02)

# Output column
out_x = col_start_x + (col_w + gap) * 2
add_rounded_rect(slide, out_x, col_y, col_w, col_h, LIGHT_GREEN, line_color=GREEN, line_w=2)
add_textbox(slide, out_x, col_y, col_w, Inches(0.5),
            "📤 OUTPUT", size=18, bold=True, color=GREEN, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
add_textbox(slide, out_x, col_y + Inches(0.5), col_w, Inches(0.4),
            "ユーザーが受け取るもの", size=10, color=DARK_GRAY, align=PP_ALIGN.CENTER)

out_items = [
    "📈 ダッシュボード",
    "・引用率トレンドグラフ\n・FAQ管理表/記録表/集計表\n・ヒートマップ",
    "🚨 リアルタイムアラート",
    "・Slack通知(閾値超え)\n・競合急伸通知\n・センチメント悪化通知",
    "📄 月次レポート",
    "・ホワイトラベル化PDF\n・代理店ロゴ・カラー反映\n・経営層向け3行サマリ",
    "💡 改修レコメンド",
    "・改修対象ページ提案\n・JSON-LD/Schemaコード\n・競合勝ちパターン",
    "🗂️ データエクスポート",
    "・CSV全データダウンロード\n・API/MCPデータ取得\n・Looker/BigQuery連携(V2)",
]
item_y = col_y + Inches(1.0)
for item in out_items:
    is_header = item.startswith(("📈", "🚨", "📄", "💡", "🗂️"))
    h = Inches(0.32) if is_header else Inches(0.6)
    add_textbox(slide, out_x + Inches(0.15), item_y, col_w - Inches(0.3), h,
                item, size=11 if is_header else 9, bold=is_header,
                color=GREEN if is_header else DARK_GRAY,
                align=PP_ALIGN.LEFT, va=MSO_ANCHOR.TOP)
    item_y += h + Inches(0.02)

# Big arrows between columns
arrow_y = col_y + Inches(2.5)
add_arrow(slide, input_x + col_w, arrow_y, proc_x, arrow_y, color=NAVY, width=4)
add_arrow(slide, proc_x + col_w, arrow_y, out_x, arrow_y, color=NAVY, width=4)

# Bottom value statement
val_y = Inches(6.45)
add_rect(slide, Inches(0.5), val_y, Inches(12.3), Inches(0.45), NAVY)
add_textbox(slide, Inches(0.5), val_y, Inches(12.3), Inches(0.45),
            "→ 結果: 年間1,500時間の手作業を100時間に圧縮、Citation Rate +25pt、AI経由パイプライン3倍",
            size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

add_footer(slide, 4)


# ============================================================
# Slide 6: AIO Pyramid × Feature Mapping
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_bar(slide, "⑤AIOピラミッド × AIOutmation機能 対応図", "概念視点: 瀧内本の体系をどう実装しているか", color=RED)

# Pyramid on left side (5 levels)
# Top = 運用層, then LLMO, GEO, AEO, SEO (bottom)
center_x = Inches(3.5)  # pyramid center
top_y = Inches(1.0)
total_h = Inches(5.5)

# Build pyramid as 5 horizontal layers
pyramid_layers = [
    # (label_short, technology, color, width_ratio, top_offset_ratio)
    ("最上層: 運用層", "PDCA・運用", RED, 0.30, 0.0),
    ("3層目: 学習対象化", "LLMO", PURPLE, 0.50, 0.18),
    ("2層目: 引用層", "GEO", ORANGE, 0.65, 0.36),
    ("1層目: 理解層", "AEO", BLUE, 0.80, 0.54),
    ("土台: 基礎構造", "SEO", GREEN, 0.95, 0.72),
]

layer_h_unit = total_h * 0.18
max_width = Inches(5.5)

for i, (label, tech, c, w_ratio, top_off) in enumerate(pyramid_layers):
    w = max_width * w_ratio
    y = top_y + total_h * top_off
    x = center_x - w/2
    # Light fill
    light_fill = {
        RED: LIGHT_RED, PURPLE: LIGHT_PURPLE, ORANGE: LIGHT_ORANGE,
        BLUE: LIGHT_BLUE, GREEN: LIGHT_GREEN
    }[c]
    add_rect(slide, x, y, w, layer_h_unit, light_fill, line_color=c, line_w=2)
    add_textbox(slide, x, y, w, layer_h_unit * 0.6,
                label, size=11, bold=True, color=c, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
    add_textbox(slide, x, y + layer_h_unit * 0.5, w, layer_h_unit * 0.5,
                tech, size=14, bold=True, color=c, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

# AIOutmation features on right side, mapped to each layer with arrows
right_x = Inches(8.0)
right_w = Inches(5.0)

feature_layers = [
    # (y_pos, primary_features, color)
    (top_y + Inches(0.0), ["F4: ダッシュボード+アラート", "F6: AIコメンタリー(3行サマリ)", "F5: 月次レポート自動生成"], RED),
    (top_y + Inches(1.0), ["F6拡張(V2): ROI物語AI", "V3: llms.txt自動生成", "F3: 継続的引用蓄積"], PURPLE),
    (top_y + Inches(2.0), ["F3: 引用率測定エンジン ★中核", "F4: 競合分析", "F8(V2): Evidence/Reference"], ORANGE),
    (top_y + Inches(3.0), ["F2: プロンプト自動生成", "F8(V2): JSON-LD自動生成", "V3: 文章構成最適化AI"], BLUE),
    (top_y + Inches(4.0), ["F7(V2): SEO×AIO監査エンジン", "F9(V2): CMS PR自動起票", "Schema実装支援"], GREEN),
]

for ry, features, c in feature_layers:
    # Feature box
    box_h = Inches(0.95)
    light_fill = {
        RED: LIGHT_RED, PURPLE: LIGHT_PURPLE, ORANGE: LIGHT_ORANGE,
        BLUE: LIGHT_BLUE, GREEN: LIGHT_GREEN
    }[c]
    add_rounded_rect(slide, right_x, ry, right_w, box_h, light_fill, line_color=c, line_w=1.5)
    text = "\n".join("・" + f for f in features)
    add_textbox(slide, right_x + Inches(0.1), ry, right_w - Inches(0.2), box_h,
                text, size=10, color=DARK_GRAY, align=PP_ALIGN.LEFT, va=MSO_ANCHOR.MIDDLE)

# Arrows from pyramid to feature boxes
for i, (ry, _, c) in enumerate(feature_layers):
    arrow_y_start = top_y + total_h * pyramid_layers[i][4] + layer_h_unit / 2
    # Pyramid right edge x at this layer
    w_ratio = pyramid_layers[i][3]
    px = center_x + (max_width * w_ratio) / 2
    add_arrow(slide, px, arrow_y_start, right_x, ry + Inches(0.5), color=c, width=1.5)

# Header label
add_textbox(slide, Inches(0.5), top_y - Inches(0.4), Inches(6.5), Inches(0.3),
            "AIOピラミッド (瀧内本)", size=12, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)
add_textbox(slide, right_x, top_y - Inches(0.4), right_w, Inches(0.3),
            "AIOutmation の対応機能", size=12, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)

# Bottom message
msg_y = Inches(6.7)
add_rect(slide, Inches(0.5), msg_y, Inches(12.3), Inches(0.45), NAVY)
add_textbox(slide, Inches(0.5), msg_y, Inches(12.3), Inches(0.45),
            "瀧内本「AIO教科書」が示す全層を、ソフトウェアで一気通貫に実装する世界初の統合プラットフォーム",
            size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

add_footer(slide, 5)


# Save
output = "/home/user/AIO/AIOutmation_Explanatory_Diagrams.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
