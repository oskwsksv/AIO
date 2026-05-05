"""Generate PPT for the pain-resolution analysis (3 tools vs Top 5 pains)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# ======================== DATA ========================

OVERVIEW_MATRIX = [
    # (Pain, Profound, Peec, Passionfruit, Common residual)
    ("1. 時間消費が桁違い", "○", "○", "○", "戦略・解釈・改修工数は残る"),
    ("2. 属人化・知識流出リスク", "○", "△", "△", "「なぜそのプロンプトか」の文脈は人依存"),
    ("3. ノイズ vs シグナル分離", "△", "△", "△", "3社とも統計的処理は未実装"),
    ("4. 経営/クライアント説得力", "○", "△", "○", "自動3行サマリ・業界ベンチは未搭載"),
    ("5. スケール壁", "◎", "◎", "○", "日本語LLM・多言語は Peec が突出"),
]

# Per-pain detail tables: rows = aspects, cols = (Profound, Peec, Passionfruit)
PAIN_1_DETAIL = {
    "title": "Pain 1: 時間消費が桁違い",
    "subtitle": "手作業1,500時間/年クラスのデータ収集と運用",
    "rows": [
        ("データ収集の自動化", "◎ 10+モデル毎日自動", "◎ UIスクレイピング 24h", "◎ 5モデル自動実行"),
        ("プロンプト初期設計", "○ CSMガイド(4週間)", "△ 自助/テンプレ", "△ 自助"),
        ("コンテンツ改修工数", "△ 測定するが改修しない", "△ 同", "○ Agentic Recommendations"),
        ("レポートコメント執筆", "△ ダッシュボード提示まで", "△ 同", "△ 同"),
        ("戦略立案・解釈", "✕ 人が解釈", "✕ 人が解釈", "✕ 人が解釈"),
    ],
    "verdict": "◆ 実態:測定の自動化(年700〜1,000h削減)は実現するが、コンテンツ改修・PR・経営報告の作文は残る(年500〜700h)\n◆ 残存ペイン: 戦略・解釈・改修・レポート作文という「人が考える部分」は依然手作業",
}

PAIN_2_DETAIL = {
    "title": "Pain 2: 属人化・知識流出リスク",
    "subtitle": "担当者の休暇・退職で運用停止",
    "rows": [
        ("データ・ダッシュボード永続化", "◎", "◎", "◎"),
        ("「なぜこのプロンプトか」の文脈記録", "△ コメント機能のみ", "△ プロンプトタグのみ", "△"),
        ("改修ロジック・勝ちパターン共有", "△", "△", "△"),
        ("新人オンボーディング教材", "○ Profound Academy", "△ Docsのみ", "△ Docsのみ"),
        ("退職時の引継ぎSOP生成", "✕", "✕", "✕"),
    ],
    "verdict": "◆ 実態: データそのものは残るが、「どの数字を、なぜ、どう解釈するか」のナレッジは個人脳内\n◆ 残存ペイン: 戦略文脈・改修判断ロジック・社内SOPは依然属人的",
}

PAIN_3_DETAIL = {
    "title": "Pain 3: ノイズ vs シグナル分離不能 ★最大の未解決領域",
    "subtitle": "LLM揺らぎを「変化」と誤検知 or 本物の変化を見逃し",
    "rows": [
        ("統計的信頼区間の提示", "✕", "✕", "✕"),
        ("モデル仕様変更の検知", "△ ログのみ", "△", "△"),
        ("同一プロンプトの複数回実行で揺らぎ吸収", "△ 1日1回", "△ 1日1回", "△"),
        ("アラート閾値", "○ ユーザー設定", "○ 同", "○ 同"),
        ("「本物の変化」AI判定", "✕", "✕", "✕"),
    ],
    "verdict": "◆ 実態: LLM評価研究(Ai2 Signal/Noise framework等)は活発だが、3社とも商用に未実装\n◆ 残存ペイン: 業界全体の未開拓領域。あなたのプロダクトの差別化候補として最有望",
}

PAIN_4_DETAIL = {
    "title": "Pain 4: 経営/クライアント説得力の欠如",
    "subtitle": "ROI証明・予算獲得・解約防止に直結",
    "rows": [
        ("可視化(Visibility Radar/Heatmap等)", "◎ 多彩", "○ シンプル", "○"),
        ("ROI・売上アトリビューション", "△ ファネル接続あるが弱め", "✕ なし", "◎ GA4/GSC連携で売上直結"),
        ("業界ベンチマーク同梱", "△ Prompt Volumesは独自", "✕", "✕"),
        ("CMO向け3行サマリ自動生成", "✕", "✕", "✕"),
        ("競合との差分文章自動生成", "✕", "✕", "✕"),
        ("取締役会向けPDF自動生成", "△ エクスポートのみ", "△ 同", "△ 同"),
    ],
    "verdict": "◆ 実態: Passionfruitの収益アトリビューションが唯一の救い。3行サマリ・物語生成は3社とも未搭載\n◆ 残存ペイン: 所感コメント・ROI物語・業界ベンチ参照は依然手作業",
}

PAIN_5_DETAIL = {
    "title": "Pain 5: スケール壁(30→100→1,000本、多言語、サブブランド)",
    "subtitle": "戦略の幅・グローバル展開・サブブランド対応",
    "rows": [
        ("プロンプト規模上限", "◎ 数千〜万本(エンプラ)", "○ Pro 150本〜", "○ Enterprise $499+/月"),
        ("対応LLM数", "◎ 10+", "○ 6〜8(add-on含)", "○ 5"),
        ("対応言語数", "○ 40+", "◎ 115+(業界最広)", "△ 明示なし"),
        ("対応リージョン数", "◎ 200+", "◎ 全国対応", "○"),
        ("サブブランド管理", "◎", "◎", "○"),
        ("日本語LLM特化(Felo/Genspark等)", "△", "△ Perplexity日本語あり", "△"),
    ],
    "verdict": "◆ 実態: ProfoundとPeecは技術的スケール壁を概ね解消。Peecは115言語対応で多言語最強\n◆ 残存ペイン: 日本語特化LLMエコシステム、エンタープライズ価格帯($2k〜5k/月)の壁",
}

# Time allocation tables
TIME_A = [
    # (業務, 非導入時, 導入後, 削減量, 評価)
    ("データ収集・再計測", "800h", "80h", "▲720h", "◎"),
    ("プロンプト設計・タグ付け", "100h", "60h", "▲40h", "○"),
    ("コンテンツ改修・ブリーフ", "300h", "250h", "▲50h", "△"),
    ("経営レポート作文", "200h", "180h", "▲20h", "△"),
    ("戦略・解釈・部門調整", "150h", "150h", "±0", "✕"),
    ("【合計】", "1,550h", "720h", "▲830h(54%減)", ""),
]

TIME_B = [
    ("多クライアント手動運用", "600h", "60h", "▲540h", "◎"),
    ("月次レポート作成", "600h", "200h", "▲400h", "○"),
    ("ブリーフ作成・QA", "300h", "250h", "▲50h", "△"),
    ("クライアント説明・更新交渉", "150h", "150h", "±0", "✕"),
    ("【合計】", "1,650h", "660h", "▲990h(60%減)", ""),
]

# White space opportunities
WHITE_SPACE = [
    ("1", "ノイズ vs シグナル分離(統計的信頼区間)", "3社とも未実装、研究フロンティア",
     "「Signal/Noise Engine」: 複数回実行で揺らぎを統計処理し、信頼区間付きで変化を提示。学術研究(Ai2 framework)を実装", "★最有望"),
    ("2", "CMO向け3行サマリ自動生成", "全社ダッシュボード提示まで",
     "「Executive Summarizer Agent」: 前月比/カテゴリ/競合差分の3行を自動生成、ROI物語テンプレを内蔵", ""),
    ("3", "業界ベンチマーク同梱(Pain 4補強)", "Profound独自データ以外なし",
     "「Industry Benchmark Library」: 業種別中央値・上位10%値を同梱、自社の立ち位置を即可視化", ""),
    ("4", "計測→CMS改修PRの自動起票", "3社とも「測定で終わる」",
     "「Optimization Agent」: Schema・FAQ・原子段落構造をHeadless CMS/Webflowに自動PR起票", ""),
    ("5", "属人化解消のSOP自動生成", "全社未対応",
     "「Knowledge Auto-Documentation」: プロンプト追加時に「なぜ」「どう解釈するか」をAIが自動文書化", ""),
    ("6", "日本語LLM特化(Felo/Genspark等)", "Peec 115言語対応だが日本語LLMエコシステムは別問題",
     "「日本語AI検索完全対応」: Felo/Genspark/Yahoo!知恵袋AI/PKSHA等の日本独自LLMをカバー", ""),
    ("7", "コンテンツ・改修・PR・SEO統合", "3社とも測定特化、改修は別ツール",
     "「クローズドループ」: 測定→改修→再計測を1プラットフォームで完結", ""),
]


# ======================== BUILD ========================

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
BLUE = RGBColor(0x2E, 0x5C, 0x8A)
ORANGE = RGBColor(0xC9, 0x6E, 0x29)
RED = RGBColor(0xB0, 0x3A, 0x2E)
GREEN = RGBColor(0x2E, 0x7A, 0x4E)
LIGHT_BLUE = RGBColor(0xE8, 0xEE, 0xF7)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
LIGHT_RED = RGBColor(0xFF, 0xF0, 0xF0)
LIGHT_GREEN = RGBColor(0xF0, 0xF8, 0xE8)
LIGHT_YELLOW = RGBColor(0xFF, 0xF8, 0xE0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x22, 0x22, 0x22)


def set_text(tf, text, size=10, bold=False, color=BLACK, font_name="Yu Gothic", align=PP_ALIGN.LEFT):
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = MSO_ANCHOR.TOP
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
        r.font.size = Pt(size)
        r.font.name = font_name
        r.font.bold = bold
        r.font.color.rgb = color


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = fill_color
    if line_color is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line_color
        s.line.width = Pt(0.5)
    s.shadow.inherit = False
    return s


def add_box(slide, left, top, width, height, text, size=10, bold=False, color=BLACK, align=PP_ALIGN.LEFT, fill=None, va=MSO_ANCHOR.TOP):
    if fill is not None:
        add_rect(slide, left, top, width, height, fill, line_color=RGBColor(0x99, 0x99, 0x99))
    tb = slide.shapes.add_textbox(left, top, width, height)
    set_text(tb.text_frame, text, size=size, bold=bold, color=color, align=align)
    tb.text_frame.vertical_anchor = va
    return tb


def title_bar(slide, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.65), NAVY)
    add_box(slide, Inches(0.3), Inches(0.07), Inches(13), Inches(0.35), title, size=20, bold=True, color=WHITE)
    if subtitle:
        add_box(slide, Inches(0.3), Inches(0.4), Inches(13), Inches(0.25), subtitle, size=10, color=WHITE)


def grade_color(grade):
    g = grade.strip()[0] if grade else " "
    if g == "◎":
        return GREEN
    if g == "○":
        return BLUE
    if g == "△":
        return ORANGE
    if g == "✕":
        return RED
    return BLACK


# === Slide builders ===

def cover():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, prs.slide_height, NAVY)
    add_box(slide, Inches(0.5), Inches(1.8), Inches(12.3), Inches(1.0),
            "ベンチマーク3社は", size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(2.7), Inches(12.3), Inches(1.4),
            "「Top 5ペイン」をどこまで解決するか", size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.6),
            "Profound / Peec AI / Passionfruit Labs ─ 解決度・残存ペイン・ホワイトスペース分析", size=14, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(5.3), Inches(12.3), Inches(0.4),
            "結論: 完全解決は約4割、残り6割は形を変えて残存", size=14, bold=True, color=LIGHT_YELLOW, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(6.1), Inches(12.3), Inches(0.4),
            "2026年5月 作成", size=11, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


def divider(text, color, sub=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, prs.slide_height, color)
    add_box(slide, Inches(0.5), Inches(2.8), Inches(12.3), Inches(1.5),
            text, size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if sub:
        add_box(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(0.5),
                sub, size=16, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


def overview_matrix_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, "全体マップ：解決度評価", "凡例: ◎=ほぼ解決 / ○=部分解決 / △=表面的 / ✕=未解決")

    margin_left = Inches(0.3)
    margin_top = Inches(0.95)
    cols = [Inches(3.4), Inches(1.4), Inches(1.4), Inches(1.7), Inches(4.8)]
    headers = ["Top 5 ペイン", "Profound", "Peec AI", "Passionfruit", "共通残存ペイン"]
    header_h = Inches(0.55)
    row_h = Inches(0.85)

    x = margin_left
    for w, h in zip(cols, headers):
        add_box(slide, x, margin_top, w, header_h, h,
                size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY, va=MSO_ANCHOR.MIDDLE)
        x += w

    for i, (pain, p, pe, pf, residual) in enumerate(OVERVIEW_MATRIX):
        y = margin_top + header_h + row_h * i
        x = margin_left
        items = [
            (pain, LIGHT_BLUE, True, PP_ALIGN.LEFT, NAVY, 12),
            (p, WHITE, True, PP_ALIGN.CENTER, grade_color(p), 28),
            (pe, WHITE, True, PP_ALIGN.CENTER, grade_color(pe), 28),
            (pf, WHITE, True, PP_ALIGN.CENTER, grade_color(pf), 28),
            (residual, LIGHT_RED, False, PP_ALIGN.LEFT, BLACK, 11),
        ]
        for w, (txt, fill, bold, align, color, sz) in zip(cols, items):
            add_box(slide, x, y, w, row_h, txt, size=sz, bold=bold, color=color, align=align, fill=fill, va=MSO_ANCHOR.MIDDLE)
            x += w

    note_y = margin_top + header_h + row_h * len(OVERVIEW_MATRIX) + Inches(0.15)
    add_box(slide, margin_left, note_y, Inches(12.7), Inches(0.45),
            "★ Pain 3「ノイズ vs シグナル分離」は3社すべて△=業界全体の未開拓領域、ホワイトスペースとして最有望",
            size=12, bold=True, color=RED, fill=LIGHT_YELLOW)


def pain_detail_slide(pain_data, accent_color):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, pain_data["title"], pain_data["subtitle"])
    add_rect(slide, Inches(0), Inches(0.65), Inches(0.15), prs.slide_height - Inches(0.65), accent_color)

    margin_left = Inches(0.3)
    margin_top = Inches(0.9)
    cols = [Inches(4.5), Inches(2.7), Inches(2.7), Inches(2.7)]
    headers = ["観点", "Profound", "Peec AI", "Passionfruit"]
    header_h = Inches(0.5)
    n_rows = len(pain_data["rows"])
    available_h = Inches(7.5 - 0.9 - 0.5 - 1.4)  # leave space for verdict
    row_h = Inches(available_h.inches / n_rows)

    x = margin_left
    for w, h in zip(cols, headers):
        add_box(slide, x, margin_top, w, header_h, h,
                size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY, va=MSO_ANCHOR.MIDDLE)
        x += w

    for i, row in enumerate(pain_data["rows"]):
        aspect = row[0]
        values = row[1:]
        y = margin_top + header_h + row_h * i
        add_box(slide, margin_left, y, cols[0], row_h, aspect,
                size=11, bold=True, color=NAVY, align=PP_ALIGN.LEFT, fill=LIGHT_BLUE, va=MSO_ANCHOR.MIDDLE)
        for j, val in enumerate(values):
            xx = margin_left + sum(c for c in cols[:j+1])
            length = len(val)
            if length > 30:
                fs = 10
            elif length > 15:
                fs = 11
            else:
                fs = 13
            color = grade_color(val)
            add_box(slide, xx, y, cols[j+1], row_h, val,
                    size=fs, bold=True, color=color, align=PP_ALIGN.CENTER, fill=WHITE, va=MSO_ANCHOR.MIDDLE)

    # Verdict box at bottom
    verdict_y = margin_top + header_h + row_h * n_rows + Inches(0.15)
    add_box(slide, margin_left, verdict_y, Inches(12.7), Inches(1.1),
            pain_data["verdict"], size=12, color=BLACK, fill=LIGHT_YELLOW)


def time_slide(title, subtitle, data, accent_color):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, title, subtitle)
    add_rect(slide, Inches(0), Inches(0.65), Inches(0.15), prs.slide_height - Inches(0.65), accent_color)

    margin_left = Inches(1.0)
    margin_top = Inches(1.2)
    cols = [Inches(4.0), Inches(2.0), Inches(2.0), Inches(2.5), Inches(0.8)]
    headers = ["業務", "非導入時(手動)", "ツール導入後", "削減量", "評価"]
    header_h = Inches(0.55)
    row_h = Inches(0.7)

    x = margin_left
    for w, h in zip(cols, headers):
        add_box(slide, x, margin_top, w, header_h, h,
                size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY, va=MSO_ANCHOR.MIDDLE)
        x += w

    for i, row in enumerate(data):
        y = margin_top + header_h + row_h * i
        x = margin_left
        is_total = "合計" in row[0]
        bg = LIGHT_BLUE if is_total else WHITE
        for j, val in enumerate(row):
            color = NAVY if is_total else BLACK
            if j == 4:  # 評価 column
                color = grade_color(val) if val else BLACK
                sz = 18
            elif is_total:
                sz = 13
            else:
                sz = 12
            align = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            add_box(slide, x, y, cols[j], row_h, val,
                    size=sz, bold=is_total or (j == 4), color=color, align=align, fill=bg, va=MSO_ANCHOR.MIDDLE)
            x += cols[j]


def whitespace_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, "3社で解決されないペイン:7つのホワイトスペース", "あなたのプロダクトの差別化候補")

    margin_left = Inches(0.2)
    margin_top = Inches(0.9)
    cols = [Inches(0.4), Inches(2.8), Inches(2.5), Inches(6.4), Inches(0.8)]
    headers = ["#", "未解決ペイン", "競合状況", "プロダクト機会", ""]
    header_h = Inches(0.45)
    row_h = Inches(0.78)

    x = margin_left
    for w, h in zip(cols, headers):
        add_box(slide, x, margin_top, w, header_h, h,
                size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY, va=MSO_ANCHOR.MIDDLE)
        x += w

    for i, (num, pain, status, opportunity, badge) in enumerate(WHITE_SPACE):
        y = margin_top + header_h + row_h * i
        x = margin_left
        is_top = badge == "★最有望"
        row_bg = LIGHT_YELLOW if is_top else WHITE
        items = [
            (num, LIGHT_BLUE, True, PP_ALIGN.CENTER, NAVY, 14),
            (pain, row_bg, True, PP_ALIGN.LEFT, RED if is_top else BLACK, 10),
            (status, row_bg, False, PP_ALIGN.LEFT, BLACK, 9),
            (opportunity, LIGHT_GREEN if is_top else WHITE, False, PP_ALIGN.LEFT, BLACK, 9),
            (badge, LIGHT_YELLOW if badge else WHITE, True, PP_ALIGN.CENTER, RED, 10),
        ]
        for w, (txt, fill, bold, align, color, sz) in zip(cols, items):
            add_box(slide, x, y, w, row_h, txt, size=sz, bold=bold, color=color, align=align, fill=fill, va=MSO_ANCHOR.MIDDLE)
            x += w


def positioning_matrix_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, "戦略ポジショニング:未解決ペイン × 既存3社の関心度", "★Signal/Noise Engine が最有望ホワイトスペース")

    # Draw axes
    chart_left = Inches(2.5)
    chart_top = Inches(1.2)
    chart_w = Inches(8.5)
    chart_h = Inches(5.5)

    # Background
    add_rect(slide, chart_left, chart_top, chart_w, chart_h, LIGHT_BLUE)

    # Quadrant divider
    mid_x = chart_left + Inches(chart_w.inches / 2)
    mid_y = chart_top + Inches(chart_h.inches / 2)
    add_rect(slide, mid_x - Inches(0.005), chart_top, Inches(0.01), chart_h, NAVY)
    add_rect(slide, chart_left, mid_y - Inches(0.005), chart_w, Inches(0.01), NAVY)

    # Axes labels
    add_box(slide, Inches(0.4), Inches(1.2), Inches(2.0), Inches(0.4),
            "高 ↑", size=14, bold=True, color=NAVY)
    add_box(slide, Inches(0.4), Inches(3.5), Inches(2.0), Inches(1.0),
            "解決価値", size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
    add_box(slide, Inches(0.4), Inches(6.4), Inches(2.0), Inches(0.4),
            "低", size=14, bold=True, color=NAVY)

    add_box(slide, chart_left, Inches(6.85), chart_w, Inches(0.4),
            "低 ←──── 既存3社の関心度 ────→ 高", size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # Place dots for opportunities
    # Format: (x_pos_in, y_pos_in, label)
    points = [
        (chart_left.inches + 1.5, chart_top.inches + 0.6, "★ Signal/Noise Engine\n(#1) 最有望", LIGHT_YELLOW, RED, True),
        (chart_left.inches + 1.0, chart_top.inches + 1.6, "業界ベンチ (#3)", WHITE, NAVY, False),
        (chart_left.inches + 5.5, chart_top.inches + 1.0, "CMO 3行サマリ (#2)", LIGHT_GREEN, NAVY, False),
        (chart_left.inches + 1.0, chart_top.inches + 2.6, "日本語LLM (#6)", WHITE, NAVY, False),
        (chart_left.inches + 1.2, chart_top.inches + 3.6, "SOP自動化 (#5)", WHITE, NAVY, False),
        (chart_left.inches + 5.5, chart_top.inches + 3.5, "CMS PR起票 (#4)", WHITE, NAVY, False),
        (chart_left.inches + 5.5, chart_top.inches + 4.7, "クローズドループ (#7)", WHITE, NAVY, False),
    ]
    for x, y, label, fill, txt_color, bold in points:
        w = 2.4
        h = 0.7
        add_box(slide, Inches(x - w/2), Inches(y - h/2), Inches(w), Inches(h), label,
                size=10, bold=bold, color=txt_color, align=PP_ALIGN.CENTER, fill=fill, va=MSO_ANCHOR.MIDDLE)


def next_steps_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, "次のステップ候補", "プロダクト仮説の絞り込みと検証へ")

    items = [
        ("A", "プロトタイプ対象を1〜2個に絞る", "上記7つのホワイトスペースから選定するワークショップ"),
        ("B", "「Signal/Noise Engine」の技術仕様検討", "複数回実行アーキテクチャ、統計手法選定(Ai2 framework等)"),
        ("C", "「日本語LLMエコシステム特化」のリサーチ", "対象LLM一覧(Felo/Genspark/PKSHA/国産)、対応コスト試算"),
        ("D", "Willingness-to-Pay 仮説検証", "ペイン別の支払意欲を顧客インタビューで検証"),
    ]

    margin_left = Inches(0.5)
    margin_top = Inches(1.0)
    cols = [Inches(0.6), Inches(4.0), Inches(7.5)]
    header_h = Inches(0.55)
    row_h = Inches(1.2)

    x = margin_left
    for w, h in zip(cols, ["#", "アクション", "内容"]):
        add_box(slide, x, margin_top, w, header_h, h,
                size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY, va=MSO_ANCHOR.MIDDLE)
        x += w

    for i, (letter, headline, body) in enumerate(items):
        y = margin_top + header_h + row_h * i
        x = margin_left
        cells = [
            (letter, LIGHT_YELLOW, True, PP_ALIGN.CENTER, NAVY, 18),
            (headline, LIGHT_BLUE, True, PP_ALIGN.LEFT, NAVY, 14),
            (body, WHITE, False, PP_ALIGN.LEFT, BLACK, 12),
        ]
        for w, (txt, fill, bold, align, color, sz) in zip(cols, cells):
            add_box(slide, x, y, w, row_h, txt, size=sz, bold=bold, color=color, align=align, fill=fill, va=MSO_ANCHOR.MIDDLE)
            x += w


# === Build deck ===
cover()

divider("全体マップ:解決度評価", NAVY, "5ペイン × 3ツールのマトリクス")
overview_matrix_slide()

divider("ペイン別 詳細分析", BLUE, "Pain 1〜5を観点別に深掘り")
pain_detail_slide(PAIN_1_DETAIL, BLUE)
pain_detail_slide(PAIN_2_DETAIL, BLUE)
pain_detail_slide(PAIN_3_DETAIL, RED)  # ★ Most important
pain_detail_slide(PAIN_4_DETAIL, BLUE)
pain_detail_slide(PAIN_5_DETAIL, BLUE)

divider("解決度を時間で可視化", GREEN, "年間時間配分の Before / After")
time_slide("ペルソナA(社内担当)の年間時間配分", "ツール導入により54%削減、ただし戦略・解釈は残る", TIME_A, GREEN)
time_slide("ペルソナB(代理店担当)の年間時間配分", "マルチテナント対応で60%削減、提案・交渉は不変", TIME_B, GREEN)

divider("ホワイトスペース", ORANGE, "3社が解決していない7つの未開拓領域")
whitespace_slide()
positioning_matrix_slide()

divider("次のステップ", RED, "プロダクト仮説の絞り込みへ")
next_steps_slide()

output = "/home/user/AIO/PainResolutionAnalysis_GEO_AEO.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
