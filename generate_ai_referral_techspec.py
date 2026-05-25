"""Generate engineer-facing technical spec for AI referral channel detection in GA4."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def style_run(run, size=10, bold=False, italic=False, color=None, font_name="Yu Gothic"):
    run.font.size = Pt(size)
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def set_cell_bg(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def write_cell(cell, text, size=9, bold=False, bg=None, color=None, align=None, mono=False):
    cell.text = ""
    if bg:
        set_cell_bg(cell, bg)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    font = "Consolas" if mono else "Yu Gothic"
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(0)
        if align:
            p.alignment = align
        run = p.add_run(line)
        style_run(run, size=size, bold=bold, color=color, font_name=font)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level <= 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    sizes = {0: 24, 1: 17, 2: 13, 3: 11}
    colors = {0: "1F3A5F", 1: "1F3A5F", 2: "2E5C8A", 3: "444444"}
    style_run(run, size=sizes.get(level, 11), bold=True, color=colors.get(level, "000000"))


def add_paragraph(doc, text, size=10, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(text)
    style_run(run, size=size, bold=bold, color=color)


def add_bullet(doc, text, size=10, indent=0):
    p = doc.add_paragraph(style="List Bullet")
    if indent:
        p.paragraph_format.left_indent = Cm(0.5 + indent * 0.5)
    run = p.add_run(text)
    style_run(run, size=size)


def add_code_block(doc, code, lang_label=None):
    """Add a monospace code block with light gray background."""
    if lang_label:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(lang_label)
        style_run(r, size=8, bold=True, color="888888")
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, "F4F4F4")
    cell.text = ""
    lines = code.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.1
        r = p.add_run(line if line else " ")
        style_run(r, size=8.5, color="1A1A1A", font_name="Consolas")


def add_callout(doc, title, body, color="FFF4E0", title_color="C04000"):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, color)
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(3)
    r = p1.add_run(title)
    style_run(r, size=10, bold=True, color=title_color)
    p2 = cell.add_paragraph()
    p2.paragraph_format.line_spacing = 1.3
    r = p2.add_run(body)
    style_run(r, size=10)


def add_table(doc, headers, rows, widths):
    n = len(headers)
    t = doc.add_table(rows=len(rows) + 1, cols=n)
    t.autofit = False
    for r in t.rows:
        for i in range(n):
            r.cells[i].width = Cm(widths[i])
    for i, h in enumerate(headers):
        write_cell(t.cell(0, i), h, size=10, bold=True, bg="1F3A5F", color="FFFFFF")
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            bg = "E8EEF7" if ci == 0 else None
            write_cell(t.cell(ri + 1, ci), val, size=9, bold=(ci == 0), bg=bg)


def setup_doc(doc):
    section = doc.sections[0]
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    style = doc.styles["Normal"]
    style.font.name = "Yu Gothic"
    style.font.size = Pt(10)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "Yu Gothic")


# ============== Build ==============
doc = Document()
setup_doc(doc)

# Title
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(40)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("AI流入チャネル判定 技術仕様書")
style_run(r, size=26, bold=True, color="1F3A5F")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("リファラドメインの正規表現判定による「AI検索」チャネル分離")
style_run(r, size=14, color="2E5C8A")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(20)
r = p.add_run("対象: 自社エンジニアチーム  /  関連プロダクト: AIOutmation (引用率測定・収益アトリビューション)")
style_run(r, size=11, color="555555")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("作成: 2026年5月  /  Version 1.0")
style_run(r, size=10, color="888888")

doc.add_page_break()

# 1. Background
add_heading(doc, "1. 背景と課題", level=1)
add_paragraph(doc,
"ユーザーがChatGPT・Perplexity等のAI回答内リンクをクリックすると、ブラウザは document.referrer に chatgpt.com 等を送信する。しかしGA4のデフォルトのチャネル定義には「AI検索」が存在しないため、これらの流入は『Referral』(リファラあり) または『Direct』(リファラ欠落時) に分類され、AI経由トラフィックの効果測定が不可能になっている。",
size=11)
add_callout(doc, "本仕様書の目的",
"リファラ/ソースのドメインを正規表現で判定し、独自の『AI検索』チャネルへ再分類する実装方法を、3つのレイヤー(GA4 UI / BigQuery / GTM)で具体的に定義する。MVP段階では『方法A支援 + 方法B実装』から着手し、V2で方法C・AI Overviews補完へ拡張する。",
color="E0F4E0", title_color="2E7A4E")

# 2. Architecture overview
add_heading(doc, "2. 実装レイヤー全体像", level=1)
add_table(doc,
    ["レイヤー", "方法", "用途", "実装時期"],
    [
        ("顧客GA4の即時可視化", "方法A: GA4カスタムチャネルグループ", "ノーコードで顧客環境に設定支援", "MVP"),
        ("プロダクト本体の集計", "方法B: BigQueryエクスポート + SQL", "プログラマティック・大規模対応", "MVP"),
        ("リファラ欠落対策", "方法C: GTM + dataLayer", "捕捉精度向上(高度プラン)", "V2"),
        ("AI Overviews補完", "Signal/Noise Engine + GSC相関", "業界未解決領域の差別化", "V2"),
    ],
    widths=[4.5, 6.0, 4.5, 2.0])

doc.add_page_break()

# 3. Method A
add_heading(doc, "3. 方法A: GA4 カスタムチャネルグループ (ノーコード)", level=1)
add_paragraph(doc, "用途: GA4のレポート画面・標準レポートで即座にAI流入を可視化したい場合。顧客オンボーディング時に設定支援する。", size=10)

add_heading(doc, "3.1 設定手順", level=2)
add_bullet(doc, "GA4管理画面 → 管理 → データの表示 → チャネルグループ")
add_bullet(doc, "「新しいチャネルグループを作成」をクリック")
add_bullet(doc, "新規チャネル「AI検索 (AI Search)」を追加")
add_bullet(doc, "条件設定: ディメンション「ソース(Source)」/ マッチタイプ「正規表現に一致」/ 値は下記正規表現")
add_bullet(doc, "【最重要】作成した「AI検索」チャネルを「Referral」「Organic Search」より上にドラッグ")

add_callout(doc, "⚠️ チャネル順序が最重要",
"GA4のチャネル分類は上から順に評価され、最初にマッチしたチャネルが適用される。「AI検索」が「Referral」より下にあると、chatgpt.com が先に「Referral」にマッチして取られ、AI検索チャネルが0セッションになる。必ず Referral より上に配置すること。",
color="FFE0E0", title_color="A02020")

add_heading(doc, "3.2 制約", level=2)
add_bullet(doc, "カスタムチャネルグループは作成時点以降のデータに適用(過去遡及は一部レポートのみ)")
add_bullet(doc, "GA4の正規表現は RE2構文")
add_bullet(doc, "デフォルトのチャネルグループは編集不可。新規チャネルグループを作成して使う")

doc.add_page_break()

# 4. Method B
add_heading(doc, "4. 方法B: BigQueryエクスポート (★プロダクト本命実装)", level=1)
add_paragraph(doc, "用途: AIOutmation本体で顧客のGA4データをプログラマティックに処理する。これが収益アトリビューション機能の中核実装。", size=10)

add_heading(doc, "4.1 前提", level=2)
add_bullet(doc, "顧客のGA4 → BigQueryエクスポートをON(GA4標準機能、無料枠あり)")
add_bullet(doc, "events_YYYYMMDD テーブルが日次生成される")
add_bullet(doc, "顧客のGCPプロジェクトへのBigQuery read権限、または自社プロジェクトへのデータ転送設計が必要")

add_heading(doc, "4.2 SQL実装例", level=2)
add_code_block(doc, """-- AI検索チャネルのセッション・収益を集計
WITH ai_sessions AS (
  SELECT
    user_pseudo_id,
    (SELECT value.int_value FROM UNNEST(event_params)
     WHERE key = 'ga_session_id') AS session_id,
    -- collected_traffic_source がGA4の最新リファラ情報
    collected_traffic_source.manual_source AS source,
    traffic_source.source AS first_source,
    -- page_referrer からドメイン抽出
    (SELECT value.string_value FROM UNNEST(event_params)
     WHERE key = 'page_referrer') AS page_referrer,
    event_name,
    ecommerce.purchase_revenue AS revenue
  FROM `your_project.analytics_XXXXXX.events_*`
  WHERE _TABLE_SUFFIX BETWEEN '20260401' AND '20260430'
)
SELECT
  CASE
    WHEN REGEXP_CONTAINS(
      COALESCE(source, page_referrer, ''),
      r'(chatgpt\\.com|chat\\.openai\\.com|openai\\.com|perplexity\\.ai|claude\\.ai|gemini\\.google\\.com|copilot\\.microsoft\\.com|deepseek\\.com|meta\\.ai|grok\\.com|felo\\.ai|genspark\\.ai)'
    ) THEN 'AI Search'
    ELSE 'Other'
  END AS channel,
  COUNT(DISTINCT CONCAT(user_pseudo_id, CAST(session_id AS STRING))) AS sessions,
  SUM(IF(event_name = 'purchase', revenue, 0)) AS total_revenue
FROM ai_sessions
GROUP BY channel;""", lang_label="BigQuery Standard SQL")

add_heading(doc, "4.3 実装ポイント", level=2)
add_bullet(doc, "collected_traffic_source (GA4の比較的新しいフィールド) が最も正確なリファラ情報を持つ")
add_bullet(doc, "フォールバックで page_referrer イベントパラメータも参照する(COALESCE)")
add_bullet(doc, "LLM別に分けたい場合は CASE を細分化 (chatgpt → 'ChatGPT', perplexity → 'Perplexity'...)")
add_bullet(doc, "セッション一意化は user_pseudo_id + ga_session_id の組合せで行う")
add_bullet(doc, "BigQueryの正規表現はバックスラッシュのエスケープに注意(r'...' のraw文字列を使う)")

add_heading(doc, "4.4 LLM別集計への拡張例", level=2)
add_code_block(doc, """CASE
  WHEN REGEXP_CONTAINS(src, r'chatgpt\\.com|openai\\.com') THEN 'ChatGPT'
  WHEN REGEXP_CONTAINS(src, r'perplexity\\.ai')            THEN 'Perplexity'
  WHEN REGEXP_CONTAINS(src, r'claude\\.ai|anthropic\\.com') THEN 'Claude'
  WHEN REGEXP_CONTAINS(src, r'gemini\\.google\\.com')       THEN 'Gemini'
  WHEN REGEXP_CONTAINS(src, r'copilot\\.microsoft\\.com')   THEN 'Copilot'
  WHEN REGEXP_CONTAINS(src, r'felo\\.ai')                  THEN 'Felo'
  WHEN REGEXP_CONTAINS(src, r'genspark\\.ai')              THEN 'Genspark'
  ELSE 'Other AI'
END AS llm_source""", lang_label="BigQuery - LLM別分類")

doc.add_page_break()

# 5. Method C
add_heading(doc, "5. 方法C: GTM + dataLayer (リファラ欠落対策・最堅牢)", level=1)
add_paragraph(doc, "用途: 一部のAIクライアントがリファラをstripする問題に対処したい場合(高度プラン顧客向け、V2)。", size=10)

add_heading(doc, "5.1 仕組み", level=2)
add_paragraph(doc, "ランディング時に document.referrer をJSで捕捉し、GA4のカスタムディメンションに送信する。", size=10)

add_heading(doc, "5.2 実装コード (GTM カスタムHTMLタグ)", level=2)
add_code_block(doc, """// GTM カスタムHTMLタグ (全ページ・DOM Ready時に発火)
(function() {
  var ref = document.referrer || '';
  var aiPattern = /(chatgpt\\.com|chat\\.openai\\.com|openai\\.com|perplexity\\.ai|claude\\.ai|gemini\\.google\\.com|copilot\\.microsoft\\.com|deepseek\\.com|meta\\.ai|grok\\.com|felo\\.ai|genspark\\.ai)/i;
  var aiSource = 'none';
  if (aiPattern.test(ref)) {
    aiSource = ref.match(aiPattern)[1];
  }
  // GA4にカスタムディメンションとして送信
  dataLayer.push({
    'event': 'ai_referral_check',
    'ai_source': aiSource
  });
})();""", lang_label="JavaScript (GTM Custom HTML)")

add_heading(doc, "5.3 実装ポイント", level=2)
add_bullet(doc, "GA4側で ai_source をカスタムディメンションとして登録する")
add_bullet(doc, "初回ランディング時のみ有効。セッション全体に紐付けるならlanding時にsessionStorage等へ保存する設計が必要")
add_bullet(doc, "リファラがあるセッションは確実に捕捉できるが、リファラ自体が無い(stripされた)場合は捕捉不可")

doc.add_page_break()

# 6. Regex full
add_heading(doc, "6. 完全版・正規表現 (日本語LLM含む)", level=1)
add_code_block(doc, """(chatgpt\\.com|chat\\.openai\\.com|openai\\.com|perplexity\\.ai|claude\\.ai|anthropic\\.com|gemini\\.google\\.com|bard\\.google\\.com|copilot\\.microsoft\\.com|deepseek\\.com|meta\\.ai|grok\\.com|x\\.ai|you\\.com|phind\\.com|poe\\.com|felo\\.ai|genspark\\.ai)""", lang_label="共通正規表現 (RE2 / JS両対応)")

add_table(doc,
    ["分類", "対象ドメイン"],
    [
        ("グローバル主要", "chatgpt.com, chat.openai.com, openai.com, perplexity.ai, claude.ai, anthropic.com, gemini.google.com, copilot.microsoft.com"),
        ("その他グローバル", "deepseek.com, meta.ai, grok.com / x.ai, you.com, phind.com, poe.com, bard.google.com"),
        ("日本語LLM(差別化要素)", "felo.ai, genspark.ai  ※今後の国産LLMは随時追加"),
    ],
    widths=[4.5, 15.0])

doc.add_page_break()

# 7. Edge cases
add_heading(doc, "7. 重要な技術的注意点 (エッジケース)", level=1)

add_heading(doc, "7.1 Google AI Overviews / AI Mode は分離困難", level=2)
add_paragraph(doc,
"AI Overviewsの引用リンク経由の流入は google.com からのorganic として来るため、リファラだけでは通常のGoogle検索と区別できない。",
size=10)
add_bullet(doc, "対処1: Google側が付与するURLパラメータ(utm_source等)を監視")
add_bullet(doc, "対処2: 推定モデル(GSCのAI Mode系インプレッション急増との相関)で補完")
add_bullet(doc, "完全な分離は現状の業界全体の未解決課題 = AIOutmationの Signal/Noise Engine が補完する領域(V2)")

add_heading(doc, "7.2 OpenAIが utm_source=chatgpt.com を付与し始めている", level=2)
add_paragraph(doc,
"2025年以降、ChatGPTが一部リンクに utm_source=chatgpt.com を自動付与。これは Source に明示的に現れるので捕捉が容易。正規表現は Source と page_referrer の両方を見るべき。",
size=10)

add_heading(doc, "7.3 リファラストリッピング", level=2)
add_paragraph(doc,
"一部のAIアプリ(ネイティブアプリ版ChatGPT等)はリファラを送らない → 「Direct」に落ちる。これは原理的に100%は捕捉不可。GSCのブランド検索ハロー(Passionfruit社の手法)で間接補完する。",
size=10)

add_heading(doc, "7.4 サブドメイン・パスでの判別", level=2)
add_paragraph(doc,
"x.com/i/grok のようにパスで判別が必要なケースは、page_referrer(フルURL)に対する正規表現が必要(Sourceドメインだけでは不足)。",
size=10)

# 8. Validation
add_heading(doc, "8. 検証方法", level=1)
add_bullet(doc, "Google Tag Assistant(Chrome拡張)をインストール")
add_bullet(doc, "実際にChatGPT/Perplexityの回答リンクから自社サイトを開く(または同僚に依頼)")
add_bullet(doc, "GA4管理画面 → DebugView でイベントストリームを確認")
add_bullet(doc, "Session source がAIドメインになっているか、カスタムチャネル「AI検索」に分類されているかをチェック")
add_bullet(doc, "BigQuery実装の場合: 上記SQLを過去データに対して実行し、AI Searchセッション数が0でないことを確認")

# 9. Implementation roadmap
add_heading(doc, "9. AIOutmationでの実装ロードマップ", level=1)
add_table(doc,
    ["フェーズ", "実装内容", "担当レイヤー"],
    [
        ("MVP", "方法A: 顧客オンボーディングで設定ガイド + テンプレート提供", "CS / フロントエンド"),
        ("MVP", "方法B: BigQueryコネクタ + 集計バッチ(日次)", "バックエンド / データ基盤"),
        ("MVP", "ダッシュボードにAI Search流入・収益を表示", "フロントエンド"),
        ("V2", "方法C: GTMタグ自動生成 + カスタムディメンション設定支援", "バックエンド"),
        ("V2", "LLM別収益アトリビューション(ChatGPT/Perplexity/Claude別)", "データ基盤"),
        ("V2", "AI Overviews補完(GSC相関 + Signal/Noise Engine)", "データサイエンス"),
        ("V2", "B2B拡張: Salesforce/HubSpot連携でMQL→SQL→受注アトリビューション", "バックエンド / 連携基盤"),
    ],
    widths=[2.5, 11.0, 6.0])

add_callout(doc, "B2B特化の差別化ポイント",
"Passionfruit LabsのGA4連携は主にEC・DTC(purchase/revenueイベントが取りやすい業態)に最適化されている。我々のターゲット(上場B2B SaaS)は購買が『商談→数ヶ月検討→受注』のためGA4だけでは追えない。GA4連携(Passionfruit同等)に加え、Salesforce/HubSpot連携でMQL→SQL→受注までアトリビューションすることが差別化軸となる。",
color="E0F4E0", title_color="2E7A4E")

add_paragraph(doc, "", size=8)
add_paragraph(doc, "─ 本仕様書はVer 1.0。実装着手時に最新のGA4スキーマ・各LLMのリファラ仕様を再確認すること。AIドメインは頻繁に追加・変更されるため、正規表現は設定ファイル化して運用中に更新可能な設計を推奨する。",
              size=9, color="888888")

doc.save("/home/user/AIO/AI_Referral_Channel_TechSpec.docx")
print("Saved: AI_Referral_Channel_TechSpec.docx")
