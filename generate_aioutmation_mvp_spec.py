"""Generate AIOutmation MVP Feature Specification Word doc for dev partner discussion."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_bg(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def style_run(run, size=10, bold=False, color=None, font_name="Yu Gothic"):
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
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def write_cell(cell, text, size=9, bold=False, bg=None, color=None, align=None):
    cell.text = ""
    if bg:
        set_cell_bg(cell, bg)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        if align:
            p.alignment = align
        run = p.add_run(line)
        style_run(run, size=size, bold=bold, color=color)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level <= 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    sizes = {0: 24, 1: 18, 2: 14, 3: 12, 4: 11}
    colors = {0: "1F3A5F", 1: "1F3A5F", 2: "2E5C8A", 3: "444444", 4: "555555"}
    style_run(run, size=sizes.get(level, 11), bold=True, color=colors.get(level, "000000"))


def add_paragraph(doc, text, size=10, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    style_run(run, size=size, bold=bold, color=color)


def add_bullet(doc, text, size=10, indent=0):
    p = doc.add_paragraph(style="List Bullet")
    if indent:
        p.paragraph_format.left_indent = Cm(0.5 + indent * 0.5)
    run = p.add_run(text)
    style_run(run, size=size)


def add_callout(doc, title, body, color="FFF4E0", title_color="C04000"):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, color)
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(3)
    r = p1.add_run(title)
    style_run(r, size=11, bold=True, color=title_color)
    p2 = cell.add_paragraph()
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
    new_w, new_h = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_w
    section.page_height = new_h
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    style = doc.styles["Normal"]
    style.font.name = "Yu Gothic"
    style.font.size = Pt(10)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "Yu Gothic")


def cover_page(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(80)
    r = p.add_run("AIOutmation")
    style_run(r, size=44, bold=True, color="1F3A5F")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MVP機能仕様書")
    style_run(r, size=22, bold=True, color="2E5C8A")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(30)
    r = p.add_run("委託開発会社ディスカッション用ドキュメント")
    style_run(r, size=14, color="555555")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(60)
    r = p.add_run("──── アウトプット③ ────")
    style_run(r, size=11, color="888888")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    r = p.add_run("MVP範囲 + 機能仕様 + 技術スタック + データモデル + 開発計画 + 論点")
    style_run(r, size=11, color="555555")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(40)
    r = p.add_run("作成: 2026年5月")
    style_run(r, size=10, color="888888")
    doc.add_page_break()


# ====== Build Document ======

doc = Document()
setup_doc(doc)
cover_page(doc)

# === 1. Executive Summary ===
add_heading(doc, "1. エグゼクティブサマリー", level=1)
add_paragraph(doc, """AIOutmationのMVPは、瀧内本「AIO教科書」のChapter 4 PDCAサイクルを完全自動化することを最優先とする。リーンキャンバスで定義した8機能のうち6機能を24週間(約6ヶ月)で実装し、3-5社のデザインパートナーとβ運用を経て、商用ローンチを目指す。

MVP範囲の意思決定基準:
1. 競合(Profound/Peec/Passionfruit)の最低限ラインに追いつく(=計測の自動化)
2. 代理店セグメントの必須要件(マルチテナント+ホワイトラベル)を満たす
3. 顕著な時間削減(月4時間→30分)で初期顧客のROIを即証明できる
4. R&D色の強い差別化機能(Signal/Noise Engine, CMS PR自動起票)はV2/V3に延期し、PMFを優先する""", size=11)

add_callout(doc,
    "MVP の北極星指標",
    "①β顧客5社の継続(NRR 100%以上) ②各顧客の引用率改善(導入前→導入3ヶ月後で平均15ポイント増) ③代理店パートナー1社獲得(クライアント3社以上を本プラットフォームで運用)",
    color="E0F4E0", title_color="2E7A4E")

# === 2. MVP Scope ===
add_heading(doc, "2. MVP範囲(In/Out of Scope)", level=1)

add_heading(doc, "2.1 MVP に含める機能(Phase 1)", level=2)
add_table(doc,
    ["#", "機能名", "リーンキャンバス対応", "教科書対応"],
    [
        ("F1", "マルチテナント基盤+認証", "⑧マルチテナント運用(基本)", "─"),
        ("F2", "プロンプト管理+AI自動生成", "③プロンプト自動生成エンジン", "Ch.1+4"),
        ("F3", "引用率測定エンジン(Citation Rate Engine)", "⑤運用(PDCA Check部分)", "Ch.4"),
        ("F4", "ダッシュボード+アラート", "⑤運用(可視化)", "Ch.4"),
        ("F5", "ホワイトラベル+月次レポート", "⑧マルチテナント+⑦経営報告(基本)", "Ch.4"),
        ("F6", "AIコメンタリー(3行サマリ)", "⑦Executive Reporter(基本)", "Ch.4 Act"),
    ],
    widths=[1.5, 6.5, 7.0, 4.0])

add_heading(doc, "2.2 V2 に延期する機能(Phase 2: 24-36週)", level=2)
add_bullet(doc, "F7: SEO×AIO監査エンジン(初期診断レポート) — Ch.1-3 統合監査")
add_bullet(doc, "F8: 文書構造化エンジン(JSON-LD自動生成) — Ch.3対応")
add_bullet(doc, "F9: CMS PR自動起票(GitHub/Webflow/WordPress連携) — Ch.3対応")
add_bullet(doc, "F10: ROI物語AI(過去成功事例DB+類似ストーリー提示) — 経営説明強化")
add_bullet(doc, "F11: 業界ベンチマーク同梱 — 顧客横断データの匿名化集計")
add_bullet(doc, "F12: Salesforce/HubSpot/Gong連携(プロンプト自動抽出) — オンボード強化")

add_heading(doc, "2.3 V3 に延期する機能(Phase 3: 36週以降)", level=2)
add_bullet(doc, "F13: 文章構成最適化AI(PREP/FAQ/HowTo自動リライト) — Ch.2対応 / 重い")
add_bullet(doc, "F14: Signal/Noise Engine(統計的信頼区間) — R&D色が強く市場検証後")
add_bullet(doc, "F15: 知識資産化アシスタント(llms.txt自動生成) — Ch.5対応")
add_bullet(doc, "F16: AI別最適化レコメンド(Perplexity/Gemini/ChatGPT別) — Ch.5対応")
add_bullet(doc, "F17: SOC 2 Type II対応 — エンタープライズ拡大時")

add_heading(doc, "2.4 MVPに含めない理由(意思決定根拠)", level=2)
add_table(doc,
    ["延期機能", "延期理由", "代替手段"],
    [
        ("CMS PR自動起票", "CMSの種類が多く(WordPress, Webflow, Headless各種)、各連携の実装コストが大。MVPではChapter 3の改修推奨を「テキストアウトプット」で提供。", "改修推奨をJSON/Markdownで出力、顧客側で手動適用"),
        ("Signal/Noise Engine", "学術研究レベルの実装(Ai2 framework等)で開発期間予測困難。閾値ベースアラートで競合と同等水準は確保可能。", "閾値ベースアラート+将来のV3でアップグレード"),
        ("文章構成最適化AI", "長文生成のLLMコストが顧客あたり月数千円〜数万円となりマージン悪化。需要検証後にスコープ確定。", "MVPでは改修ガイドのみ(リライト推奨は人手)"),
        ("業界ベンチマーク", "顧客横断データの蓄積に時間が必要。最低3-6ヶ月の運用データが必要。", "リリース後6ヶ月で順次提供"),
    ],
    widths=[5.0, 9.5, 5.0])

doc.add_page_break()

# === 3. Feature Specifications ===
add_heading(doc, "3. MVP機能仕様(F1〜F6)", level=1)

# Feature F1
add_heading(doc, "F1. マルチテナント基盤+認証", level=2)
add_callout(doc, "目的", "代理店向け運用(複数クライアント管理)+企業向けチーム運用の双方を1つの基盤で実現する。後付けは技術負債になるため、Phase 1で必須。", color="FFF4E0")

add_heading(doc, "スコープ", level=3)
add_table(doc,
    ["項目", "内容"],
    [
        ("認証方式", "Email/Password + Google SSO + Microsoft 365 SSO(MVPはGoogle SSOのみ必須、他はV2)"),
        ("組織構造", "3階層: Account(テナント) → Workspace(クライアント案件) → User(メンバー)"),
        ("権限", "RBAC: Owner / Admin / Editor / Viewer の4ロール"),
        ("シート制限", "全プラン無制限シート(=競合差別化として明確に打ち出す)"),
        ("Workspace切替", "サイドバーでアクティブWorkspaceを切替、URLにWorkspace IDを含める"),
        ("監査ログ", "MVP: 認証ログのみ / V2: 全操作ログ"),
    ],
    widths=[4.0, 15.5])

add_heading(doc, "ユーザーストーリー", level=3)
add_bullet(doc, "代理店PMとして、5社のクライアント案件をWorkspaceで分離管理したい")
add_bullet(doc, "代理店アカウント担当として、自分が担当するWorkspaceのみアクセスできるようにしたい")
add_bullet(doc, "事業会社マーケPMとして、SEO担当・コンテンツ担当・経営層をそれぞれの権限で招待したい")

add_heading(doc, "成功基準", level=3)
add_bullet(doc, "1代理店アカウントで5クライアント案件を運用可能")
add_bullet(doc, "Workspace切替が3秒以内に完了")
add_bullet(doc, "認証成功率99.5%以上")

# Feature F2
add_heading(doc, "F2. プロンプト管理+AI自動生成", level=2)
add_callout(doc, "目的", "プロンプト初期設計の属人性を解消し、オンボード工数を半日→1時間以内に圧縮。", color="FFF4E0")

add_heading(doc, "スコープ", level=3)
add_table(doc,
    ["項目", "内容"],
    [
        ("プロンプトCRUD", "作成・編集・削除・アーカイブ。最大5,000本/Workspace"),
        ("タグ機能", "ファネルステージ / ペルソナ / プロダクト / 地域 / 言語 の5軸タグ"),
        ("AI自動生成", "入力: 自社URL + 競合URL(最大10) + 業界キーワード\n出力: 30〜200本のプロンプト案(タグ付き)を5分以内に生成"),
        ("一括インポート", "CSV/TSVから一括インポート(タグ列対応)"),
        ("プロンプト文脈記録", "各プロンプトに「対象ICP/ファネル/採用根拠」をAIが自動文書化(編集可)"),
        ("優先順位スコア", "前月の引用率×推定ROI×競合プレッシャーから自動算出"),
    ],
    widths=[4.0, 15.5])

add_heading(doc, "AI自動生成の処理フロー", level=3)
add_bullet(doc, "Step 1: 自社URLを取得 → タイトル・h1・descriptionを抽出")
add_bullet(doc, "Step 2: 競合URLを同様に解析")
add_bullet(doc, "Step 3: LLM(Claude Sonnet等)に「ICP × 購買ファネル × 製品カテゴリ」のマトリクスでプロンプト生成を依頼")
add_bullet(doc, "Step 4: 生成プロンプトをタグ付きで返却、ユーザーがレビュー→採用")

add_heading(doc, "成功基準", level=3)
add_bullet(doc, "1セッションで30本以上のプロンプトをタグ付きで生成")
add_bullet(doc, "ユーザーが採用するプロンプトの割合が60%以上")
add_bullet(doc, "生成時間5分以内、コスト1Workspace当たり50円以下")

# Feature F3
add_heading(doc, "F3. 引用率測定エンジン(Citation Rate Engine)", level=2)
add_callout(doc, "目的", "Chapter 4 Check段階を完全自動化。競合(Profound/Peec/Passionfruit)の最低限ラインに到達。", color="FFE0E0", title_color="C04000")

add_heading(doc, "対応LLM(MVP範囲)", level=3)
add_table(doc,
    ["LLM", "取得方法", "実装難度", "備考"],
    [
        ("Perplexity", "UIスクレイピング(Playwright)", "中", "出典URL明示で最優先実装"),
        ("ChatGPT(検索機能付き)", "OpenAI API(検索拡張)", "低", "API利用、コスト管理重要"),
        ("Google AI Overviews", "SERP API経由(SerpAPI/ScrapingBee等)", "中", "外部APIサブスクリプション"),
        ("Gemini", "UIスクレイピング(Playwright)", "中", "Google認証管理が必要"),
        ("【V2追加】Claude", "Anthropic API", "低", "検索機能未対応のため後回し"),
        ("【V2追加】Felo / Genspark / 国産LLM", "UIスクレイピング", "高", "日本語特化、差別化要素"),
    ],
    widths=[3.5, 6.0, 2.5, 7.5])

add_heading(doc, "実行スケジュールと頻度", level=3)
add_bullet(doc, "標準: 各プロンプト × 各LLM = 1日1回、深夜時間帯に分散実行")
add_bullet(doc, "高頻度プラン(Enterprise): 1日3回(朝/昼/夜)で揺らぎ吸収")
add_bullet(doc, "再試行: 失敗時は3回まで自動リトライ、それでも失敗時はエラーレコード+アラート")

add_heading(doc, "引用判定ロジック", level=3)
add_bullet(doc, "1. LLM応答から出典URL/ドメインを抽出(正規表現+構造化レスポンス解析)")
add_bullet(doc, "2. 顧客のドメインリスト(自社+サブドメイン+買収ブランド)とマッチング")
add_bullet(doc, "3. マッチ時: 引用あり(O) / 不一致時: 引用なし(X)")
add_bullet(doc, "4. 「内容類似だが出典なし」のケースは別フラグで記録(類似度スコア参考値)")
add_bullet(doc, "5. 引用順位(出典1番目/2番目/...)を併せて記録")

add_heading(doc, "Citation Rate計算式", level=3)
add_paragraph(doc, "Citation Rate (%) = (引用回数 ÷ チェック総数) × 100", size=11, bold=True, color="1F3A5F")
add_bullet(doc, "プロンプト単位/LLM単位/Workspace全体 の3粒度で算出")
add_bullet(doc, "30日移動平均でトレンド可視化")

add_heading(doc, "成功基準", level=3)
add_bullet(doc, "1日100プロンプト × 4LLM = 400実行 を稼働率99%で完遂")
add_bullet(doc, "1実行あたりLLMコスト30円以下を維持")
add_bullet(doc, "引用判定の精度: True Positive 95%以上 / False Positive 5%以下(手動検証50件で評価)")

doc.add_page_break()

# Feature F4
add_heading(doc, "F4. ダッシュボード+アラート", level=2)
add_callout(doc, "目的", "Citation Rateを「見える化」し、顧客が日次で価値を体感できる体験を提供。書籍のFAQ管理表/記録表/集計表をデジタル化。", color="FFF4E0")

add_heading(doc, "画面構成", level=3)
add_table(doc,
    ["画面", "内容", "書籍対応"],
    [
        ("Overview", "Workspace全体のCitation Rateトレンド+主要KPI(引用回数/プロンプト数/競合比較)", "─"),
        ("Prompts(FAQ管理表ビュー)", "プロンプト一覧+各プロンプトのCitation Rate+優先順位スコア", "Ch.4 FAQ管理表"),
        ("Records(FAQ記録表ビュー)", "全LLM実行ログ。確認日/AI/検索クエリ/引用?/メモ", "Ch.4 FAQ記録表"),
        ("Heatmap", "プロンプト×日付のCitation Rateヒートマップ(緑/黄/赤)", "Ch.4 ヒートマップ"),
        ("Competitors", "競合ドメインの引用シェア比較(時系列)", "─"),
        ("Reports", "月次レポート一覧+PDFダウンロード", "Ch.4 月次レポート"),
        ("Alerts", "アラート設定+履歴", "─"),
    ],
    widths=[3.5, 12.0, 4.0])

add_heading(doc, "アラート機能", level=3)
add_bullet(doc, "閾値アラート: Citation Rate 50%/60%/70% の3段階で設定可")
add_bullet(doc, "急変アラート: 前週比-10ポイント以上で自動通知")
add_bullet(doc, "競合急伸アラート: 競合の引用シェアが前月比+15ポイント以上で通知")
add_bullet(doc, "通知先: Slack(Webhook) / Email / Webhook(任意エンドポイント)")
add_bullet(doc, "アラート集約: 同一Workspace内のアラートを1日1通にまとめる(通知疲れ対策)")

add_heading(doc, "エクスポート", level=3)
add_bullet(doc, "CSV: 全データのフルエクスポート(Sheets継続利用顧客向け)")
add_bullet(doc, "PDF: 月次レポート(F5でホワイトラベル化)")
add_bullet(doc, "API: REST APIで主要データ取得可能(Looker/BigQuery連携の前段)")

add_heading(doc, "成功基準", level=3)
add_bullet(doc, "主要画面表示が3秒以内")
add_bullet(doc, "プロンプト1,000本でも快適に動作")
add_bullet(doc, "アラート配信成功率99%以上")

# Feature F5
add_heading(doc, "F5. ホワイトラベル+月次レポート", level=2)
add_callout(doc, "目的", "代理店セグメントの必須要件。クライアントレポートを代理店ブランドで提供できるようにする。", color="FFF4E0")

add_heading(doc, "ホワイトラベル設定項目(MVP)", level=3)
add_table(doc,
    ["項目", "MVP", "V2", "V3"],
    [
        ("ロゴ", "○", "─", "─"),
        ("メインカラー(1色)", "○", "─", "─"),
        ("クライアント名(ブランド)上書き", "○", "─", "─"),
        ("レポート表紙のカスタムテキスト", "○", "─", "─"),
        ("独自ドメイン(代理店.aioutmation.com)", "─", "○", "─"),
        ("送信元メールアドレス", "─", "○", "─"),
        ("完全プライベートホスティング", "─", "─", "○"),
    ],
    widths=[8.0, 2.5, 2.5, 2.5])

add_heading(doc, "月次レポート構成", level=3)
add_bullet(doc, "Page 1: 表紙(代理店ロゴ+クライアントブランド+期間)")
add_bullet(doc, "Page 2: エグゼクティブサマリー(F6のAIコメンタリー)")
add_bullet(doc, "Page 3-4: KPIダッシュボード(Citation Rate推移、引用回数、競合比較)")
add_bullet(doc, "Page 5: ヒートマップ(プロンプト別パフォーマンス)")
add_bullet(doc, "Page 6: Top/Bottomプロンプト分析")
add_bullet(doc, "Page 7: 競合詳細分析")
add_bullet(doc, "Page 8: 推奨アクション(F6 AIコメンタリー)")

add_heading(doc, "生成方式", level=3)
add_bullet(doc, "MVP: PythonベースPDF生成(reportlab/python-pptx)")
add_bullet(doc, "V2: Browserlessでのブラウザレンダリング(Web版と同一見た目)")
add_bullet(doc, "毎月1日に自動生成→アラート通知")

add_heading(doc, "成功基準", level=3)
add_bullet(doc, "代理店ロゴ反映で1分以内にレポート出力")
add_bullet(doc, "月次レポートの自動配信成功率99%")

# Feature F6
add_heading(doc, "F6. AIコメンタリー(3行サマリ)", level=2)
add_callout(doc, "目的", "経営報告/クライアント説明の作文工数を月数時間→数分に圧縮。", color="FFF4E0")

add_heading(doc, "出力フォーマット", level=3)
add_paragraph(doc, "【3行サマリ(必須)】", size=11, bold=True)
add_bullet(doc, "1行目: 全体トレンド(前月比/前期比/業界平均との位置)")
add_bullet(doc, "2行目: カテゴリ別ハイライト(製品Aは好調、製品Bは低下等)")
add_bullet(doc, "3行目: 競合差分(自社が上回ったポイント/下回ったポイント)")

add_paragraph(doc, "【詳細解説(任意)】", size=11, bold=True)
add_bullet(doc, "300-500文字のナラティブ。原因仮説と推奨アクション含む")

add_heading(doc, "生成プロンプト戦略", level=3)
add_bullet(doc, "入力: 月次集計データ(Citation Rate推移、Top/Bottomプロンプト、競合データ)")
add_bullet(doc, "テンプレート: 3パターン(改善傾向/横ばい/低下)で初期は固定、A/Bテストで改善")
add_bullet(doc, "LLM: Claude Sonnet 4.5(日本語クオリティ重視) または GPT-4o(コスト重視)")
add_bullet(doc, "コスト: 1Workspace月次1回、200円以下")

add_heading(doc, "成功基準", level=3)
add_bullet(doc, "顧客が「コメントをそのまま提出可能」と評価する割合70%以上")
add_bullet(doc, "ファクトエラー率5%以下(数値の取り違え・誤った因果推論)")

doc.add_page_break()

# === 4. Technical Architecture ===
add_heading(doc, "4. 技術アーキテクチャ", level=1)

add_heading(doc, "4.1 推奨技術スタック", level=2)
add_table(doc,
    ["レイヤー", "推奨", "代替案", "選定理由"],
    [
        ("フロントエンド", "Next.js 14 (App Router) + TypeScript + Tailwind + shadcn/ui",
         "Remix / Vite + React",
         "日本のSaaS開発で最も人材確保しやすい、Vercelでの即時デプロイ、ストリーミングUI(LLM応答待ち)対応"),
        ("バックエンド API", "Node.js (NestJS) + TypeScript",
         "Python (FastAPI), Go (Echo)",
         "フロントとの型共有、LLM SDK公式サポート充実、TS人材豊富"),
        ("LLM呼び出し基盤", "Vercel AI SDK / LangChain.js",
         "LiteLLM(マルチプロバイダ抽象化)",
         "OpenAI/Anthropic/Google切替容易、ストリーミング標準対応"),
        ("UIスクレイピング", "Playwright(Containerized) on AWS Fargate",
         "Browserless.io(SaaS), Puppeteer",
         "Perplexity/Geminiなどブラウザ操作必須、コスト最適化"),
        ("データベース", "PostgreSQL 16 + Prisma ORM",
         "MySQL, MongoDB",
         "JSON型サポート、スケーラビリティ、運用人材豊富"),
        ("キャッシュ/キュー", "Redis (ElastiCache) + BullMQ",
         "AWS SQS, Cloud Tasks",
         "バックグラウンドジョブ(LLM実行)に最適、観測しやすい"),
        ("認証", "Clerk または Auth0",
         "自前実装",
         "MVPでは買って時間節約、SOC 2対応も容易"),
        ("ホスティング", "AWS (ECS Fargate + RDS + ElastiCache + S3 + CloudFront)",
         "GCP, Vercel + Supabase",
         "エンタープライズ顧客のセキュリティ要件に対応しやすい"),
        ("監視", "Datadog または Sentry + Mixpanel",
         "New Relic, OpenTelemetry自前",
         "LLM API監視・コスト追跡が容易"),
        ("ファイル/データレイク", "S3 + Athena (BigQueryも可)",
         "GCS + BigQuery",
         "業界ベンチマーク用データ蓄積基盤(V2準備)"),
    ],
    widths=[3.5, 5.5, 4.0, 6.5])

add_heading(doc, "4.2 アーキテクチャ概要図(テキスト版)", level=2)
add_paragraph(doc, """
[ユーザー] → [Vercel CDN] → [Next.js Frontend]
                                   ↓ (API)
                              [NestJS API Gateway]
                                   ↓
            ┌──────────────────┼──────────────────┐
            ↓                   ↓                   ↓
       [Auth Service]      [Core API]        [LLM Worker Pool]
            ↓                   ↓                   ↓
            └────────→ [PostgreSQL] ←───────[Redis Queue (BullMQ)]
                            ↑                       ↓
                       [BackOffice]          [Playwright Containers]
                                                    ↓
                                          [LLM API Providers]
                                          (OpenAI/Anthropic/Google/SerpAPI)
                                                    ↓
                                          [LLM UI Sites]
                                          (Perplexity/Gemini/...)
""", size=9, color="333333")

add_heading(doc, "4.3 LLM実行ワーカーの設計ポイント", level=2)
add_bullet(doc, "並列度制御: 1Workspaceあたり同時実行3、全体で50並列上限(LLMレート制限対応)")
add_bullet(doc, "リトライ: 指数バックオフ(2s/4s/8s)で3回")
add_bullet(doc, "コスト保護: Workspaceごとの月間コスト上限を設定、超過時にアラート+一時停止")
add_bullet(doc, "プロキシローテーション: UIスクレイピング時のIP多様化(住宅プロキシ含む)")
add_bullet(doc, "セッション管理: 各LLMサイトのログイン状態をRedisで保持、定期的に再認証")
add_bullet(doc, "レスポンスキャッシュ: 24時間以内の同一プロンプトはキャッシュ返却")

add_heading(doc, "4.4 セキュリティ要件", level=2)
add_bullet(doc, "通信: TLS 1.3, HSTS, CSP")
add_bullet(doc, "保存: PostgreSQL暗号化(AES-256), S3 SSE-KMS")
add_bullet(doc, "シークレット管理: AWS Secrets Manager")
add_bullet(doc, "テナント分離: Row-Level Security(PostgreSQL) + Application層チェック2重")
add_bullet(doc, "監査: 認証ログ最低90日保持(エンタープライズはV2で1年)")
add_bullet(doc, "データ削除: 顧客解約時、30日以内に全データ完全削除(GDPR/個人情報保護法対応)")

doc.add_page_break()

# === 5. Data Model ===
add_heading(doc, "5. データモデル概要", level=1)

add_heading(doc, "5.1 主要エンティティ", level=2)
add_table(doc,
    ["エンティティ", "主要属性", "関係"],
    [
        ("Account", "id, name, plan, billingEmail, createdAt", "1 Account → N Workspace"),
        ("Workspace", "id, accountId, name, brandName, primaryColor, logoUrl, customDomain(V2), settings(JSON)", "1 Workspace → N Project"),
        ("Project", "id, workspaceId, name, archivedAt", "1 Project → N Brand, N Prompt"),
        ("User", "id, email, name, accountId, defaultRole", "M User × N Workspace (UserWorkspaceRole)"),
        ("UserWorkspaceRole", "userId, workspaceId, role(Owner/Admin/Editor/Viewer)", "結合テーブル"),
        ("Brand", "id, projectId, name, primaryDomain, additionalDomains(JSON), keywords(JSON)", "ターゲットブランド情報"),
        ("Competitor", "id, projectId, name, primaryDomain, additionalDomains(JSON)", "競合ブランド情報"),
        ("Prompt", "id, projectId, text, tags(JSON), priorityScore, contextNote, createdAt, archivedAt", "プロンプト本体"),
        ("LLMRun", "id, promptId, llmProvider, executedAt, status, responseText, sourceUrls(JSON), costJpy, errorMessage", "LLM実行ログ(膨大)"),
        ("CitationRecord", "id, llmRunId, promptId, brandId, citedFlag, citationPosition, similarityScore, createdAt", "引用判定結果"),
        ("Page", "id, projectId, url, title, h1, lastCrawledAt, structuredDataPresent(bool), aeoScore", "ターゲットページ(対象ブランドの自社ページ)"),
        ("AlertRule", "id, workspaceId, name, conditions(JSON), notificationChannels(JSON), enabled", "アラート設定"),
        ("AlertEvent", "id, alertRuleId, triggeredAt, payload(JSON), acknowledged", "アラート発生履歴"),
        ("Report", "id, workspaceId, periodStart, periodEnd, generatedAt, pdfUrl, summary(JSON)", "月次レポート"),
        ("Commentary", "id, reportId, generatedText, model, costJpy, createdAt", "AIコメンタリー"),
    ],
    widths=[4.0, 11.0, 4.5])

add_heading(doc, "5.2 想定データ量(リリース1年後)", level=2)
add_table(doc,
    ["エンティティ", "想定件数", "備考"],
    [
        ("Account", "100", "エンプラ50+代理店30+SMB200弱"),
        ("Workspace", "500", "代理店1社あたり5-10、エンプラ1社=1Workspace"),
        ("Project", "1,000", "1Workspace平均2Project"),
        ("Prompt", "200,000", "1Project平均200本"),
        ("LLMRun", "292,000,000/年", "200K × 4LLM × 365日。partition必須"),
        ("CitationRecord", "292,000,000/年", "LLMRunと1:1。partition必須"),
        ("Page", "100,000", "1Project平均100ページ"),
        ("Report", "12,000/年", "1Workspace月次"),
    ],
    widths=[5.0, 4.0, 10.5])

add_heading(doc, "5.3 パーティショニング戦略", level=2)
add_bullet(doc, "LLMRun / CitationRecord: 月次パーティション(2026-05, 2026-06, ...)")
add_bullet(doc, "古いパーティションは1年後にS3へアーカイブ→Athenaクエリ可能に")
add_bullet(doc, "AlertEvent: 90日でアーカイブ")

doc.add_page_break()

# === 6. Sprint Plan ===
add_heading(doc, "6. 開発計画(24週間)", level=1)

add_heading(doc, "6.1 スプリント概要", level=2)
add_table(doc,
    ["Sprint", "週", "主要成果物", "達成基準"],
    [
        ("Sprint 0", "W0-2", "アーキ確定、DB設計、認証基盤、開発環境構築", "Hello World API + 認証動作"),
        ("Sprint 1", "W3-4", "F1: Account/Workspace/User/RBAC + UI骨格", "代理店アカウントで5Workspace作成可能"),
        ("Sprint 2", "W5-6", "F2: プロンプト管理(CRUD+CSVインポート+タグ)", "1Workspaceに100プロンプト登録"),
        ("Sprint 3", "W7-8", "F2: AI自動生成 + F3基盤(LLM Workerプール)", "プロンプト30本自動生成 + Perplexity 1日1回実行"),
        ("Sprint 4", "W9-10", "F3: ChatGPT/Gemini/AI Overviews対応 + 引用判定エンジン", "4LLM × 100プロンプト = 400/日 稼働"),
        ("Sprint 5", "W11-12", "F4: ダッシュボード(Overview/Prompts/Records/Heatmap)", "Citation Rate可視化、3秒以内表示"),
        ("Sprint 6", "W13-14", "F4: アラート機能(Slack/Email)", "閾値+急変+競合急伸の3種アラート稼働"),
        ("Sprint 7", "W15-16", "F5: ホワイトラベル(ロゴ・カラー・ブランド名)", "代理店ブランドでレポート出力"),
        ("Sprint 8", "W17-18", "F5: 月次レポート自動生成 + F6: AIコメンタリー", "月次PDF自動配信 + 3行サマリ生成"),
        ("Sprint 9", "W19-20", "βリリース準備(オンボーディング、ドキュメント、サポート体制)", "デザインパートナー3-5社にオンボーディング"),
        ("Sprint 10", "W21-22", "βフィードバック反映 + バグ修正 + パフォーマンス改善", "重大バグゼロ、稼働率99%"),
        ("Sprint 11-12", "W23-24", "GAリリース準備(プライシング、契約、課金、マーケ)", "商用ローンチ"),
    ],
    widths=[2.0, 1.5, 11.0, 4.0])

add_heading(doc, "6.2 マイルストーン", level=2)
add_table(doc,
    ["マイルストーン", "達成時期", "意義"],
    [
        ("M1: アーキ凍結", "W2末", "技術的負債の予防"),
        ("M2: クローズドα", "W10末", "社内+1社で計測エンジンの実証"),
        ("M3: オープンβ", "W18末", "デザインパートナー3-5社"),
        ("M4: GA(商用ローンチ)", "W24末", "課金開始、PR、セールス本格化"),
        ("M5: PMF確認", "W36末(GA+12週)", "NRR 100%以上、有料転換成立"),
    ],
    widths=[5.0, 3.5, 11.0])

add_heading(doc, "6.3 想定開発体制", level=2)
add_bullet(doc, "プロダクトマネージャー: 1名(委託先ではなく弊社側)")
add_bullet(doc, "テックリード/アーキテクト: 1名")
add_bullet(doc, "バックエンドエンジニア: 2-3名")
add_bullet(doc, "フロントエンドエンジニア: 1-2名")
add_bullet(doc, "SRE/DevOps: 0.5名(兼任)")
add_bullet(doc, "QA: 0.5名(後半から1名)")
add_bullet(doc, "デザイナー: 0.5名")
add_bullet(doc, "合計: 5-7名 / 24週間 = 約120-170人週")

doc.add_page_break()

# === 7. Discussion Points ===
add_heading(doc, "7. 委託先との論点(意思決定が必要な事項)", level=1)

add_heading(doc, "7.1 Build vs Buy 判断", level=2)
add_table(doc,
    ["項目", "選択肢", "推奨", "コスト想定"],
    [
        ("認証基盤", "Build / Clerk / Auth0", "Clerk(MVPまで)→Auth0(エンプラ拡大時)", "Clerk: $99/月〜"),
        ("LLMマルチプロバイダ抽象化", "Build / LangChain / LiteLLM / Vercel AI SDK", "Vercel AI SDK", "OSS"),
        ("ジョブキュー", "BullMQ / SQS / Inngest", "BullMQ(Redis)", "Redis稼働費のみ"),
        ("PDF生成", "Playwright / reportlab / Pandoc / Browserless", "MVP: reportlab、V2: Browserless", "MVPは無料、V2は$200/月〜"),
        ("UIスクレイピング基盤", "自前Playwright / Browserless / Bright Data", "MVP自前、スケール時にBrowserless検討", "自前なら運用工数のみ"),
        ("SerpAPI(Google AI Overviews)", "SerpAPI / ScrapingBee / 自前", "SerpAPI", "$50/月〜(プロンプト数依存)"),
        ("監視", "Datadog / New Relic / Sentry+自作", "Sentry + Mixpanel(MVP)→Datadog(GA)", "Sentry: $26/月〜"),
        ("分析DB(将来)", "BigQuery / Athena / Snowflake", "MVPは不要、V2でAthena", "─"),
    ],
    widths=[5.0, 4.5, 5.5, 4.5])

add_heading(doc, "7.2 LLM API選定", level=2)
add_table(doc,
    ["用途", "推奨LLM", "代替", "コスト想定"],
    [
        ("プロンプト自動生成", "Claude Sonnet 4.5", "GPT-4o", "$3/1M tokens (input)"),
        ("引用判定(自社ドメインマッチ)", "ロジック実装で十分(LLM不要)", "─", "─"),
        ("AIコメンタリー", "Claude Sonnet 4.5(日本語品質重視)", "GPT-4o", "$3/1M tokens"),
        ("ChatGPT検索クエリ実行", "OpenAI GPT-4o w/ search", "─", "$5/1M tokens"),
        ("Perplexity実行", "UIスクレイピング(LLM不要)", "Perplexity API", "API: $5/1K queries"),
        ("Gemini実行", "UIスクレイピング", "Gemini API", "API: $1/1M tokens"),
    ],
    widths=[5.0, 5.5, 4.0, 5.0])

add_paragraph(doc, "1顧客あたり月間LLMコスト想定: 5,000〜30,000円(プロンプト50-500本、4LLM、月次レポート)→ プラン価格設計の根拠", size=10, bold=True, color="C04000")

add_heading(doc, "7.3 オープン論点(委託先と要相談)", level=2)
add_bullet(doc, "Q1: UIスクレイピングの法的リスク評価(Perplexity/Geminiの利用規約レビュー必要)")
add_bullet(doc, "Q2: 顧客のプロンプト・引用データの法的位置付け(規約整備)")
add_bullet(doc, "Q3: LLM API値上げ時の価格転嫁ポリシー")
add_bullet(doc, "Q4: 障害発生時のSLA(エンプラ顧客向け99.9%が現実的か)")
add_bullet(doc, "Q5: バックアップ・DR(Disaster Recovery)体制")
add_bullet(doc, "Q6: 個人情報保護法対応(顧客のCRMデータ連携時)")
add_bullet(doc, "Q7: SOC 2 Type II取得タイミング(MVP後V2想定で1,500-3,000万円)")
add_bullet(doc, "Q8: マルチリージョン展開(将来の海外展開で必要か議論)")

add_heading(doc, "7.4 委託契約のスコープ案", level=2)
add_bullet(doc, "成果物ベース: F1〜F6完成+βリリースまで一括 → リスク分担曖昧、変更管理困難")
add_bullet(doc, "準委任ベース(推奨): スプリント単位での発注、PdMの仕様判断は弊社、技術判断は委託先")
add_bullet(doc, "ハイブリッド: コア機能(F3, F4)は成果物、その他は準委任")
add_paragraph(doc, "推奨: 準委任(アジャイル開発)。ペイン解像度・市場フィードバック・LLM技術進化への適応性を優先。", size=11, bold=True, color="2E7A4E")

doc.add_page_break()

# === 8. Risks ===
add_heading(doc, "8. リスクと対策", level=1)

add_table(doc,
    ["#", "リスク", "影響度", "発生可能性", "対策"],
    [
        ("R1", "LLM API値上げでマージン悪化", "高", "高",
         "・複数LLMマルチプロバイダ前提で設計\n・プラン価格にバッファ20%含める\n・コスト上限機能で顧客を制御"),
        ("R2", "Perplexity/Gemini等のUIスクレイピング拒否", "高", "中",
         "・公式API出る都度移行\n・複数LLMで分散(1社依存しない)\n・住宅プロキシ等の対策"),
        ("R3", "競合(Profound等)が日本市場に本格参入", "高", "中",
         "・日本語LLM対応で先行(Felo/Genspark等)\n・代理店パートナー網を早期に構築\n・瀧内氏連携で著者監修ブランド確立"),
        ("R4", "開発遅延(24週→36週)", "中", "中",
         "・Sprint 0-4を最重要、後半は柔軟に\n・F6(AIコメンタリー)等は最悪V2へ延期可\n・隔週デモで進捗可視化"),
        ("R5", "βパートナー獲得失敗", "高", "中",
         "・並行で5-10社にアプローチ\n・初期は無料運用で実績作り\n・代理店経由のリード獲得"),
        ("R6", "引用判定の精度問題(False Positive多発)", "中", "中",
         "・手動検証50件でリリース判定\n・お客様にもフラグ修正UIを提供\n・継続的な精度改善"),
        ("R7", "代理店向け要件の見落とし", "中", "中",
         "・Sprint 1-2で代理店PMにヒアリング\n・βで代理店2-3社を必ず含める"),
        ("R8", "個人情報・データ取扱い違反", "極高", "低",
         "・初期から法務レビュー\n・データ最小化原則\n・規約・プライバシーポリシー整備"),
    ],
    widths=[1.0, 5.0, 1.5, 2.0, 9.0])

# === 9. Decision Asks ===
add_heading(doc, "9. 委託先との初回ディスカッション議題(チェックリスト)", level=1)

add_paragraph(doc, "下記項目を初回〜2回目のミーティングで合意することを推奨します:", size=10)
add_bullet(doc, "□ 開発チーム体制と役割(PdM/テックリード/エンジニア/QA/デザイナー)")
add_bullet(doc, "□ 開発手法(スクラム/カンバン)とスプリント長さ(2週間推奨)")
add_bullet(doc, "□ 契約形態(準委任/成果物ベース/ハイブリッド)")
add_bullet(doc, "□ 技術スタックの最終確定(本資料4.1の表をたたき台に)")
add_bullet(doc, "□ ブランチ戦略(Git Flow / Trunk-based / GitHub Flow)")
add_bullet(doc, "□ コードレビュー基準とCI/CDパイプライン")
add_bullet(doc, "□ テスト戦略(ユニット/統合/E2E)とカバレッジ目標")
add_bullet(doc, "□ 監視・アラート設計の責任範囲")
add_bullet(doc, "□ デプロイ環境(dev/staging/production)")
add_bullet(doc, "□ 委託先に開示する情報の範囲(顧客データ、商談情報等)")
add_bullet(doc, "□ 知的財産権の取扱い(コード、AIプロンプトテンプレート、業界ベンチマークデータ)")
add_bullet(doc, "□ 成果物の納品形式とドキュメント要件")
add_bullet(doc, "□ 障害対応のエスカレーションパス")
add_bullet(doc, "□ 月次振り返り・四半期戦略レビューの頻度")

# === Appendix ===
add_heading(doc, "Appendix: 関連リソース", level=1)
add_bullet(doc, "AIO_Textbook_Summary_50p.docx / AIO_Textbook_Summary_10p.docx (アウトプット①)")
add_bullet(doc, "AIOutmation_LeanCanvas.docx (アウトプット②)")
add_bullet(doc, "CustomerJourneyMap_GEO_AEO.docx/.pptx (顧客ペインのジャーニーマップ)")
add_bullet(doc, "PainResolutionAnalysis_GEO_AEO.pptx (ベンチマーク3社の解決度分析)")
add_bullet(doc, "出典: 瀧内賢『これからはじめるAIO AI最適化の教科書』技術評論社, 2026年3月発行")

add_paragraph(doc, "", size=8)
add_paragraph(doc, "本ドキュメントは委託先との合意形成を目的とした初版です。技術的判断は委託先のアーキテクトと協議の上、確定してください。",
              size=9, color="888888")

doc.save("/home/user/AIO/AIOutmation_MVP_Spec.docx")
print("Saved: AIOutmation_MVP_Spec.docx")
