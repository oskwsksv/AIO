"""Generate a printable Word document of the GEO/AEO Customer Journey Maps."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_ORIENT
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


def set_cell_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "999999")
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def style_run(run, size=10, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.name = "Yu Gothic"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), "Yu Gothic")
    rFonts.set(qn("w:ascii"), "Yu Gothic")
    rFonts.set(qn("w:hAnsi"), "Yu Gothic")
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def write_cell(cell, text, size=9, bold=False, bg=None, color=None):
    cell.text = ""
    set_cell_borders(cell)
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
        run = p.add_run(line)
        style_run(run, size=size, bold=bold, color=color)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    sizes = {0: 18, 1: 15, 2: 13, 3: 11}
    colors = {0: "1F3A5F", 1: "1F3A5F", 2: "2E5C8A", 3: "444444"}
    style_run(run, size=sizes.get(level, 11), bold=True, color=colors.get(level, "000000"))


def add_paragraph(doc, text, size=10, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    style_run(run, size=size, bold=bold, color=color)


def add_persona_table(doc, rows):
    table = doc.add_table(rows=len(rows), cols=2)
    table.autofit = False
    widths = [Cm(4.5), Cm(20)]
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]
    for i, (k, v) in enumerate(rows):
        write_cell(table.cell(i, 0), k, size=9, bold=True, bg="E8EEF7")
        write_cell(table.cell(i, 1), v, size=9)


# Customer Journey Map data structure
ASPECT_LABELS = ["タスク／ジョブ", "行動", "思考・問い", "感情", "🔥 ペイン（コア）", "タッチポイント", "🎯 機会（プロダクト示唆）"]
ASPECT_COLORS = ["F5F5F5", "F5F5F5", "F5F5F5", "F5F5F5", "FFF0F0", "F5F5F5", "F0F8E8"]


def add_phase_table(doc, phase_title, sub_headers, sub_data):
    """sub_data: list of dict keyed by aspect label, length matches sub_headers"""
    add_heading(doc, phase_title, level=2)

    n_subs = len(sub_headers)
    table = doc.add_table(rows=len(ASPECT_LABELS) + 1, cols=n_subs + 1)
    table.autofit = False

    label_w = Cm(3.0)
    sub_w = Cm((24.5 - 3.0) / n_subs)
    for row in table.rows:
        row.cells[0].width = label_w
        for i in range(1, n_subs + 1):
            row.cells[i].width = sub_w

    write_cell(table.cell(0, 0), "観点", size=10, bold=True, bg="1F3A5F", color="FFFFFF")
    for j, h in enumerate(sub_headers):
        write_cell(table.cell(0, j + 1), h, size=10, bold=True, bg="2E5C8A", color="FFFFFF")

    for i, label in enumerate(ASPECT_LABELS):
        bg = ASPECT_COLORS[i]
        write_cell(table.cell(i + 1, 0), label, size=9, bold=True, bg="E8EEF7")
        for j in range(n_subs):
            content = sub_data[j].get(label, "")
            write_cell(table.cell(i + 1, j + 1), content, size=9, bg=bg)


# ====== DATA ======

PERSONA_A_INFO = [
    ("想定ロール", "Senior Manager, Digital Marketing(GEO/AEO Lead) 30代後半"),
    ("所属", "ARR $100M〜$1B のグローバルB2B SaaS企業(例:MongoDB / DocuSign規模)"),
    ("上司ライン", "Director → VP Marketing → CMO"),
    ("個人KGI", "パイプライン貢献、ブランドVisibility"),
    ("個人KPI", "AI Mention Rate / Share of Voice / Sentiment / AI経由トラフィック・パイプライン"),
    ("チーム", "自分1〜3名(コンテンツストラテジスト / スキーマ専門 / データアナリスト)"),
    ("周辺ステークホルダー", "SEOチーム、コンテンツチーム、PR、Web開発、法務、情シス／セキュリティ、財務"),
    ("既存ツール", "GA4、Search Console、Semrush/Ahrefs、HubSpot、Salesforce、Segment、Looker"),
]

PERSONA_B_INFO = [
    ("想定ロール", "Account Director / GEO Lead(兼 SEOマネージャ) 30代前半〜後半"),
    ("所属", "50〜300名規模のデジタル運用代理店(パフォーマンス・SEO・コンテンツを提供)"),
    ("上司ライン", "Group Director → VP / Partner"),
    ("個人KGI", "クライアントリテンション、新規受注、リテイナー単価"),
    ("個人KPI", "アカウント数、月次リテイナー、解約率、NPS、提案勝率"),
    ("チーム", "アカウント担当・SEOアナリスト・コンテンツライター(外部含む)"),
    ("担当クライアント", "大手B2B SaaS 5〜12社、月額$3k〜$30kリテイナー"),
    ("既存ツール", "Semrush / Ahrefs / GA4 / Looker Studio / Google Workspace / Slack / Asana"),
]

# --- Persona A overview phases (1-4 and 9 from earlier map) ---
PERSONA_A_OVERVIEW_PHASES = [
    ("① 異変認知", "オーガニック流入の減少傾向を察知。CMOから「ChatGPTで自社が出てこない」と一言。\n\n【ペイン】既存ツール(GA/Semrush)でAI流入が見えない／社内に前例ゼロで何から手を付けるか不明／CMOから個別質問を受けても即答できない"),
    ("② 情報収集・学習", "GEO/AEO/LLMOの定義整理、競合動向、業界レポート読み込み。\n\n【ペイン】情報が断片的で各社ベンダー発信に偏る／日本語LLM対応の情報が極端に少ない／意思決定者向け一枚絵を短時間で作る必要"),
    ("③ 戦略立案・KPI設計", "自社ICP×購買ファネルでプロンプト300本を選定。北極星KPI決定。\n\n【ペイン】プロンプト300本の優先順位付けに正解がない／どのKPIをCMO提出指標に置くか合意が取れない／既存SEO KPIとの連携設計が難しい"),
    ("④ ツール選定・稟議", "3〜5社をPoC比較、セキュリティ審査、年間予算化。\n\n【ペイン】SOC2/SSO/監査ログ等のエンタープライズ要件を満たすベンダーが限られる／Profoundは高額、安いツールはセキュリティ未充足／購買・法務・情シスの三段審査で4〜6ヶ月"),
    ("⑨ 横展開・スケール", "多言語・多リージョン展開、サブブランド追加、社内ナレッジ化。\n\n【ペイン】ツールが多言語・多リージョン本格対応していない／グローバルでSOPを統一できない／社内人材育成・引継ぎ資料作成が個人依存"),
]

# --- Persona A: Phase ⑤ ---
P_A_5 = [
    {
        "タスク／ジョブ": "ブランド／サブブランド／競合／対象モデル／対象国・言語／SSO／RBAC／監査ログ／API・BI連携を設定",
        "行動": "CSMキックオフ、競合リスト整備、Salesforce/HubSpot/GA4/Looker接続設定、社内SSO発行、IP範囲共有",
        "思考・問い": "「セキュリティ部門の追加質問が増えそう」\n「他システム連携を最初に詰めないと後で詰む」",
        "感情": "😅 設定地獄",
        "🔥 ペイン（コア）": "・各種SaaS連携の認証・権限取得に部門横断で2〜4週間\n・サブブランド／買収企業／グローバル各リージョンの設定が手作業の繰り返し\n・SOC2/RBAC等の運用ルール文書化に時間",
        "タッチポイント": "CSM Slack / 管理コンソール / 情シス・購買",
        "🎯 機会（プロダクト示唆）": "連携テンプレートライブラリ(Salesforce/HubSpot/GA4/Looker)＋サブブランド一括コピー",
    },
    {
        "タスク／ジョブ": "ICP×購買ファネル×製品ライン×ペルソナでプロンプト300〜1,000本を起案・選定・タグ付け",
        "行動": "プロダクトマーケ・SE・営業・ABMチームへヒアリング、競合プロンプト調査、PMM資料・営業トーク・Win/Lossから抽出",
        "思考・問い": "「プロンプトの粒度・本数の正解値は？」\n「日本語と英語を分けて持つべき？」\n「ABM向けの社名×プロダクト系まで含めるか」",
        "感情": "🤔 正解のなさへの不安・属人性",
        "🔥 ペイン（コア）": "・プロンプト選定の正解SOPがない(ベンダー提供テンプレが業界・自社に合わない)\n・1,000本規模になると重複・ファネル偏り・タグ運用が破綻\n・営業現場の生プロンプト(実際の商談で出る質問)を吸い上げる仕組みがない",
        "タッチポイント": "プロンプト管理画面 / Notion / 営業ヒアリングMTG",
        "🎯 機会（プロダクト示唆）": "プロンプトAI自動生成(自社ドメイン解析×競合×ICPから提案)＋営業データ自動取込(Gong/Salesforce連携で実プロンプト抽出)",
    },
    {
        "タスク／ジョブ": "プロンプト初回実行→ベースライン記録、ダッシュボード分割、アラート閾値・受信者・チャネル設計、QBR/レビュー会議体設定",
        "行動": "初回スナップショット取得、Looker埋め込み、Slackチャネル作成、Sentiment急落・SoV低下のアラート設計",
        "思考・問い": "「アラート過多になると無視される」\n「閾値を厳しくしすぎると見逃す」\n「経営に出すKPIから逆算してダッシュボードを切るべき」",
        "感情": "🙂 進捗感／🔔 通知設計の悩み",
        "🔥 ペイン（コア）": "・初期データのノイズ・モデル変動を「異常値」として誤検知しがち\n・経営報告KPIに合わせたダッシュボードを後付けで作り直す\n・アラート設計の正解値(閾値・頻度)が経験則に依存",
        "タッチポイント": "ダッシュボード / Slackアラート / Looker / カレンダー(QBR)",
        "🎯 機会（プロダクト示唆）": "アラート閾値の自動推奨(ベースライン分散から提案)、経営KPIテンプレート(CMO向けカスタマイズ済)",
    },
]

# --- Persona A: Phase ⑥ ---
P_A_6 = [
    {
        "タスク／ジョブ": "朝のSlackアラート確認、緊急性判定、関係部門にエスカレーション、誤検知の振り分け",
        "行動": "通知巡回、PR部門・営業部門・SEOチームへのエスカレ、「無視」「監視継続」「即対応」の三択判断",
        "思考・問い": "「これはノイズか実シグナルか」\n「PRに即連絡すべき案件か」",
        "感情": "😵 通知疲れ・即応プレッシャー",
        "🔥 ペイン（コア）": "・毎日数十件のアラートで本質的なシグナルが埋もれる\n・LLMの確率的揺らぎを「変化」と誤検知する頻度が高い\n・PR/営業の各部門との連絡経路が個人Slackに依存",
        "タッチポイント": "Slack / メール / 管理コンソール",
        "🎯 機会（プロダクト示唆）": "シグナル要約Agent(「今朝の本物の異変は2件」「他は確率変動」と判定)＋ノイズ自動フィルタ",
    },
    {
        "タスク／ジョブ": "Top10プロンプトのMention Rate / Citation Rate / SoV / First-Mention Rateの推移を確認、要因分析",
        "行動": "Looker/GEOツールで週次トレンドレポート確認、競合の特定ページ調査、社内Slackで共有",
        "思考・問い": "「上昇／低下の主要因は何か」\n「Claude/Geminiで差が出ている理由は？」",
        "感情": "🧐 分析モード",
        "🔥 ペイン（コア）": "・週次レポート作成に毎週3〜5時間手作業で消費\n・モデル間の差分(ChatGPTは上がってGeminiは下がる等)の解釈に正解がない\n・SEO KPIとAEO KPIの統合ビューが存在せず別ツール併用",
        "タッチポイント": "Looker / Notion / 週次レビューMTG",
        "🎯 機会（プロダクト示唆）": "週次サマリ自動生成(要因の差分言語化)＋SEO/AEO統合ビュー",
    },
    {
        "タスク／ジョブ": "センチメント悪化／競合急伸／自社引用喪失をクラスタリング、月次プランへ反映",
        "行動": "センチメント引用元のソース確認、競合の被引用ページ分析、四半期計画にリスト化",
        "思考・問い": "「センチメント悪化の根源はどこか」\n「競合のどのコンテンツが引用されている？」",
        "感情": "🔥 戦略思考／😬 競合焦り",
        "🔥 ペイン（コア）": "・センチメント悪化の根本原因(引用元ページ・ニュース・口コミ)まで辿るのに半日〜1日\n・競合のページ単位citation分析が手作業\n・四半期計画とのつなぎ込みが属人的",
        "タッチポイント": "センチメントダッシュボード / 競合ページ / 社内Confluence",
        "🎯 機会（プロダクト示唆）": "根本原因AI(引用元PDFやニュースまで辿って要因提示)＋競合被引用ページ自動アーカイブ",
    },
]

# --- Persona A: Phase ⑦ ---
P_A_7 = [
    {
        "タスク／ジョブ": "プロンプトギャップ・トピックギャップ・引用先ページ分析、ROIが大きい順に改修候補をリスト化",
        "行動": "ツールで「未引用だが商機あり」プロンプトを抽出、競合被引用ページを構造分解、Jiraチケット起票",
        "思考・問い": "「どこから手を付ければ最速で結果が出る？」\n「コンテンツ施策とPR施策のどちらが効く？」",
        "感情": "🧠 戦略集中",
        "🔥 ペイン（コア）": "・「未引用だが商機あり」の機会を手作業で抽出しなければならない\n・コンテンツ vs PR vs 構造化のどれが効くか判断材料が少ない\n・JiraやAsanaへの起票が分断・属人化",
        "タッチポイント": "プロンプトギャップ画面 / Jira / Notion",
        "🎯 機会（プロダクト示唆）": "ギャップ→Jira/Asana自動起票＋ROI予測スコア(改修コスト×期待引用増)",
    },
    {
        "タスク／ジョブ": "H2/H3の質問化、回答サマリ冒頭挿入、原子段落化、FAQ追加、Article/Product/FAQPage Schema実装、E-E-A-T強化(著者情報・引用源)",
        "行動": "コンテンツライターへのブリーフ作成、Web開発へSchema実装依頼、社内法務レビュー、PR/PMM連携",
        "思考・問い": "「LLMに引用される構造とは具体的に何か」\n「既存ブランドガイドラインと衝突しないか」",
        "感情": "😤 部門横断調整の苦労",
        "🔥 ペイン（コア）": "・コンテンツ・Web開発・法務・PMMの4部門横断調整で1施策あたり2〜6週間\n・スキーマやFAQ実装はWeb開発のスプリントに乗せられないことが多い\n・ライターがAEO要件を理解しておらず再修正発生\n・著者情報・E-E-A-T準備が想像以上に重い",
        "タッチポイント": "CMS / Figma / Web開発スプリント / コンテンツ会議",
        "🎯 機会（プロダクト示唆）": "CMSへの改修PR自動生成(Schema・FAQ・原子段落構造)＋AEO対応ライティングAI＋著者プロフィール自動整備",
    },
    {
        "タスク／ジョブ": "CMS公開、サイトマップ送信、再クロール促進、4〜8週間で再計測しBefore/Afterを記録",
        "行動": "デプロイ→GSC再クロール→再プロンプト実行→ダッシュボードで前後比較→Win/Lossをドキュメント化",
        "思考・問い": "「ラグはどれぐらい？」\n「再計測タイミングは？」\n「成功施策の再現性は？」",
        "感情": "😌 達成感／📈 次の打ち手探し",
        "🔥 ペイン（コア）": "・効果反映に4〜8週間のラグでPDCAが遅い\n・LLMキャッシュ・学習周期が読めず再計測タイミング属人化\n・成功施策のナレッジ化・横展開が個人依存\n・改修したのに引用されず原因不明、というケースの分析困難",
        "タッチポイント": "GSC / GEOダッシュボード / Slack共有",
        "🎯 機会（プロダクト示唆）": "再計測自動スケジューリング＋Before/After自動レポート＋成功事例ナレッジAI(再現可能な型を抽出)",
    },
]

# --- Persona A: Phase ⑧ ---
P_A_8 = [
    {
        "タスク／ジョブ": "GEOツール×GA4×CRM(Salesforce)×コンテンツKPIを統合、経営向け3行サマリ生成",
        "行動": "各ツールエクスポート、Looker/スプレッドシート集約、GEO→AI流入→MQL→SQL→パイプラインのファネル接続、コメント記述",
        "思考・問い": "「経営は400行のプロンプトログを見ない、3行で要約が必要」\n「AI Visibility → 売上の因果をどう示すか」",
        "感情": "😬 集約地獄",
        "🔥 ペイン（コア）": "・ツールが分断しており、データ統合に毎月8〜20時間消費\n・GEO→売上の因果推論手法が確立していない\n・指標粒度が経営×実務でずれる(経営は3行、実務は400行)",
        "タッチポイント": "Looker / GEOツール / スプレッドシート / Salesforce",
        "🎯 機会（プロダクト示唆）": "CMO向け3行サマリ自動生成(前月比/カテゴリ/競合差分)＋AI Visibility→Pipelineアトリビューション",
    },
    {
        "タスク／ジョブ": "QBR用デック作成(Visibility Radar / Heatmap / Stack Ranking / 競合差分 / ROI / 次四半期計画)",
        "行動": "スライド構成(Executive Summary / Goals vs Results / Funnel Health / SoV / 競合 / 次期計画)、PMOチェック",
        "思考・問い": "「Visibility Radarは見栄えはいいが、説得力は弱い」\n「ROIストーリーをどう組み立てるか」",
        "感情": "🧠 ストーリーテリング集中",
        "🔥 ペイン（コア）": "・スライド作成に毎月数日〜1週間\n・ROI物語の型がなく毎回ゼロから構成\n・競合比較データの鮮度が古い(手作業更新のため)\n・他チャネル(広告、PR、SEO)との寄与按分が困難",
        "タッチポイント": "PowerPoint/Keynote / Figma / 社内レビュー",
        "🎯 機会（プロダクト示唆）": "QBRデック自動生成(Profound・Conductorの先行事例を参照可能なテンプレ)＋ROI物語構築AI(過去成功事例DBから類似ストーリー提示)",
    },
    {
        "タスク／ジョブ": "経営会議でプレゼン、想定質問対応、次期予算・人員リソース要求、四半期計画合意",
        "行動": "役員プレゼン、Q&A対応、CMOから上層部・取締役会への展開、人員・予算交渉",
        "思考・問い": "「予算削減リスクをどう抑えるか」\n「来期人員(コンテンツ・スキーマ専門)どう確保するか」",
        "感情": "😰 緊張・プレッシャー／😌 承認時の安堵",
        "🔥 ペイン（コア）": "・「AIO/GEOは金になるのか」を疑う役員への説明難度が高い\n・予算ROIで負けると即座に投資縮小\n・人員確保が経営優先度競争で常に劣後\n・成果に時間ラグがあるため毎四半期説明が消耗",
        "タッチポイント": "経営会議 / 取締役会 / 1on1",
        "🎯 機会（プロダクト示唆）": "役員想定質問Q&Aジェネレータ＋他チャネル按分モデル＋業界ベンチマーク自動引用で説得力強化",
    },
]

# --- Persona B overview phases ---
PERSONA_B_OVERVIEW_PHASES = [
    ("① サービスライン化", "GEO/AEOを既存SEO/広告に追加する商品設計、価格表、SOP整備。\n\n【ペイン】GEOの正解SOPがまだ業界に存在しない／既存SEOチームのリスキル必要／価格設計・原価設計が手探り"),
    ("② 提案・受注(ピッチ)", "見込み客に「AI上での御社の現状」を可視化して提案。\n\n【ペイン】無料で即見せるPitch環境がないと差別化困難／意思決定者の決裁ハードルが高い／クライアントは数字を即見たい"),
    ("③ オンボーディング", "契約締結、KPI合意、エンティティ・ICP・競合のヒアリング。\n\n【ペイン】クライアントごとにツール再設定の工数／シート課金のツールはチーム全員に配れない／データアクセス権取得に時間"),
    ("⑧ 更新・アップセル", "四半期レビュー、追加スコープ提案、解約防止。\n\n【ペイン】売上アトリビューションが弱いと更新不能／Sentiment悪化やSoV低下で即詰問／追加スコープ提案の根拠データが弱い"),
    ("⑨ ポートフォリオ拡大", "業界横展開・新規ピッチ・ケーススタディ作成。\n\n【ペイン】他社事例を出せないNDA問題／人材採用が追いつかず案件断念／大手代理店との価格競争"),
]

# --- Persona B: Phase ④ ---
P_B_4 = [
    {
        "タスク／ジョブ": "クライアントブランドのKnowledge Graph位置確認、競合エンティティ・ネットワーク・引用元ページの構造分析",
        "行動": "クライアントのPMM/PRと壁打ち、Wikidata/Wikipedia/レビューサイト調査、競合の被引用ページを構造分解",
        "思考・問い": "「自社クライアントが競合より構造的に弱い理由を1枚絵で示したい」\n「KGの未掲載部分は要修復」",
        "感情": "🧠 分析・発見モード",
        "🔥 ペイン（コア）": "・KG/Wikidata/外部レビュー含めたエンティティ全体像を見るツールが分断\n・競合の被引用ページ構造分析が手作業\n・コンサル工数が見積を超過しやすい",
        "タッチポイント": "Wikidata / 競合サイト / 自社Notion",
        "🎯 機会（プロダクト示唆）": "エンティティ統合ダッシュボード(KG/Wiki/レビュー/被引用構造を1画面)＋競合構造分解AI",
    },
    {
        "タスク／ジョブ": "クライアントの売上ドライバーに直結するMoney Prompt(10〜30本)を選定、その後ファネル・ペルソナ別に300本へ拡張",
        "行動": "クライアントのCRM Top商談・営業の生質問・カスタマーサクセスのFAQから抽出、ICP×ファネル×プロダクトでマトリクス化",
        "思考・問い": "「30本に絞り込むのが難しい」\n「クライアントの担当者と意思決定者で関心プロンプトが違う」",
        "感情": "🤔 取捨選択の悩み",
        "🔥 ペイン（コア）": "・Money Promptに正解がなく属人的\n・クライアントの営業データへのアクセスが限定的(NDA/権限)\n・プロンプトをファネル拡張する自動化がない",
        "タッチポイント": "クライアントCRM / 営業MTG / プロンプト管理画面",
        "🎯 機会（プロダクト示唆）": "クライアントCRM/HubSpot連携でMoney Prompt自動抽出AI＋ファネル拡張AI",
    },
    {
        "タスク／ジョブ": "全プロンプトをChatGPT/Perplexity/Gemini/Copilotで一斉実行→ベースライン記録、KPIと90日ロードマップをクライアント承認",
        "行動": "レポートテンプレ展開、KPI(Mention/Citation/SoV/Sentiment)合意、改修ロードマップ提示、契約スコープ確定",
        "思考・問い": "「ベースラインで期待値ズレが起きないように説明したい」\n「90日で見える成果は何か事前に握りたい」",
        "感情": "😎 提案ピーク／😬 期待値調整プレッシャー",
        "🔥 ペイン（コア）": "・ベースライン提示の場で「思ったより低い」と落胆される\n・90日ロードマップに根拠データが弱いと契約スコープ縮小される\n・KPI合意に複数MTGが必要",
        "タッチポイント": "キックオフ報告会 / 提案資料 / 契約書",
        "🎯 機会（プロダクト示唆）": "ベースライン期待値マネジメントテンプレ(業界ベンチマーク同梱)＋90日ロードマップ自動生成",
    },
]

# --- Persona B: Phase ⑤ ---
P_B_5 = [
    {
        "タスク／ジョブ": "朝一でマルチクライアントダッシュボードを確認、緑(順調)/黄(要注意)/赤(即対応)でクライアント分類",
        "行動": "ダッシュボード一覧表示、Slackで担当者にエスカレ、CS担当との同期",
        "思考・問い": "「今日触るべきは5社中どこか」\n「赤クライアントが3社あると詰む」",
        "感情": "😵 マルチタスク",
        "🔥 ペイン（コア）": "・マルチクライアントダッシュボード非対応ツールでは毎回切替\n・クライアントごとKPI閾値が異なり、信号灯統一が困難\n・シート課金だとチーム全員に渡せない",
        "タッチポイント": "マルチテナント管理画面 / Slack / Asana",
        "🎯 機会（プロダクト示唆）": "マルチテナント＋信号灯ビュー(クライアント横断、緑/黄/赤)＋全プラン無制限シート",
    },
    {
        "タスク／ジョブ": "クライアントごとSLA(応答時間、報告頻度)に合わせてアラートをトリアージ、担当者へルーティング",
        "行動": "アラート優先度マトリクス(影響×緊急性)、Slack/メール/Asanaで担当アサイン、対応履歴記録",
        "思考・問い": "「このアラートはクライアント報告すべきか自社で吸収か」",
        "感情": "🚨 緊急対応・優先順位ストレス",
        "🔥 ペイン（コア）": "・アラート粒度がクライアントニーズと合わない(過剰／過少)\n・SLA違反リスク(24時間以内応答等)の可視化なし\n・対応履歴が個人Slack/メールに分散",
        "タッチポイント": "アラート管理 / メール / クライアントSlackコネクト",
        "🎯 機会（プロダクト示唆）": "アラート優先度AI(クライアントSLA連携)＋対応履歴の自動記録・引き継ぎ",
    },
    {
        "タスク／ジョブ": "週次社内レビュー、各クライアントQA(プロンプト健全性、データ整合)、スコープ外作業のチェック",
        "行動": "週次MTG、データ品質チェック、契約範囲再確認、超過工数の上長相談",
        "思考・問い": "「スコープ外作業を吸収しすぎて利益率悪化していないか」",
        "感情": "😩 収益性プレッシャー",
        "🔥 ペイン（コア）": "・プロンプト枠が固定だと配分の融通効かず\n・スコープ外対応が利益を圧迫するが断りづらい\n・社内QA基準が属人的、新人教育コスト高",
        "タッチポイント": "週次MTG / Notion SOP / 工数管理ツール",
        "🎯 機会（プロダクト示唆）": "クレジット制(プロンプト枠を流動配分)＋SOP自動QAチェック＋スコープ超過アラート",
    },
]

# --- Persona B: Phase ⑥ ---
P_B_6 = [
    {
        "タスク／ジョブ": "プロンプトギャップ→改修対象ページ→ブリーフ作成(質問形H2、回答サマリ、原子段落、Schema要件、引用源指定)",
        "行動": "競合被引用ページ参照、被引用形式テンプレ展開、ライターアサイン",
        "思考・問い": "「LLMに引用される具体的構造をライターに伝えるのが難しい」",
        "感情": "🧠 設計集中",
        "🔥 ペイン（コア）": "・AEO対応ブリーフのテンプレが業界に確立されていない\n・ライター毎に品質ばらつき大\n・被引用構造の勝ちパターンが言語化されていない",
        "タッチポイント": "ブリーフドキュメント / Notion / Figma",
        "🎯 機会（プロダクト示唆）": "AEOブリーフ自動生成AI(被引用構造テンプレ＋クライアントトーン学習)＋勝ちパターンライブラリ",
    },
    {
        "タスク／ジョブ": "外部ライター発注、ファクトチェック、Schema検証、スキャン性レビュー、ブランドトーン整合",
        "行動": "ブリーフ送付、ドラフト受領、QAチェックリスト適用、Schema妥当性検証ツール、社内ピアレビュー",
        "思考・問い": "「ファクトの正確性とAEO最適化を両立させる工数が大きい」",
        "感情": "🔥 納期プレッシャー",
        "🔥 ペイン（コア）": "・Schema実装はクライアントWeb開発の協力必須だが優先順位後回し\n・ファクトチェック工数が読めない\n・ブランドトーンとAEO最適化の衝突(短文化・原子段落化への抵抗)",
        "タッチポイント": "ライターSlack / Google Docs / Schema検証",
        "🎯 機会（プロダクト示唆）": "Schema自動生成・検証＋AEO対応QAチェックリスト自動適用＋ライタートレーニング教材同梱",
    },
    {
        "タスク／ジョブ": "クライアントレビュー→修正→公開→GSC再クロール→4〜8週間後再計測",
        "行動": "クライアントへ提出、修正対応、CMS入稿(クライアント側 or 自社)、GEOツールで再プロンプト実行",
        "思考・問い": "「クライアント承認に2〜3週間かかる」\n「再計測で結果が出ない場合の言い訳をどう用意するか」",
        "感情": "😬 承認待ちストレス／📈 結果待ち",
        "🔥 ペイン（コア）": "・クライアント承認プロセスが長く納期スリップ\n・4〜8週間ラグで結果証明が遅い\n・効果が出ない場合の説明が毎回ゼロから",
        "タッチポイント": "クライアントメール / CMS / GEOツール",
        "🎯 機会（プロダクト示唆）": "クライアント承認ワークフロー(コメント・承認・履歴管理一体)＋再計測自動化＋結果未達時の原因分析AI",
    },
]

# --- Persona B: Phase ⑦ ---
P_B_7 = [
    {
        "タスク／ジョブ": "GEOツール、GA4、GSC、CRM、広告データを各クライアント分集約・整合",
        "行動": "データエクスポート、Looker Studio/AgencyAnalytics連携、各クライアント用プリセット適用",
        "思考・問い": "「毎月クライアント分のデータ収集に2〜5日かかる」",
        "感情": "😩 集約地獄",
        "🔥 ペイン（コア）": "・データソース分断で毎月収集に膨大工数\n・クライアントごとデータ構造が微妙に違う\n・シート課金でチームに渡せず1名に集中",
        "タッチポイント": "Looker Studio / AgencyAnalytics / 各種ツールAPI",
        "🎯 機会（プロダクト示唆）": "マルチソース自動統合(GEO+GA4+GSC+CRM+広告)＋クライアント別プリセット＋全データAPI/MCP",
    },
    {
        "タスク／ジョブ": "クライアントロゴ・カラーで月次PDF/HTMLレポート生成、サマリコメント・洞察・次月推奨アクションを記述",
        "行動": "テンプレレポートに数字差し込み、アナリスト所感記述、Visibility Radar/Heatmap/Stack Rankingで構成",
        "思考・問い": "「3社で似たコメントを書き分けねばならず、品質ばらつき発生」\n「定型コメントに飽きられる」",
        "感情": "🧠 ストーリーテリング・🥱 反復作業",
        "🔥 ペイン（コア）": "・完全ホワイトラベル対応ツールが少ない(ロゴ・カラー・独自ドメイン・送信元メールまで)\n・コメント・推奨アクションのテキスト生成が手作業で月数十時間\n・テンプレ化するとクライアントから「定型的」と不評",
        "タッチポイント": "レポートテンプレ / DashThis / PDF",
        "🎯 機会（プロダクト示唆）": "完全ホワイトラベル(ロゴ・カラー・独自ドメイン・送信元メール)＋所感コメント生成AI(定型回避＋クライアント業界特化)",
    },
    {
        "タスク／ジョブ": "クライアントレビューMTGで報告、Q&A対応、次月スコープ・追加施策合意、リテイナー継続確認",
        "行動": "クライアント担当者・上司・関連部門と月例会、追加プロンプト枠提案、新ブランド追加案",
        "思考・問い": "「数字伸び悩み時の説明が肝」\n「アップセル提案を入れたいが時間配分難しい」",
        "感情": "😬 解約リスク不安／🚀 アップセル機会",
        "🔥 ペイン（コア）": "・ROI証明が弱いと即解約・予算削減\n・アップセル提案の根拠データ準備に追加工数\n・複数クライアントの月例MTGが集中して時間枯渇",
        "タッチポイント": "Zoom / クライアント会議室 / 提案資料",
        "🎯 機会（プロダクト示唆）": "収益アトリビューション(GA4/CRM連携でROI継続証明)＋アップセル提案AI(次月施策候補×期待効果ROIを自動提示)",
    },
]

# Opportunity summary
OPPORTUNITIES = [
    ("1", "A⑤-2／B④-2", "プロンプト選定の属人性", "CRM/Gong/Salesforce連携でMoney Prompt自動抽出AI"),
    ("2", "A⑥-1／B⑤-2", "アラートのノイズ過多", "シグナル要約Agent(本物の異変だけ抽出、確率変動を除外)"),
    ("3", "A⑦-2／B⑥-2", "コンテンツ改修の部門横断停滞", "CMSへの改修PR自動起票(Schema・FAQ・原子段落)"),
    ("4", "A⑧-1／B⑦-2", "経営/クライアント向け文章生成負荷", "3行サマリ自動生成＋ROI物語AI"),
    ("5", "A⑧-3／B⑦-3", "ROI証明が継続できない", "AI Visibility→Pipeline/Revenueアトリビューション"),
    ("6", "A⑦-3／B⑥-3", "効果計測ラグの不可視", "再計測自動スケジューリング＋未達原因分析AI"),
    ("7", "B⑤-1／B⑦-1", "多クライアント運用効率", "マルチテナント＋クレジット制＋無制限シート＋完全ホワイトラベル"),
    ("8", "B⑥-1", "AEO対応ブリーフが書けない", "AEOブリーフ自動生成＋勝ちパターンライブラリ"),
]


# ====== BUILD DOCUMENT ======
doc = Document()

# Landscape orientation
section = doc.sections[0]
new_width, new_height = section.page_height, section.page_width
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = new_width
section.page_height = new_height
section.left_margin = Cm(1.2)
section.right_margin = Cm(1.2)
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)

# Set default style
style = doc.styles["Normal"]
style.font.name = "Yu Gothic"
style.font.size = Pt(10)
rpr = style.element.get_or_add_rPr()
rfonts = rpr.find(qn("w:rFonts"))
if rfonts is None:
    rfonts = OxmlElement("w:rFonts")
    rpr.append(rfonts)
rfonts.set(qn("w:eastAsia"), "Yu Gothic")

# Title
add_heading(doc, "カスタマージャーニーマップ(GEO/AEO 領域)", level=0)
add_paragraph(doc, "ベンチマーク3社(Profound / Peec AI / Passionfruit Labs)の機能設計、業界KPI、エージェンシーSOPを統合し、2ペルソナのジャーニーを設計。", size=10, color="555555")
add_paragraph(doc, "2026年5月 作成", size=9, color="888888")

# === Persona A ===
doc.add_page_break()
add_heading(doc, "【ペルソナA】大手SaaS企業 デジタルマーケ部 GEO/AEO担当者", level=1)
add_heading(doc, "ペルソナ詳細", level=2)
add_persona_table(doc, PERSONA_A_INFO)

add_heading(doc, "ジャーニー全体像(Phase ①〜④, ⑨)", level=2)
add_paragraph(doc, "※ Phase ⑤〜⑧は次ページ以降でサブステップに分解", size=9, color="888888")
overview_table = doc.add_table(rows=len(PERSONA_A_OVERVIEW_PHASES), cols=2)
overview_table.autofit = False
for row in overview_table.rows:
    row.cells[0].width = Cm(4.5)
    row.cells[1].width = Cm(20)
for i, (phase, summary) in enumerate(PERSONA_A_OVERVIEW_PHASES):
    write_cell(overview_table.cell(i, 0), phase, size=10, bold=True, bg="E8EEF7")
    write_cell(overview_table.cell(i, 1), summary, size=9)

doc.add_page_break()
add_heading(doc, "Phase ⑤ オンボード・初期設定(標準4週間)", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["⑤-1 環境セットアップ\n(Week 1-2)", "⑤-2 プロンプト設計・タグ付け\n(Week 2-3)", "⑤-3 ベースライン取得・ダッシュボード/アラート設計\n(Week 3-4)"],
    P_A_5,
)

doc.add_page_break()
add_heading(doc, "Phase ⑥ 日次運用・モニタリング", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["⑥-1 デイリー\nアラート対応・ノイズ振分け", "⑥-2 ウィークリー\nMention/Citation/SoVトレンドレビュー", "⑥-3 マンスリー\nセンチメント・競合シフトのトリアージ"],
    P_A_6,
)

doc.add_page_break()
add_heading(doc, "Phase ⑦ 改善(コンテンツ/構造化)", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["⑦-1 ギャップ特定・優先順位付け", "⑦-2 コンテンツ改修\nSchema/FAQ/構造化＋ライティング", "⑦-3 公開・効果計測\n再計測ループ"],
    P_A_7,
)

doc.add_page_break()
add_heading(doc, "Phase ⑧ 経営レポート", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["⑧-1 データ集約・解釈\n前月比3行サマリ生成", "⑧-2 CMO/役員向けデック作成\nQBR、ROI物語", "⑧-3 経営会議プレゼン\n予算交渉・次期計画"],
    P_A_8,
)

# === Persona B ===
doc.add_page_break()
add_heading(doc, "【ペルソナB】デジタル広告代理店担当者(大手SaaSクライアント担当)", level=1)
add_heading(doc, "ペルソナ詳細", level=2)
add_persona_table(doc, PERSONA_B_INFO)

add_heading(doc, "ジャーニー全体像(Phase ①〜③, ⑧〜⑨)", level=2)
add_paragraph(doc, "※ Phase ④〜⑦は次ページ以降でサブステップに分解", size=9, color="888888")
overview_table = doc.add_table(rows=len(PERSONA_B_OVERVIEW_PHASES), cols=2)
overview_table.autofit = False
for row in overview_table.rows:
    row.cells[0].width = Cm(4.5)
    row.cells[1].width = Cm(20)
for i, (phase, summary) in enumerate(PERSONA_B_OVERVIEW_PHASES):
    write_cell(overview_table.cell(i, 0), phase, size=10, bold=True, bg="E8EEF7")
    write_cell(overview_table.cell(i, 1), summary, size=9)

doc.add_page_break()
add_heading(doc, "Phase ④ 戦略構築・ベースライン", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["④-1 エンティティマッピング・競合分析", "④-2 マネープロンプト選定\n10〜30本→拡張", "④-3 ベースライン取得\nKPI／ロードマップ合意"],
    P_B_4,
)

doc.add_page_break()
add_heading(doc, "Phase ⑤ 多クライアント並行運用", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["⑤-1 ポートフォリオ巡回\n信号灯・優先度判定", "⑤-2 アラートトリアージ\nSLA管理", "⑤-3 週次レビュー\n社内QA・スコープ管理"],
    P_B_5,
)

doc.add_page_break()
add_heading(doc, "Phase ⑥ コンテンツ・最適化納品", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["⑥-1 ブリーフ作成\n被引用構造に書ける指示書", "⑥-2 制作・QA\nSchema・ファクト・スキャン性", "⑥-3 公開・クライアント承認・再計測"],
    P_B_6,
)

doc.add_page_break()
add_heading(doc, "Phase ⑦ 月次レポーティング", level=1)
add_phase_table(
    doc,
    "サブステップ詳細",
    ["⑦-1 データ収集・統合\n複数ソース×複数クライアント", "⑦-2 ホワイトラベルレポート作成\nコメント・推奨アクション", "⑦-3 クライアントMTGプレゼン\n次月プラン合意"],
    P_B_7,
)

# === Opportunity Summary ===
doc.add_page_break()
add_heading(doc, "サブステップレベルでの追加機会(プロダクト示唆まとめ)", level=1)
opp_table = doc.add_table(rows=len(OPPORTUNITIES) + 1, cols=4)
opp_table.autofit = False
widths = [Cm(1.2), Cm(4.5), Cm(7.0), Cm(11.8)]
for row in opp_table.rows:
    for i, cell in enumerate(row.cells):
        cell.width = widths[i]
headers = ["#", "サブフェーズ", "機会の本質", "解決策(プロダクト要件)"]
for j, h in enumerate(headers):
    write_cell(opp_table.cell(0, j), h, size=10, bold=True, bg="1F3A5F", color="FFFFFF")
for i, (num, sub, essence, solution) in enumerate(OPPORTUNITIES):
    write_cell(opp_table.cell(i + 1, 0), num, size=9, bold=True, bg="F5F5F5")
    write_cell(opp_table.cell(i + 1, 1), sub, size=9, bg="F5F5F5")
    write_cell(opp_table.cell(i + 1, 2), essence, size=9)
    write_cell(opp_table.cell(i + 1, 3), solution, size=9, bg="F0F8E8")

# Sources page
doc.add_page_break()
add_heading(doc, "参考資料(Sources)", level=1)
sources = [
    "Which GEO/AEO delivers onboarding for AI results - Brandlight",
    "GEO Audit Checklist For Agencies - Wellows",
    "GEO Audit Agency Walkthrough 7-Step - DemandLocal",
    "GEO Playbook for Agencies - Superlines",
    "AEO Workflows in n8n or Zapier - Agenxus",
    "Answer Engine Optimization Guide - Frase.io",
    "How to optimize content for AI answer engines - Contentstack",
    "The State of AEO/GEO 2026 CMO Investment Report - Conductor",
    "AEO Reporting Guide for Marketing Agencies - Cairrot",
    "Board deck template for CMOs - CMO Alliance",
    "Best AEO Tools for Agencies 2026 - Rankability",
    "White Label Client Reporting for Agencies - ALM Corp",
    "HubSpot AEO Brand Sentiment Analysis",
    "Profound Pricing / Profound Partners / Profound for Agencies",
    "Peec AI Pricing for Agencies / White Label Client Report",
    "Passionfruit Labs Product Guide",
]
for s in sources:
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(s)
    style_run(run, size=9)

output_path = "/home/user/AIO/CustomerJourneyMap_GEO_AEO.docx"
doc.save(output_path)
print(f"Saved: {output_path}")
