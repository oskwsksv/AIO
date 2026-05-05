"""Generate a presentable PowerPoint of the GEO/AEO Customer Journey Maps."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Reuse the data from generate_cjm.py
import importlib.util

spec = importlib.util.spec_from_file_location("cjm", "/home/user/AIO/generate_cjm.py")
# We can't run it directly because it builds a docx; instead, we re-define data here.

# ----- Reimport data definitions by reading the module without running side-effect -----
# Simpler: copy data inline.

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

ASPECT_LABELS = ["タスク／ジョブ", "行動", "思考・問い", "感情", "🔥 ペイン", "タッチポイント", "🎯 機会"]

# Sub-phase data: each phase = list of 3 dicts
P_A_5 = [
    {
        "header": "⑤-1 環境セットアップ\n(Week 1-2)",
        "タスク／ジョブ": "ブランド／サブブランド／競合／対象モデル／対象国・言語／SSO／RBAC／監査ログ／API・BI連携を設定",
        "行動": "CSMキックオフ、競合リスト整備、Salesforce/HubSpot/GA4/Looker接続設定、社内SSO発行、IP範囲共有",
        "思考・問い": "「セキュリティ部門の追加質問が増えそう」「他システム連携を最初に詰めないと後で詰む」",
        "感情": "😅 設定地獄",
        "🔥 ペイン": "・SaaS連携の認証・権限取得に部門横断で2〜4週間\n・サブブランド／買収企業／グローバル各リージョンの設定が手作業の繰り返し\n・SOC2/RBAC等の運用ルール文書化に時間",
        "タッチポイント": "CSM Slack / 管理コンソール / 情シス・購買",
        "🎯 機会": "連携テンプレートライブラリ＋サブブランド一括コピー",
    },
    {
        "header": "⑤-2 プロンプト設計・タグ付け\n(Week 2-3)",
        "タスク／ジョブ": "ICP×購買ファネル×製品ライン×ペルソナでプロンプト300〜1,000本を起案・選定・タグ付け",
        "行動": "プロダクトマーケ・SE・営業・ABMチームへヒアリング、競合プロンプト調査、PMM資料・営業トーク・Win/Lossから抽出",
        "思考・問い": "「プロンプトの粒度・本数の正解値は？」「日本語と英語を分けて持つべき？」「ABM向けの社名×プロダクト系まで含めるか」",
        "感情": "🤔 正解のなさへの不安・属人性",
        "🔥 ペイン": "・プロンプト選定の正解SOPがない\n・1,000本規模になると重複・ファネル偏り・タグ運用が破綻\n・営業現場の生プロンプトを吸い上げる仕組みがない",
        "タッチポイント": "プロンプト管理画面 / Notion / 営業ヒアリングMTG",
        "🎯 機会": "プロンプトAI自動生成＋営業データ自動取込(Gong/Salesforce連携で実プロンプト抽出)",
    },
    {
        "header": "⑤-3 ベースライン取得・ダッシュボード/アラート設計\n(Week 3-4)",
        "タスク／ジョブ": "プロンプト初回実行→ベースライン記録、ダッシュボード分割、アラート閾値・受信者・チャネル設計、QBR/レビュー会議体設定",
        "行動": "初回スナップショット取得、Looker埋め込み、Slackチャネル作成、Sentiment急落・SoV低下のアラート設計",
        "思考・問い": "「アラート過多になると無視される」「閾値を厳しくしすぎると見逃す」「経営に出すKPIから逆算してダッシュボードを切るべき」",
        "感情": "🙂 進捗感／🔔 通知設計の悩み",
        "🔥 ペイン": "・初期データのノイズ・モデル変動を「異常値」として誤検知しがち\n・経営報告KPIに合わせたダッシュボードを後付けで作り直す\n・アラート設計の正解値が経験則に依存",
        "タッチポイント": "ダッシュボード / Slackアラート / Looker / カレンダー(QBR)",
        "🎯 機会": "アラート閾値の自動推奨＋経営KPIテンプレート(CMO向けカスタマイズ済)",
    },
]

P_A_6 = [
    {
        "header": "⑥-1 デイリー\nアラート対応・ノイズ振分け",
        "タスク／ジョブ": "朝のSlackアラート確認、緊急性判定、関係部門にエスカレーション、誤検知の振り分け",
        "行動": "通知巡回、PR部門・営業部門・SEOチームへのエスカレ、「無視」「監視継続」「即対応」の三択判断",
        "思考・問い": "「これはノイズか実シグナルか」「PRに即連絡すべき案件か」",
        "感情": "😵 通知疲れ・即応プレッシャー",
        "🔥 ペイン": "・毎日数十件のアラートで本質的なシグナルが埋もれる\n・LLMの確率的揺らぎを「変化」と誤検知する頻度が高い\n・PR/営業との連絡経路が個人Slackに依存",
        "タッチポイント": "Slack / メール / 管理コンソール",
        "🎯 機会": "シグナル要約Agent(本物の異変だけ抽出)＋ノイズ自動フィルタ",
    },
    {
        "header": "⑥-2 ウィークリー\nMention/Citation/SoVトレンドレビュー",
        "タスク／ジョブ": "Top10プロンプトのMention Rate / Citation Rate / SoV / First-Mention Rateの推移を確認、要因分析",
        "行動": "Looker/GEOツールで週次トレンドレポート確認、競合の特定ページ調査、社内Slackで共有",
        "思考・問い": "「上昇／低下の主要因は何か」「Claude/Geminiで差が出ている理由は？」",
        "感情": "🧐 分析モード",
        "🔥 ペイン": "・週次レポート作成に毎週3〜5時間手作業で消費\n・モデル間の差分の解釈に正解がない\n・SEO KPIとAEO KPIの統合ビューが存在せず別ツール併用",
        "タッチポイント": "Looker / Notion / 週次レビューMTG",
        "🎯 機会": "週次サマリ自動生成(要因の差分言語化)＋SEO/AEO統合ビュー",
    },
    {
        "header": "⑥-3 マンスリー\nセンチメント・競合シフトのトリアージ",
        "タスク／ジョブ": "センチメント悪化／競合急伸／自社引用喪失をクラスタリング、月次プランへ反映",
        "行動": "センチメント引用元のソース確認、競合の被引用ページ分析、四半期計画にリスト化",
        "思考・問い": "「センチメント悪化の根源はどこか」「競合のどのコンテンツが引用されている？」",
        "感情": "🔥 戦略思考／😬 競合焦り",
        "🔥 ペイン": "・センチメント悪化の根本原因まで辿るのに半日〜1日\n・競合のページ単位citation分析が手作業\n・四半期計画とのつなぎ込みが属人的",
        "タッチポイント": "センチメントダッシュボード / 競合ページ / Confluence",
        "🎯 機会": "根本原因AI＋競合被引用ページ自動アーカイブ",
    },
]

P_A_7 = [
    {
        "header": "⑦-1 ギャップ特定・優先順位付け",
        "タスク／ジョブ": "プロンプトギャップ・トピックギャップ・引用先ページ分析、ROIが大きい順に改修候補をリスト化",
        "行動": "ツールで「未引用だが商機あり」プロンプトを抽出、競合被引用ページを構造分解、Jiraチケット起票",
        "思考・問い": "「どこから手を付ければ最速で結果が出る？」「コンテンツ施策とPR施策のどちらが効く？」",
        "感情": "🧠 戦略集中",
        "🔥 ペイン": "・「未引用だが商機あり」の機会を手作業で抽出\n・コンテンツ vs PR vs 構造化のどれが効くか判断材料が少ない\n・JiraやAsanaへの起票が分断・属人化",
        "タッチポイント": "プロンプトギャップ画面 / Jira / Notion",
        "🎯 機会": "ギャップ→Jira/Asana自動起票＋ROI予測スコア",
    },
    {
        "header": "⑦-2 コンテンツ改修\nSchema/FAQ/構造化＋ライティング",
        "タスク／ジョブ": "H2/H3の質問化、回答サマリ冒頭挿入、原子段落化、FAQ追加、Schema実装、E-E-A-T強化",
        "行動": "コンテンツライターへのブリーフ作成、Web開発へSchema実装依頼、社内法務レビュー、PR/PMM連携",
        "思考・問い": "「LLMに引用される構造とは具体的に何か」「既存ブランドガイドラインと衝突しないか」",
        "感情": "😤 部門横断調整の苦労",
        "🔥 ペイン": "・コンテンツ・Web開発・法務・PMMの4部門横断調整で1施策あたり2〜6週間\n・Schema/FAQ実装はWeb開発のスプリントに乗りにくい\n・著者情報・E-E-A-T準備が想像以上に重い",
        "タッチポイント": "CMS / Figma / Web開発スプリント / コンテンツ会議",
        "🎯 機会": "CMSへの改修PR自動生成＋AEO対応ライティングAI＋著者プロフィール自動整備",
    },
    {
        "header": "⑦-3 公開・効果計測\n再計測ループ",
        "タスク／ジョブ": "CMS公開、サイトマップ送信、再クロール促進、4〜8週間で再計測しBefore/Afterを記録",
        "行動": "デプロイ→GSC再クロール→再プロンプト実行→ダッシュボードで前後比較→Win/Lossをドキュメント化",
        "思考・問い": "「ラグはどれぐらい？」「再計測タイミングは？」「成功施策の再現性は？」",
        "感情": "😌 達成感／📈 次の打ち手探し",
        "🔥 ペイン": "・効果反映に4〜8週間のラグでPDCAが遅い\n・LLMキャッシュ・学習周期が読めず再計測タイミング属人化\n・改修したのに引用されず原因不明",
        "タッチポイント": "GSC / GEOダッシュボード / Slack共有",
        "🎯 機会": "再計測自動スケジューリング＋Before/After自動レポート＋成功事例ナレッジAI",
    },
]

P_A_8 = [
    {
        "header": "⑧-1 データ集約・解釈\n前月比3行サマリ生成",
        "タスク／ジョブ": "GEOツール×GA4×CRM(Salesforce)×コンテンツKPIを統合、経営向け3行サマリ生成",
        "行動": "各ツールエクスポート、Looker/スプレッドシート集約、GEO→AI流入→MQL→SQL→パイプラインのファネル接続",
        "思考・問い": "「経営は400行のプロンプトログを見ない、3行で要約が必要」「AI Visibility → 売上の因果をどう示すか」",
        "感情": "😬 集約地獄",
        "🔥 ペイン": "・ツールが分断しており、データ統合に毎月8〜20時間消費\n・GEO→売上の因果推論手法が確立していない\n・指標粒度が経営×実務でずれる",
        "タッチポイント": "Looker / GEOツール / スプレッドシート / Salesforce",
        "🎯 機会": "CMO向け3行サマリ自動生成＋AI Visibility→Pipelineアトリビューション",
    },
    {
        "header": "⑧-2 CMO/役員向けデック作成\nQBR、ROI物語",
        "タスク／ジョブ": "QBR用デック作成(Visibility Radar / Heatmap / Stack Ranking / 競合差分 / ROI / 次四半期計画)",
        "行動": "スライド構成(Executive Summary / Goals vs Results / Funnel Health / SoV / 競合 / 次期計画)",
        "思考・問い": "「Visibility Radarは見栄えはいいが、説得力は弱い」「ROIストーリーをどう組み立てるか」",
        "感情": "🧠 ストーリーテリング集中",
        "🔥 ペイン": "・スライド作成に毎月数日〜1週間\n・ROI物語の型がなく毎回ゼロから構成\n・他チャネル(広告、PR、SEO)との寄与按分が困難",
        "タッチポイント": "PowerPoint/Keynote / Figma / 社内レビュー",
        "🎯 機会": "QBRデック自動生成＋ROI物語構築AI(過去成功事例DBから類似ストーリー提示)",
    },
    {
        "header": "⑧-3 経営会議プレゼン\n予算交渉・次期計画",
        "タスク／ジョブ": "経営会議でプレゼン、想定質問対応、次期予算・人員リソース要求、四半期計画合意",
        "行動": "役員プレゼン、Q&A対応、CMOから上層部・取締役会への展開、人員・予算交渉",
        "思考・問い": "「予算削減リスクをどう抑えるか」「来期人員(コンテンツ・スキーマ専門)どう確保するか」",
        "感情": "😰 緊張・プレッシャー／😌 承認時の安堵",
        "🔥 ペイン": "・「AIO/GEOは金になるのか」を疑う役員への説明難度が高い\n・予算ROIで負けると即座に投資縮小\n・成果に時間ラグがあるため毎四半期説明が消耗",
        "タッチポイント": "経営会議 / 取締役会 / 1on1",
        "🎯 機会": "役員想定質問Q&Aジェネレータ＋他チャネル按分モデル＋業界ベンチマーク自動引用",
    },
]

P_B_4 = [
    {
        "header": "④-1 エンティティマッピング・競合分析",
        "タスク／ジョブ": "クライアントブランドのKnowledge Graph位置確認、競合エンティティ・引用元ページの構造分析",
        "行動": "クライアントPMM/PRと壁打ち、Wikidata/Wikipedia/レビューサイト調査、競合の被引用ページを構造分解",
        "思考・問い": "「自社クライアントが競合より構造的に弱い理由を1枚絵で示したい」",
        "感情": "🧠 分析・発見モード",
        "🔥 ペイン": "・KG/Wikidata/外部レビュー含めたエンティティ全体像を見るツールが分断\n・競合の被引用ページ構造分析が手作業\n・コンサル工数が見積を超過しやすい",
        "タッチポイント": "Wikidata / 競合サイト / Notion",
        "🎯 機会": "エンティティ統合ダッシュボード＋競合構造分解AI",
    },
    {
        "header": "④-2 マネープロンプト選定\n10〜30本→拡張",
        "タスク／ジョブ": "クライアントの売上ドライバーに直結するMoney Prompt(10〜30本)を選定、その後ファネル別に300本へ拡張",
        "行動": "クライアントCRM Top商談・営業の生質問・CSのFAQから抽出、ICP×ファネル×プロダクトでマトリクス化",
        "思考・問い": "「30本に絞り込むのが難しい」「担当者と意思決定者で関心プロンプトが違う」",
        "感情": "🤔 取捨選択の悩み",
        "🔥 ペイン": "・Money Promptに正解がなく属人的\n・クライアントの営業データへのアクセスが限定的\n・プロンプトをファネル拡張する自動化がない",
        "タッチポイント": "クライアントCRM / 営業MTG / プロンプト管理画面",
        "🎯 機会": "クライアントCRM/HubSpot連携でMoney Prompt自動抽出AI＋ファネル拡張AI",
    },
    {
        "header": "④-3 ベースライン取得\nKPI／ロードマップ合意",
        "タスク／ジョブ": "全プロンプトを4 LLMで一斉実行→ベースライン記録、KPIと90日ロードマップをクライアント承認",
        "行動": "レポートテンプレ展開、KPI(Mention/Citation/SoV/Sentiment)合意、改修ロードマップ提示、契約スコープ確定",
        "思考・問い": "「ベースラインで期待値ズレが起きないように説明したい」「90日で見える成果は何か事前に握りたい」",
        "感情": "😎 提案ピーク／😬 期待値調整プレッシャー",
        "🔥 ペイン": "・ベースライン提示で「思ったより低い」と落胆される\n・90日ロードマップに根拠データが弱いと契約スコープ縮小される\n・KPI合意に複数MTGが必要",
        "タッチポイント": "キックオフ報告会 / 提案資料 / 契約書",
        "🎯 機会": "ベースライン期待値マネジメントテンプレ＋90日ロードマップ自動生成",
    },
]

P_B_5 = [
    {
        "header": "⑤-1 ポートフォリオ巡回\n信号灯・優先度判定",
        "タスク／ジョブ": "朝一でマルチクライアントダッシュボードを確認、緑/黄/赤でクライアント分類",
        "行動": "ダッシュボード一覧表示、Slackで担当者にエスカレ、CS担当との同期",
        "思考・問い": "「今日触るべきは5社中どこか」「赤クライアントが3社あると詰む」",
        "感情": "😵 マルチタスク",
        "🔥 ペイン": "・マルチクライアントダッシュボード非対応ツールでは毎回切替\n・クライアントごとKPI閾値が異なる\n・シート課金だとチーム全員に渡せない",
        "タッチポイント": "マルチテナント管理画面 / Slack / Asana",
        "🎯 機会": "マルチテナント＋信号灯ビュー＋全プラン無制限シート",
    },
    {
        "header": "⑤-2 アラートトリアージ\nSLA管理",
        "タスク／ジョブ": "クライアントごとSLAに合わせてアラートをトリアージ、担当者へルーティング",
        "行動": "アラート優先度マトリクス、Slack/メール/Asanaで担当アサイン、対応履歴記録",
        "思考・問い": "「このアラートはクライアント報告すべきか自社で吸収か」",
        "感情": "🚨 緊急対応・優先順位ストレス",
        "🔥 ペイン": "・アラート粒度がクライアントニーズと合わない\n・SLA違反リスクの可視化なし\n・対応履歴が個人Slack/メールに分散",
        "タッチポイント": "アラート管理 / メール / クライアントSlackコネクト",
        "🎯 機会": "アラート優先度AI(クライアントSLA連携)＋対応履歴の自動記録・引き継ぎ",
    },
    {
        "header": "⑤-3 週次レビュー\n社内QA・スコープ管理",
        "タスク／ジョブ": "週次社内レビュー、各クライアントQA(プロンプト健全性、データ整合)、スコープ外作業のチェック",
        "行動": "週次MTG、データ品質チェック、契約範囲再確認、超過工数の上長相談",
        "思考・問い": "「スコープ外作業を吸収しすぎて利益率悪化していないか」",
        "感情": "😩 収益性プレッシャー",
        "🔥 ペイン": "・プロンプト枠が固定だと配分の融通効かず\n・スコープ外対応が利益を圧迫するが断りづらい\n・社内QA基準が属人的",
        "タッチポイント": "週次MTG / Notion SOP / 工数管理ツール",
        "🎯 機会": "クレジット制(プロンプト枠を流動配分)＋SOP自動QAチェック＋スコープ超過アラート",
    },
]

P_B_6 = [
    {
        "header": "⑥-1 ブリーフ作成\n被引用構造に書ける指示書",
        "タスク／ジョブ": "プロンプトギャップ→改修対象ページ→ブリーフ作成(質問形H2、回答サマリ、原子段落、Schema要件、引用源指定)",
        "行動": "競合被引用ページ参照、被引用形式テンプレ展開、ライターアサイン",
        "思考・問い": "「LLMに引用される具体的構造をライターに伝えるのが難しい」",
        "感情": "🧠 設計集中",
        "🔥 ペイン": "・AEO対応ブリーフのテンプレが業界に確立されていない\n・ライター毎に品質ばらつき大\n・被引用構造の勝ちパターンが言語化されていない",
        "タッチポイント": "ブリーフドキュメント / Notion / Figma",
        "🎯 機会": "AEOブリーフ自動生成AI＋勝ちパターンライブラリ",
    },
    {
        "header": "⑥-2 制作・QA\nSchema・ファクト・スキャン性",
        "タスク／ジョブ": "外部ライター発注、ファクトチェック、Schema検証、スキャン性レビュー、ブランドトーン整合",
        "行動": "ブリーフ送付、ドラフト受領、QAチェックリスト適用、Schema妥当性検証ツール、社内ピアレビュー",
        "思考・問い": "「ファクトの正確性とAEO最適化を両立させる工数が大きい」",
        "感情": "🔥 納期プレッシャー",
        "🔥 ペイン": "・Schema実装はクライアントWeb開発の協力必須だが優先順位後回し\n・ファクトチェック工数が読めない\n・ブランドトーンとAEO最適化の衝突",
        "タッチポイント": "ライターSlack / Google Docs / Schema検証",
        "🎯 機会": "Schema自動生成・検証＋AEO対応QAチェックリスト自動適用＋ライタートレーニング教材同梱",
    },
    {
        "header": "⑥-3 公開・クライアント承認・再計測",
        "タスク／ジョブ": "クライアントレビュー→修正→公開→GSC再クロール→4〜8週間後再計測",
        "行動": "クライアントへ提出、修正対応、CMS入稿、GEOツールで再プロンプト実行",
        "思考・問い": "「クライアント承認に2〜3週間かかる」「再計測で結果が出ない場合の言い訳をどう用意するか」",
        "感情": "😬 承認待ちストレス／📈 結果待ち",
        "🔥 ペイン": "・クライアント承認プロセスが長く納期スリップ\n・4〜8週間ラグで結果証明が遅い\n・効果が出ない場合の説明が毎回ゼロから",
        "タッチポイント": "クライアントメール / CMS / GEOツール",
        "🎯 機会": "クライアント承認ワークフロー＋再計測自動化＋結果未達時の原因分析AI",
    },
]

P_B_7 = [
    {
        "header": "⑦-1 データ収集・統合\n複数ソース×複数クライアント",
        "タスク／ジョブ": "GEOツール、GA4、GSC、CRM、広告データを各クライアント分集約・整合",
        "行動": "データエクスポート、Looker Studio/AgencyAnalytics連携、各クライアント用プリセット適用",
        "思考・問い": "「毎月クライアント分のデータ収集に2〜5日かかる」",
        "感情": "😩 集約地獄",
        "🔥 ペイン": "・データソース分断で毎月収集に膨大工数\n・クライアントごとデータ構造が微妙に違う\n・シート課金でチームに渡せず1名に集中",
        "タッチポイント": "Looker Studio / AgencyAnalytics / 各種ツールAPI",
        "🎯 機会": "マルチソース自動統合＋クライアント別プリセット＋全データAPI/MCP",
    },
    {
        "header": "⑦-2 ホワイトラベルレポート作成\nコメント・推奨アクション",
        "タスク／ジョブ": "クライアントロゴ・カラーで月次PDF/HTMLレポート生成、サマリコメント・洞察・次月推奨アクションを記述",
        "行動": "テンプレレポートに数字差し込み、アナリスト所感記述、Visibility Radar/Heatmap/Stack Ranking構成",
        "思考・問い": "「3社で似たコメントを書き分けねばならず、品質ばらつき発生」",
        "感情": "🧠 ストーリーテリング・🥱 反復作業",
        "🔥 ペイン": "・完全ホワイトラベル対応ツールが少ない\n・コメント・推奨アクションのテキスト生成が手作業で月数十時間\n・テンプレ化するとクライアントから「定型的」と不評",
        "タッチポイント": "レポートテンプレ / DashThis / PDF",
        "🎯 機会": "完全ホワイトラベル＋所感コメント生成AI(定型回避＋クライアント業界特化)",
    },
    {
        "header": "⑦-3 クライアントMTGプレゼン\n次月プラン合意",
        "タスク／ジョブ": "クライアントレビューMTGで報告、Q&A対応、次月スコープ・追加施策合意、リテイナー継続確認",
        "行動": "クライアント担当者・上司・関連部門と月例会、追加プロンプト枠提案、新ブランド追加案",
        "思考・問い": "「数字伸び悩み時の説明が肝」「アップセル提案を入れたいが時間配分難しい」",
        "感情": "😬 解約リスク不安／🚀 アップセル機会",
        "🔥 ペイン": "・ROI証明が弱いと即解約・予算削減\n・アップセル提案の根拠データ準備に追加工数\n・複数クライアントの月例MTGが集中して時間枯渇",
        "タッチポイント": "Zoom / クライアント会議室 / 提案資料",
        "🎯 機会": "収益アトリビューション(GA4/CRM連携でROI継続証明)＋アップセル提案AI",
    },
]

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

# =========================================================
# Build PPT
# =========================================================
prs = Presentation()
# 16:9 widescreen
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
BLUE = RGBColor(0x2E, 0x5C, 0x8A)
LIGHT_BLUE = RGBColor(0xE8, 0xEE, 0xF7)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
LIGHT_RED = RGBColor(0xFF, 0xF0, 0xF0)
LIGHT_GREEN = RGBColor(0xF0, 0xF8, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_TEXT = RGBColor(0x55, 0x55, 0x55)
BLACK = RGBColor(0x22, 0x22, 0x22)


def set_text(text_frame, text, size=10, bold=False, color=BLACK, font_name="Yu Gothic", align=PP_ALIGN.LEFT):
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.05)
    text_frame.margin_right = Inches(0.05)
    text_frame.margin_top = Inches(0.03)
    text_frame.margin_bottom = Inches(0.03)
    text_frame.vertical_anchor = MSO_ANCHOR.TOP
    p = text_frame.paragraphs[0]
    p.alignment = align
    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            r = p.add_run()
        else:
            np = text_frame.add_paragraph()
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
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.5)
    shape.shadow.inherit = False
    return shape


def add_text_box(slide, left, top, width, height, text, size=10, bold=False, color=BLACK, align=PP_ALIGN.LEFT, fill=None):
    if fill is not None:
        bg = add_rect(slide, left, top, width, height, fill, line_color=RGBColor(0x99, 0x99, 0x99))
    tb = slide.shapes.add_textbox(left, top, width, height)
    set_text(tb.text_frame, text, size=size, bold=bold, color=color, align=align)
    return tb


def add_title_bar(slide, title_text, subtitle_text=None):
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.7), NAVY)
    add_text_box(slide, Inches(0.3), Inches(0.1), Inches(13), Inches(0.55), title_text, size=20, bold=True, color=WHITE)
    if subtitle_text:
        add_text_box(slide, Inches(0.3), Inches(0.42), Inches(13), Inches(0.3), subtitle_text, size=10, color=WHITE)


def add_cover_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, prs.slide_height, NAVY)
    add_text_box(slide, Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.5),
                 "カスタマージャーニーマップ", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text_box(slide, Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.8),
                 "GEO / AEO 領域", size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text_box(slide, Inches(0.5), Inches(5.3), Inches(12.3), Inches(0.5),
                 "ベンチマーク3社(Profound / Peec AI / Passionfruit Labs)を統合", size=14, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
    add_text_box(slide, Inches(0.5), Inches(5.9), Inches(12.3), Inches(0.4),
                 "2026年5月 作成", size=11, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


def add_persona_overview_slide(title, info, color_strip):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title_bar(slide, title, "ペルソナ詳細")
    add_rect(slide, Inches(0), Inches(0.7), Inches(0.15), prs.slide_height - Inches(0.7), color_strip)

    # Two-column layout: label | value
    top = Inches(1.0)
    row_h = Inches(0.55)
    label_w = Inches(2.7)
    value_w = Inches(9.8)
    left_label = Inches(0.4)
    left_value = left_label + label_w
    for i, (k, v) in enumerate(info):
        y = top + row_h * i
        add_text_box(slide, left_label, y, label_w, row_h, k, size=11, bold=True, color=NAVY, fill=LIGHT_BLUE)
        add_text_box(slide, left_value, y, value_w, row_h, v, size=11, color=BLACK, fill=WHITE)


def add_phase_slide(phase_title, phase_subtitle, sub_data, color_strip):
    """One slide per phase, with 3 sub-phase columns × 7 aspect rows."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title_bar(slide, phase_title, phase_subtitle)
    add_rect(slide, Inches(0), Inches(0.7), Inches(0.15), prs.slide_height - Inches(0.7), color_strip)

    # Layout
    margin_left = Inches(0.25)
    margin_top = Inches(0.85)
    aspect_col_w = Inches(1.25)
    sub_col_w = Inches((13.333 - 0.5 - 1.25) / 3)
    header_h = Inches(0.55)
    row_h = Inches(0.78)

    # Header row
    add_text_box(slide, margin_left, margin_top, aspect_col_w, header_h, "観点",
                 size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY)
    for j, sub in enumerate(sub_data):
        x = margin_left + aspect_col_w + sub_col_w * j
        add_text_box(slide, x, margin_top, sub_col_w, header_h, sub["header"],
                     size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=BLUE)

    # Aspect rows
    for i, label in enumerate(ASPECT_LABELS):
        y = margin_top + header_h + row_h * i
        # Determine row color
        if "ペイン" in label:
            row_fill = LIGHT_RED
        elif "機会" in label:
            row_fill = LIGHT_GREEN
        else:
            row_fill = LIGHT_GRAY
        add_text_box(slide, margin_left, y, aspect_col_w, row_h, label,
                     size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER, fill=LIGHT_BLUE)
        for j, sub in enumerate(sub_data):
            x = margin_left + aspect_col_w + sub_col_w * j
            content = sub.get(label, "")
            # Adjust font size based on content length
            length = len(content)
            if length > 180:
                fs = 7
            elif length > 120:
                fs = 8
            else:
                fs = 8
            add_text_box(slide, x, y, sub_col_w, row_h, content,
                         size=fs, color=BLACK, fill=row_fill)


def add_section_divider(text, color):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, prs.slide_height, color)
    add_text_box(slide, Inches(0.5), Inches(3.0), Inches(12.3), Inches(1.5),
                 text, size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def add_opportunities_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title_bar(slide, "サブステップレベルでの追加機会", "プロダクト示唆まとめ")

    margin_left = Inches(0.3)
    margin_top = Inches(0.85)
    col_widths = [Inches(0.6), Inches(2.0), Inches(3.5), Inches(6.6)]
    header_h = Inches(0.45)
    row_h = Inches(0.65)

    headers = ["#", "サブフェーズ", "機会の本質", "解決策(プロダクト要件)"]
    x = margin_left
    for w, h in zip(col_widths, headers):
        add_text_box(slide, x, margin_top, w, header_h, h,
                     size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY)
        x += w

    for i, (num, sub, essence, solution) in enumerate(OPPORTUNITIES):
        y = margin_top + header_h + row_h * i
        x = margin_left
        cells = [
            (num, LIGHT_GRAY, True, PP_ALIGN.CENTER),
            (sub, LIGHT_GRAY, False, PP_ALIGN.LEFT),
            (essence, WHITE, False, PP_ALIGN.LEFT),
            (solution, LIGHT_GREEN, False, PP_ALIGN.LEFT),
        ]
        for w, (content, fill, bold, align) in zip(col_widths, cells):
            add_text_box(slide, x, y, w, row_h, content, size=10, bold=bold, color=BLACK, align=align, fill=fill)
            x += w


def add_overview_phases_slide(title, phases, color_strip):
    """Slide listing brief overview phases (not sub-phased)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title_bar(slide, title, "ジャーニー全体像(サブステップ分解外フェーズ)")
    add_rect(slide, Inches(0), Inches(0.7), Inches(0.15), prs.slide_height - Inches(0.7), color_strip)

    margin_left = Inches(0.3)
    margin_top = Inches(0.9)
    label_w = Inches(2.6)
    value_w = Inches(10.0)
    n = len(phases)
    avail_h = Inches(7.5 - 1.0)
    row_h = Emu(int(avail_h / n))
    for i, (phase, summary) in enumerate(phases):
        y = margin_top + row_h * i
        add_text_box(slide, margin_left, y, label_w, row_h, phase,
                     size=12, bold=True, color=NAVY, align=PP_ALIGN.LEFT, fill=LIGHT_BLUE)
        add_text_box(slide, margin_left + label_w, y, value_w, row_h, summary,
                     size=9, color=BLACK, fill=WHITE)


PERSONA_A_OVERVIEW_PHASES = [
    ("① 異変認知", "オーガニック流入の減少傾向を察知。CMOから「ChatGPTで自社が出てこない」と一言。\n【ペイン】既存ツール(GA/Semrush)でAI流入が見えない／社内に前例ゼロで何から手を付けるか不明"),
    ("② 情報収集・学習", "GEO/AEO/LLMOの定義整理、競合動向、業界レポート読み込み。\n【ペイン】情報が断片的でベンダー発信に偏る／日本語LLM対応の情報が極端に少ない"),
    ("③ 戦略立案・KPI設計", "自社ICP×購買ファネルでプロンプト300本を選定。北極星KPI決定。\n【ペイン】プロンプトの優先順位付けに正解がない／既存SEO KPIとの連携設計が難しい"),
    ("④ ツール選定・稟議", "3〜5社をPoC比較、セキュリティ審査、年間予算化。\n【ペイン】SOC2/SSO/監査ログ等のエンタープライズ要件を満たすベンダーが限られる／購買・法務・情シス三段審査で4〜6ヶ月"),
    ("⑨ 横展開・スケール", "多言語・多リージョン展開、サブブランド追加、社内ナレッジ化。\n【ペイン】ツールが多言語・多リージョン本格対応していない／グローバルSOP統一困難"),
]

PERSONA_B_OVERVIEW_PHASES = [
    ("① サービスライン化", "GEO/AEOを既存SEO/広告に追加する商品設計、価格表、SOP整備。\n【ペイン】GEOの正解SOPがまだ業界に存在しない／既存SEOチームのリスキル必要"),
    ("② 提案・受注(ピッチ)", "見込み客に「AI上での御社の現状」を可視化して提案。\n【ペイン】無料Pitch環境がないと差別化困難／意思決定者の決裁ハードルが高い"),
    ("③ オンボーディング", "契約締結、KPI合意、エンティティ・ICP・競合のヒアリング。\n【ペイン】クライアントごとにツール再設定の工数／シート課金のツールはチーム全員に配れない"),
    ("⑧ 更新・アップセル", "四半期レビュー、追加スコープ提案、解約防止。\n【ペイン】売上アトリビューションが弱いと更新不能／Sentiment悪化やSoV低下で即詰問"),
    ("⑨ ポートフォリオ拡大", "業界横展開・新規ピッチ・ケーススタディ作成。\n【ペイン】他社事例を出せないNDA問題／人材採用が追いつかず案件断念／大手代理店との価格競争"),
]

# Build slides
add_cover_slide()

# Persona A section
add_section_divider("【ペルソナA】\n大手SaaS企業 デジタルマーケ部\nGEO/AEO担当者", NAVY)
add_persona_overview_slide("【ペルソナA】大手SaaS企業 デジタルマーケ部 GEO/AEO担当者", PERSONA_A_INFO, color_strip=BLUE)
add_overview_phases_slide("【ペルソナA】ジャーニー全体像", PERSONA_A_OVERVIEW_PHASES, color_strip=BLUE)
add_phase_slide("【A】Phase ⑤ オンボード・初期設定(標準4週間)", "ペルソナA — サブステップ詳細", P_A_5, color_strip=BLUE)
add_phase_slide("【A】Phase ⑥ 日次運用・モニタリング", "ペルソナA — サブステップ詳細", P_A_6, color_strip=BLUE)
add_phase_slide("【A】Phase ⑦ 改善(コンテンツ/構造化)", "ペルソナA — サブステップ詳細", P_A_7, color_strip=BLUE)
add_phase_slide("【A】Phase ⑧ 経営レポート", "ペルソナA — サブステップ詳細", P_A_8, color_strip=BLUE)

# Persona B section
ORANGE = RGBColor(0xC9, 0x6E, 0x29)
add_section_divider("【ペルソナB】\nデジタル広告代理店担当者\n大手SaaSクライアント担当", ORANGE)
add_persona_overview_slide("【ペルソナB】デジタル広告代理店担当者(大手SaaSクライアント担当)", PERSONA_B_INFO, color_strip=ORANGE)
add_overview_phases_slide("【ペルソナB】ジャーニー全体像", PERSONA_B_OVERVIEW_PHASES, color_strip=ORANGE)
add_phase_slide("【B】Phase ④ 戦略構築・ベースライン", "ペルソナB — サブステップ詳細", P_B_4, color_strip=ORANGE)
add_phase_slide("【B】Phase ⑤ 多クライアント並行運用", "ペルソナB — サブステップ詳細", P_B_5, color_strip=ORANGE)
add_phase_slide("【B】Phase ⑥ コンテンツ・最適化納品", "ペルソナB — サブステップ詳細", P_B_6, color_strip=ORANGE)
add_phase_slide("【B】Phase ⑦ 月次レポーティング", "ペルソナB — サブステップ詳細", P_B_7, color_strip=ORANGE)

# Opportunities
add_section_divider("プロダクト示唆まとめ", RGBColor(0x2E, 0x7A, 0x4E))
add_opportunities_slide()

output_path = "/home/user/AIO/CustomerJourneyMap_GEO_AEO.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
print(f"Total slides: {len(prs.slides)}")
