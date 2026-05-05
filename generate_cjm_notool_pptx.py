"""Generate PPT for the No-Tool (manual) Customer Journey Maps."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# ======================== DATA ========================

PERSONA_A_PREMISE = [
    ("担当人数", "兼任1名(既存SEO担当の片手間)、エンジニア協力 月数時間"),
    ("使用ツール", "ChatGPT/Perplexity/Gemini/Claude のWeb UI、Google Sheets、Looker Studio、Semrush/Ahrefs、GA4、Search Console、Notion、Schema.org検証ツール、自作Pythonスクリプト"),
    ("計測対象プロンプト", "20〜50本(手作業の限界) ←→ ツール導入時 300〜1,000本"),
    ("計測頻度", "月1回の手動更新 ←→ ツール導入時 24時間ごと自動"),
    ("経営報告頻度", "四半期に1回 ←→ ツール導入時 月次"),
]

PERSONA_B_PREMISE = [
    ("体制", "アカウント担当1名＋SEOアナリスト共有0.5名、外部ライター"),
    ("使用ツール", "各社LLMのWeb UI、Google Sheets(クライアント別タブ)、Looker Studio、AgencyAnalytics、Semrush/Ahrefs、Notion"),
    ("担当クライアント数の現実", "手動なら同時に最大2〜3社が限界。5社以上は破綻"),
    ("クライアントあたりプロンプト数", "10〜20本(Money Promptのみ) ←→ ツール導入時 300〜1,000本"),
    ("月次レポート所要時間", "1社あたり 2〜4日 ←→ ツール導入時 数時間"),
]

ASPECT_LABELS = ["タスク／ジョブ", "行動", "思考・問い", "感情", "🔥 ペイン", "タッチポイント", "💡 ツールがあれば"]

# Persona A — 8 phases
P_A_PHASES = [
    {
        "header": "① 異変認知・問題定義",
        "タスク／ジョブ": "CMO質問「ChatGPTで弊社は？」を皮切りに、「測れていない」状態の言語化",
        "行動": "GA4のリファラ調査、社内Slackで議論、CMO 1on1で「現状検討中」と回答",
        "思考・問い": "「測れる前提でCMOが質問してくる…どう答えるか」",
        "感情": "😨 焦り・恥ずかしさ",
        "🔥 ペイン": "・「測定できていない」自体が説明困難で、CMOから次々質問が来ても応答に時間ラグ\n・既存SEOツール(Semrush/Ahrefs)はAI検索をカバーしていない\n・GA4/GSCではAI流入がほぼ識別不能",
        "タッチポイント": "GA4 / Slack / CMO 1on1",
        "💡 ツールがあれば": "─",
    },
    {
        "header": "② 情報学習・社内啓蒙",
        "タスク／ジョブ": "GEO/AEO概念を独学、社内勉強会開催、ベンダーブログ・Xで情報収集",
        "行動": "Notion個人メモ、外部ウェビナー視聴、SEOチームに共有",
        "思考・問い": "「何が業界標準KPIなのかわからない」「日本語事例少ない」",
        "感情": "🤔 情報過多・孤独",
        "🔥 ペイン": "・GEO/AEO情報がベンダー発信に偏り、客観的な業界標準KPIが不明\n・日本語LLM(Felo / Genspark / Perplexity日本語)の事例が極めて少ない\n・社内には前例ゼロで勉強しながら設計する孤独感",
        "タッチポイント": "Notion / X(Twitter) / ウェビナー / SEOチーム",
        "💡 ツールがあれば": "─",
    },
    {
        "header": "③ 手動ベースライン調査",
        "タスク／ジョブ": "自社×競合×プロダクトで20〜50プロンプトを起案、ChatGPT/Perplexity/Gemini/Claudeに手動入力",
        "行動": "1プロンプトあたり2〜3分×4モデル×30本=最低6時間 手動入力、スクショ保存、Sheet貼付",
        "思考・問い": "「30本でも半日。300本だったら週単位の作業」「手動だと揺らぎなのか変化なのか不明」",
        "感情": "😩 単純作業疲れ・腱鞘炎",
        "🔥 ペイン": "・1プロンプトあたり2〜3分×4モデルで物理的に1日30〜50本が限界\n・LLMの確率的揺らぎで再現条件を保証できない\n・スクショ保存→ファイル名管理→Sheet貼付で時間消費\n・国・言語・モデルを変えるたびに人格切替(VPN/別アカウント)が必要",
        "タッチポイント": "ChatGPT / Perplexity / Gemini / Claude のWeb UI、スクショフォルダ",
        "💡 ツールがあれば": "自動プロンプト実行、4モデル横断、揺らぎを統計処理",
    },
    {
        "header": "④ Sheets運用設計",
        "タスク／ジョブ": "6カラム構成のSheet設計(Platform/Query/Mention?/Citation URL/Sentiment/Action)、運用ルール文書化",
        "行動": "カラム設計、フィルタ／ピボット作成、関係者にURL共有、入力ルール周知",
        "思考・問い": "「Sheet設計が甘いと後でデータ統合できない」「他メンバーが同じ品質で入力できるか不安」",
        "感情": "🧐 設計集中／属人化への不安",
        "🔥 ペイン": "・Sheetのスキーマが後から変わると過去データと整合せず全部やり直し\n・複数人で入力すると入力品質ばらつき(Mention判定基準等)\n・フィルタ・ピボット・グラフが管理工数を圧迫\n・バージョン管理が破綻(コピー乱立)",
        "タッチポイント": "Google Sheets / 関係者MTG",
        "💡 ツールがあれば": "スキーマ固定、自動入力、バージョン管理",
    },
    {
        "header": "⑤ 月次手動再計測",
        "タスク／ジョブ": "同じプロンプトを月1で手動再実行、結果をSheetに入力、前月比手作業計算",
        "行動": "全プロンプト再実行、過去結果と比較、変化があれば赤字で目立たせる",
        "思考・問い": "「先月と本当に同じ条件か？」「ChatGPT仕様変更で結果が変わったらどう判断？」",
        "感情": "🥱 反復作業の倦怠",
        "🔥 ペイン": "・1モデル仕様変更で過去比較が無効化\n・ノイズ(揺らぎ)と本物の変化を分離する手法がない\n・差分計算を毎月手動で計算して人為ミス頻発\n・競合の被引用ページ調査が手作業で1社あたり半日\n・月1の頻度では急変を1ヶ月遅れで気づく",
        "タッチポイント": "Sheets / カレンダー",
        "💡 ツールがあれば": "24時間自動再実行、ノイズ除去アルゴリズム、変化検知アラート",
    },
    {
        "header": "⑥ 手動コンテンツ改修",
        "タスク／ジョブ": "プロンプトギャップから改修対象を選定、ライターにブリーフ作成、Schema実装をWeb開発に依頼",
        "行動": "競合の被引用ページを手動分析(目視)、Wordでブリーフ起草",
        "思考・問い": "「LLMに引用される構造が言語化できない」「Schema実装の優先度がスプリントで上がらない」",
        "感情": "😤 部門横断調整の苦労",
        "🔥 ペイン": "・「LLMに引用される構造」の業界共通言語化が未整備でブリーフが書けない\n・Schema実装はWeb開発のスプリントに乗せられず3〜6ヶ月待ち\n・改修後の再計測手段がなく、効果証明できない\n・ライター教育コストが大きく毎回ゼロから",
        "タッチポイント": "CMS / Jira / コンテンツ会議 / Schema検証ツール",
        "💡 ツールがあれば": "引用構造ライブラリ、CMS連携PR自動起票、ブリーフAI生成",
    },
    {
        "header": "⑦ 経営報告",
        "タスク／ジョブ": "四半期の経営会議用に、Sheetsから手作業でグラフ作成、CMO向け3行サマリ起草",
        "行動": "Sheetsから数字抽出→PowerPoint→Visibility Radar/積み上げ棒グラフ手作り",
        "思考・問い": "「Visibility Radarをパワポで自作するの何回目？」「経営は本気でこの数字を信じるか」",
        "感情": "😬 説得自信のなさ",
        "🔥 ペイン": "・スプレッドシート→PPTの変換に毎回数日\n・経営は3行で要求するが、データ抽出ロジックの説明に時間取られる\n・ROI接続できない(AI Visibility→売上の因果が示せない)\n・業界ベンチマーク不明で自社の立ち位置を経営に伝えられない\n・競合鮮度が古い(手動更新)",
        "タッチポイント": "PowerPoint / 経営会議",
        "💡 ツールがあれば": "ダッシュボード自動生成、ROIアトリビューション、業界ベンチマーク同梱",
    },
    {
        "header": "⑧ スケール壁・限界",
        "タスク／ジョブ": "プロンプト100本超への拡張・多言語化・サブブランド対応の検討で詰まる",
        "行動": "拡張試算、追加リソース要求の社内提案、自作Pythonスクリプト検討",
        "思考・問い": "「もう個人運用の限界。ツール導入の稟議を本気で書くか、諦めるか」",
        "感情": "😵 燃え尽き／🚪 諦めムード",
        "🔥 ペイン": "・100本超に拡張するリソース確保不可能\n・多言語×多リージョン×サブブランドを組み合わせると工数爆発\n・自作Pythonスクリプトを作っても、LLMのログイン仕様変更で動かなくなる\n・個人依存で休暇・退職時に運用停止\n・ナレッジが個人Notionに閉じて横展開不能",
        "タッチポイント": "社内提案書 / Python IDE",
        "💡 ツールがあれば": "スケール(1,000本以上)、多言語、サブブランド、APIによる自動化",
    },
]

# Persona B — 8 phases
P_B_PHASES = [
    {
        "header": "① クライアントから問合せ受領",
        "タスク／ジョブ": "クライアントCMOから「うちChatGPTで出てる？」と質問。社内で「やります」と回答",
        "行動": "クライアントメール確認、社内に持ち帰り、上司と相談",
        "思考・問い": "「やる方法はあるが、収益化の型がない」「断ると他社に取られる」",
        "感情": "🤔 期待半分・不安半分",
        "🔥 ペイン": "・業界共通のサービスSOPが存在せず「やります」と言ったが具体手順を即答できない\n・既存SEOチームにAEO/GEO知見ゼロで社内リスキル必要\n・クライアントの期待値合意ができない",
        "タッチポイント": "クライアントメール / 上司MTG / 社内Slack",
        "💡 ツールがあれば": "─",
    },
    {
        "header": "② サービス設計(手探り)",
        "タスク／ジョブ": "サービスメニュー、価格表、SOPをゼロから設計",
        "行動": "Notionでサービス定義、SEOチームから知見輸入、価格設定試行錯誤",
        "思考・問い": "「他社代理店はどうやって型化しているのか」「価格はリテイナーいくらで売れる？」",
        "感情": "😣 設計の孤独・焦り",
        "🔥 ペイン": "・サービスメニュー設計に正解なしで価格・スコープが他社と比較不能\n・既存SEOリテイナーとの重複・カニバリが不明\n・SOPがないため人によって品質ばらつき\n・教育資料・トーク・営業ピッチを全てゼロから",
        "タッチポイント": "Notion / 価格表エクセル / 競合代理店ブログ",
        "💡 ツールがあれば": "サービスSOPテンプレ、業界価格ベンチマーク",
    },
    {
        "header": "③ ピッチ・受注",
        "タスク／ジョブ": "提案資料作成、ピッチ。ChatGPTで競合と並べたスクショを差し込み",
        "行動": "スライド作成、ChatGPT/Perplexityで現状デモ、商談Zoom",
        "思考・問い": "「無料でも見せてあげないと意思決定者の心は掴めない」",
        "感情": "😎 提案ピーク／🥶 ピッチ準備の徹夜",
        "🔥 ペイン": "・無料でAI上の現状を見せる「Pitch環境」を作るために手動で半日\n・各LLMでスクショ取って加工してパワポ貼付\n・意思決定者が「で、どれぐらい改善できる？」と聞いてもベンチマークがない\n・ROI根拠が薄いため値引き交渉に弱い",
        "タッチポイント": "提案資料 / ChatGPT / Zoom",
        "💡 ツールがあれば": "Pitch環境(即発行・期限付き)、業界ベンチマーク",
    },
    {
        "header": "④ オンボード・ベースライン",
        "タスク／ジョブ": "契約締結、KPI合意、Money Prompt 10〜20本を選定、初回手動実行→ベースライン",
        "行動": "キックオフMTG、Sheets共有、初回データ取得(1クライアント半日〜1日)",
        "思考・問い": "「1社設定するだけで半日。これを5社やったら週末まで埋まる」",
        "感情": "😩 工数地獄スタート",
        "🔥 ペイン": "・1クライアントの初期セットアップに半日〜1日(プロンプト選定・初回実行・Sheet設計)\n・5社受注したら初週はオンボーディングで何もできない\n・KPI閾値・期待値合意が毎回ゼロから\n・データアクセス(GA4/GSC/CRM)取得に2週間",
        "タッチポイント": "キックオフMTG / Sheets / クライアントSlack",
        "💡 ツールがあれば": "プロンプトAI自動生成、テンプレ一括コピー、CRM自動連携",
    },
    {
        "header": "⑤ 多クライアント手動運用",
        "タスク／ジョブ": "5社並行で月1手動再計測。アラートなしのため週1で「なんとなく確認」",
        "行動": "クライアント別Sheetsを巡回、変化を目視、Slackでアラート手送",
        "思考・問い": "「3社目あたりから細部の記憶が混ざる」「アラートなしでは月遅れの気付き」",
        "感情": "😵 マルチタスク混乱",
        "🔥 ペイン": "・プロンプト枠が固定でなくクライアント別管理でSheets構造爆発\n・Slack/メール/Asana/Notionにアラート分散、SLA違反気付かず\n・3クライアント目で記憶混乱、入力先間違い頻発\n・休暇・退職時に運用停止、引継ぎ資料は個人Notion頼み\n・月1再計測の頻度では競合急伸を月遅れで発見",
        "タッチポイント": "各クライアントSheets / Slack / Asana",
        "💡 ツールがあれば": "マルチテナント信号灯、SLAアラート、対応履歴自動記録、無制限シート",
    },
    {
        "header": "⑥ コンテンツ改修指示・QA",
        "タスク／ジョブ": "改修対象ページ特定、ブリーフ手動作成、ライターQA、CMS入稿協議",
        "行動": "競合被引用ページを手動構造分析、ライターSlack、Schema検証ツール手動",
        "思考・問い": "「ライターごとに構造の理解度が違う。再修正コストが利益を食う」",
        "感情": "😤 品質ばらつきへのストレス",
        "🔥 ペイン": "・AEO対応ブリーフのテンプレ未確立で毎回属人的に書く\n・ライターごとに被引用構造の理解度が異なり再修正コスト膨張\n・Schema実装はクライアントWeb開発の協力必要だが優先度低\n・改修後の再計測手段がなく「効果あり」を証明できず再修正提案も弱い",
        "タッチポイント": "ブリーフドキュメント / ライターSlack / Schema検証ツール",
        "💡 ツールがあれば": "AEOブリーフ自動生成、Schema自動生成、勝ちパターンライブラリ",
    },
    {
        "header": "⑦ 月次手動レポート作成",
        "タスク／ジョブ": "クライアントごとSheets→Looker Studio→PDF化、コメント手書き",
        "行動": "クライアント別データエクスポート、コメント執筆、ロゴ・カラー手動差替え",
        "思考・問い": "「同じ作業を5社分するのはなんでこんなに時間かかるんだ」「コメントが定型的になる」",
        "感情": "😩 反復作業疲労",
        "🔥 ペイン": "・1社あたりレポート作成2〜4日＝5社で半月消える\n・ロゴ・カラー手動差替え、コメントは似た文章を手書き\n・Visibility Radar/Heatmap/Stack Rankingを毎月手動\n・業界ベンチマークを引用したくても自社過去データのみ\n・白ラベル・独自ドメイン送信を実現するには別途ツール契約",
        "タッチポイント": "Looker Studio / PowerPoint / PDF",
        "💡 ツールがあれば": "完全ホワイトラベル、コメントAI、3社分自動生成、業界ベンチマーク同梱",
    },
    {
        "header": "⑧ クライアント説明・更新交渉",
        "タスク／ジョブ": "月例会で報告、Q&A対応、追加スコープ提案、解約防止",
        "行動": "月例会Zoom、提案資料、リテイナー継続交渉",
        "思考・問い": "「成果が薄い月は数字以外で説明する技術が必要」",
        "感情": "😬 解約リスク不安／🚀 アップセル狙い",
        "🔥 ペイン": "・ROI証明が弱いと即予算削減・解約\n・月例会で「数字伸びてないですね」と言われて回答準備に時間\n・アップセル提案根拠データの追加準備でさらに工数\n・Sentiment悪化やSoV低下を数字で説明しきれず\n・他クライアントの月例会と重なると時間枯渇",
        "タッチポイント": "Zoom / 提案資料 / 契約書",
        "💡 ツールがあれば": "収益アトリビューション、アップセル提案AI",
    },
]

TOP_PAINS = [
    ("1位", "時間消費が桁違い", "担当者の80%の時間がデータ収集に消え、戦略業務が消滅", "LLMがプログラマブルアクセスを十分に提供せず、Web UIが標準"),
    ("2位", "属人化・知識流出リスク", "担当者の休暇・退職で運用停止", "SOPがなく、Sheetsの設計・運用ノウハウが個人脳内"),
    ("3位", "ノイズ vs シグナルの分離不能", "LLMの揺らぎを「変化」と誤検知 or 本物の変化を見逃し", "統計処理を手動で行えず、月1観測では揺らぎ除去できない"),
    ("4位", "経営/クライアント説得力の欠如", "予算削減・解約リスク、追加投資の承認得られず", "データ粒度薄／業界ベンチマーク非搭載／ROIアトリ不能"),
    ("5位", "スケール壁(30→100→1,000本)", "戦略の幅が狭まり、競合に劣後", "物理的人手の限界、多言語・多リージョン・サブブランド対応不能"),
]

IMPLICATIONS = [
    ("A", "価値提案の中核は「時間圧縮」", "「年間1,500時間の手作業を100時間に」のような時間単位ROIが最も訴求力ある(特に経営承認)"),
    ("B", "ハイブリッド戦略の余地", "中小／個人事業主／SMB予算向けに「Sheetsを資産にしたまま自動化レイヤを乗せる」プロダクトは未開拓ニッチ"),
    ("C", "経営説明書類の自動化", "ペルソナA⑦・B⑧のレポート工数は、ツールを導入しても残る最大ペイン。レポート自動化AIは継続価値"),
    ("D", "SOPの民主化", "ペルソナB①〜②の「業界SOPが存在しない」問題は、ツールにSOPテンプレを内蔵すれば差別化軸"),
    ("E", "段階的導入(Crawl→Walk→Run)", "30本/月1から始めて、300本/日次に拡張するパスを設計すれば、決裁ハードルを下げられる"),
]


# ======================== BUILD ========================

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
BLUE = RGBColor(0x2E, 0x5C, 0x8A)
ORANGE = RGBColor(0xC9, 0x6E, 0x29)
RED = RGBColor(0xB0, 0x3A, 0x2E)
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


def add_box(slide, left, top, width, height, text, size=10, bold=False, color=BLACK, align=PP_ALIGN.LEFT, fill=None):
    if fill is not None:
        add_rect(slide, left, top, width, height, fill, line_color=RGBColor(0x99, 0x99, 0x99))
    tb = slide.shapes.add_textbox(left, top, width, height)
    set_text(tb.text_frame, text, size=size, bold=bold, color=color, align=align)
    return tb


def title_bar(slide, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.65), NAVY)
    add_box(slide, Inches(0.3), Inches(0.07), Inches(13), Inches(0.55), title, size=20, bold=True, color=WHITE)
    if subtitle:
        add_box(slide, Inches(0.3), Inches(0.4), Inches(13), Inches(0.25), subtitle, size=10, color=WHITE)


def cover():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, prs.slide_height, NAVY)
    add_box(slide, Inches(0.5), Inches(2.0), Inches(12.3), Inches(1.2),
            "【ツール非導入時】", size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(3.0), Inches(12.3), Inches(1.2),
            "カスタマージャーニーマップ", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.6),
            "GEO / AEO 領域 (専用ツール非導入時の手動運用)", size=18, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(5.5), Inches(12.3), Inches(0.4),
            "ベンチマーク3社のような専用ツールを使わず、Web UI＋Sheets＋汎用ツールで運用するケース", size=12, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
    add_box(slide, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.4),
            "2026年5月 作成", size=11, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


def divider(text, color):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, prs.slide_height, color)
    add_box(slide, Inches(0.5), Inches(3.0), Inches(12.3), Inches(1.5),
            text, size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def premise_slide(title, premise, strip_color):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, title, "前提:手動運用時のリソース・スタック")
    add_rect(slide, Inches(0), Inches(0.65), Inches(0.15), prs.slide_height - Inches(0.65), strip_color)
    top = Inches(1.1)
    row_h = Inches(0.85)
    label_w = Inches(3.5)
    value_w = Inches(9.3)
    for i, (k, v) in enumerate(premise):
        y = top + row_h * i
        add_box(slide, Inches(0.4), y, label_w, row_h, k,
                size=12, bold=True, color=NAVY, fill=LIGHT_BLUE)
        add_box(slide, Inches(0.4) + label_w, y, value_w, row_h, v,
                size=11, color=BLACK, fill=WHITE)


def phases_slide(slide_title, subtitle, phases, strip_color):
    """Render up to 4 phases per slide, aspects as rows."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, slide_title, subtitle)
    add_rect(slide, Inches(0), Inches(0.65), Inches(0.15), prs.slide_height - Inches(0.65), strip_color)

    n = len(phases)
    margin_left = Inches(0.2)
    margin_top = Inches(0.8)
    aspect_col_w = Inches(1.1)
    sub_col_w = Inches((13.333 - 0.4 - 1.1) / n)
    header_h = Inches(0.5)
    row_h = Inches(0.78)

    add_box(slide, margin_left, margin_top, aspect_col_w, header_h, "観点",
            size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY)
    for j, ph in enumerate(phases):
        x = margin_left + aspect_col_w + sub_col_w * j
        add_box(slide, x, margin_top, sub_col_w, header_h, ph["header"],
                size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=BLUE)

    for i, label in enumerate(ASPECT_LABELS):
        y = margin_top + header_h + row_h * i
        if "ペイン" in label:
            row_fill = LIGHT_RED
        elif "ツール" in label:
            row_fill = LIGHT_GREEN
        else:
            row_fill = LIGHT_GRAY
        add_box(slide, margin_left, y, aspect_col_w, row_h, label,
                size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER, fill=LIGHT_BLUE)
        for j, ph in enumerate(phases):
            x = margin_left + aspect_col_w + sub_col_w * j
            content = ph.get(label, "")
            length = len(content)
            if length > 200:
                fs = 6
            elif length > 130:
                fs = 7
            else:
                fs = 8
            add_box(slide, x, y, sub_col_w, row_h, content,
                    size=fs, color=BLACK, fill=row_fill)


def top_pains_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, "ツール非導入時の最大ペイン Top 5", "両ペルソナから抽出")
    margin_left = Inches(0.3)
    margin_top = Inches(0.85)
    cols = [Inches(0.8), Inches(2.6), Inches(4.0), Inches(5.6)]
    headers = ["ランク", "ペイン", "影響", "構造的原因"]
    header_h = Inches(0.5)
    row_h = Inches(1.05)
    x = margin_left
    for w, h in zip(cols, headers):
        add_box(slide, x, margin_top, w, header_h, h,
                size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY)
        x += w
    for i, (rank, pain, impact, cause) in enumerate(TOP_PAINS):
        y = margin_top + header_h + row_h * i
        x = margin_left
        items = [
            (rank, LIGHT_BLUE, True, PP_ALIGN.CENTER, NAVY, 12),
            (pain, LIGHT_RED, True, PP_ALIGN.LEFT, BLACK, 11),
            (impact, WHITE, False, PP_ALIGN.LEFT, BLACK, 10),
            (cause, LIGHT_GRAY, False, PP_ALIGN.LEFT, BLACK, 10),
        ]
        for w, (txt, fill, bold, align, color, sz) in zip(cols, items):
            add_box(slide, x, y, w, row_h, txt, size=sz, bold=bold, color=color, align=align, fill=fill)
            x += w


def implications_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_bar(slide, "上流設計への含意", "あなたのプロダクトへの示唆")
    margin_left = Inches(0.3)
    margin_top = Inches(0.85)
    cols = [Inches(0.5), Inches(3.4), Inches(9.0)]
    headers = ["#", "示唆", "内容"]
    header_h = Inches(0.5)
    row_h = Inches(1.15)
    x = margin_left
    for w, h in zip(cols, headers):
        add_box(slide, x, margin_top, w, header_h, h,
                size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY)
        x += w
    for i, (letter, headline, body) in enumerate(IMPLICATIONS):
        y = margin_top + header_h + row_h * i
        x = margin_left
        items = [
            (letter, LIGHT_YELLOW, True, PP_ALIGN.CENTER, NAVY, 14),
            (headline, LIGHT_BLUE, True, PP_ALIGN.LEFT, NAVY, 12),
            (body, WHITE, False, PP_ALIGN.LEFT, BLACK, 10),
        ]
        for w, (txt, fill, bold, align, color, sz) in zip(cols, items):
            add_box(slide, x, y, w, row_h, txt, size=sz, bold=bold, color=color, align=align, fill=fill)
            x += w


# === Build the deck ===
cover()

divider("【ペルソナA】\n大手SaaS企業 デジタルマーケ部\nGEO/AEO担当者", NAVY)
premise_slide("【ペルソナA】手動運用の前提", PERSONA_A_PREMISE, BLUE)
phases_slide("【A】Phase ①〜④", "ペルソナA — ツール非導入時 ジャーニー前半", P_A_PHASES[0:4], BLUE)
phases_slide("【A】Phase ⑤〜⑧", "ペルソナA — ツール非導入時 ジャーニー後半", P_A_PHASES[4:8], BLUE)

divider("【ペルソナB】\nデジタル広告代理店担当者\n大手SaaSクライアント担当", ORANGE)
premise_slide("【ペルソナB】手動運用の前提", PERSONA_B_PREMISE, ORANGE)
phases_slide("【B】Phase ①〜④", "ペルソナB — ツール非導入時 ジャーニー前半", P_B_PHASES[0:4], ORANGE)
phases_slide("【B】Phase ⑤〜⑧", "ペルソナB — ツール非導入時 ジャーニー後半", P_B_PHASES[4:8], ORANGE)

divider("ツール非導入時の最大ペイン Top 5", RED)
top_pains_slide()

divider("上流設計への含意", RGBColor(0x2E, 0x7A, 0x4E))
implications_slide()

output = "/home/user/AIO/CustomerJourneyMap_NoTool_GEO_AEO.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
