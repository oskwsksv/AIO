"""Generate story-style product introduction Word doc for AIOutmation.
Enhanced version with high-resolution pain depiction for the protagonist."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
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


def add_chapter_title(doc, num, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"第{num}話")
    style_run(r, size=11, bold=True, color="C04000")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run(title)
    style_run(r, size=20, bold=True, color="1F3A5F")


def add_body(doc, text, size=11):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Cm(1.0)
    r = p.add_run(text)
    style_run(r, size=size)


def add_dialogue(doc, text, speaker_color="2E5C8A"):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.left_indent = Cm(1.0)
    r = p.add_run(text)
    style_run(r, size=11, color=speaker_color)


def add_inner(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.left_indent = Cm(1.5)
    r = p.add_run(text)
    style_run(r, size=11, italic=True, color="555555")


def add_scene_break(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run("◆")
    style_run(r, size=11, color="888888")


def add_callout_box(doc, body, color="FFF8E0"):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, color)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.4
    r = p.add_run(body)
    style_run(r, size=10, color="555555")


def add_pain_log(doc, title, body):
    """Specific pain detail box - high contrast red-tinted."""
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, "FFE8E8")
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(3)
    r = p1.add_run(title)
    style_run(r, size=10, bold=True, color="A02020")
    p2 = cell.add_paragraph()
    p2.paragraph_format.line_spacing = 1.4
    r = p2.add_run(body)
    style_run(r, size=10, color="333333")


def cover(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(100)
    r = p.add_run("ある日、私は")
    style_run(r, size=20, color="555555")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    r = p.add_run("「AIOマネージャー」に")
    style_run(r, size=24, bold=True, color="1F3A5F")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(5)
    r = p.add_run("任命された。")
    style_run(r, size=24, bold=True, color="1F3A5F")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(60)
    r = p.add_run("〜あるマーケティング担当者の半年間〜")
    style_run(r, size=14, color="555555")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(80)
    r = p.add_run("AIOutmation")
    style_run(r, size=32, bold=True, color="2E5C8A")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(5)
    r = p.add_run("プロダクトストーリー")
    style_run(r, size=14, color="2E5C8A")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(80)
    r = p.add_run("登場人物:")
    style_run(r, size=10, color="888888")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("山下美咲(35歳) ・・・ 上場SaaS企業のマーケ部マネージャー")
    style_run(r, size=10, color="888888")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("田島CMO ・・・ 美咲の上司、52歳")
    style_run(r, size=10, color="888888")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("夫・健一 ・・・ IT企業エンジニア / 娘・葵 ・・・ 4歳")
    style_run(r, size=10, color="888888")

    doc.add_page_break()


def setup_doc(doc):
    section = doc.sections[0]
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    style = doc.styles["Normal"]
    style.font.name = "Yu Gothic"
    style.font.size = Pt(11)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "Yu Gothic")


# ============== Build Story ==============

doc = Document()
setup_doc(doc)
cover(doc)


# === 第1話 ===
add_chapter_title(doc, 1, "突然の任命")

add_body(doc, "2026年4月、月曜の朝9時5分。クラウド人事SaaS『ピープルワン』(東証グロース上場、ARR45億円)のマーケティング部、定例朝会。山下美咲は、いつものようにアイスのカフェラテを片手に、田島CMOのスライドを眺めていた。今日もまた、Google広告のCPAが先週比でわずかに悪化した話だろう、と思っていた。")

add_body(doc, "ところがその日、田島CMOが取り出したスライドには、見慣れない言葉が並んでいた。")

add_callout_box(doc, "「AEO/GEO/LLMO ─ 新しい検索の時代へ」")

add_body(doc, "「みんな、最近気づいてる? 自社サイトへのオーガニック流入が、半年前と比べて15%も減ってる。一方で、ChatGPTやPerplexityで『クラウド人事システム おすすめ』と検索する人が、競合A社のサイトを引用するシーンが増えてる。試しにPerplexityで検索してみたんだけど ─ うちのサイト、引用されないんだよ」")

add_dialogue(doc, "「これからは『検索順位』じゃない。『AIに選ばれるか』が勝負になる。だから、うちもAIO ─ AI最適化に本気で取り組む。やる人、いない?」")

add_body(doc, "会議室がしんと静まり返った。誰もが、田島CMOと目を合わせないように資料に視線を落とした。")

add_inner(doc, "(またか、と思った。誰もやらないなら、私が手を挙げる。新しい領域だし、上手くいけば評価される ─ そう思ったのは確かだ。でも、それだけじゃなかった。なんとなく、これは時代を変える仕事になる気がしていた)")

add_body(doc, "気づくと、美咲は手を挙げていた。田島CMOが嬉しそうにうなずく。")

add_dialogue(doc, "「お、山下くん。じゃあ任せた。3ヶ月後の経営会議で、AIO戦略レポートを出してくれ。新しい予算枠もつける。期待してるよ」")

add_body(doc, "拍手の中、美咲は曖昧に笑った。心の中では、すでに焦り始めていた。AEO? GEO? LLMO? 略語が3つも並んでいるが、何ひとつとして詳細を知らない。")

add_inner(doc, "(これから3ヶ月、なにから手をつけよう ─)")

doc.add_page_break()


# === 第2話 ===
add_chapter_title(doc, 2, "教科書は買った。でも──")

add_body(doc, "1週間後。美咲のデスクには、ピンク色のカバーの分厚い本が置かれていた。『これからはじめる AIO AI最適化の教科書』(技術評論社)。464ページ。週末に一気読みした。")

add_body(doc, "概念は、思ったよりも論理的だった。AIO = AEO(理解させる) + GEO(引用させる) + LLMO(学習させる)、その土台にSEO。これらを『AIOピラミッド』として階層化して、土台から順に積み上げる。文章はPREP法→FAQ構文→HowTo構文の3つの型で書く。引用率(Citation Rate)という独自指標で運用する。")

add_inner(doc, "(理屈は分かった。でも、これを実際にやるのは別問題だ ─)")

add_scene_break(doc)

add_body(doc, "5月の連休明け、美咲は教科書のChapter 4『PDCA運用』を開いて、Excelを起動した。教科書の指示通り、3種類のスプレッドシートを作る。『FAQ管理表』『FAQ記録表』『FAQ集計表』。列の名前を一つずつ手で打ち込んだ。タイトル、質問、回答、最終更新日、引用回数、引用率(%)、改善メモ ─ これだけで、午前中が消えた。")

add_pain_log(doc,
    "【ペイン①: プロンプトの正解SOPがどこにもない】",
    "そもそも『どんな質問を測るべきか』が分からない。教科書には『売上に直結するMoney Promptを10〜30本選ぶ』と書いてある。でも、自社の人事SaaSにとっての『売上に直結する質問』とは何だろう? 営業に聞きに行ったが、忙しい商談の合間に「お客さんが最初によく聞くのは、料金とサポート体制かな…たまに他社比較も…」と曖昧な答え。Win/Lossミーティングの議事録を引っ張り出して、Notionで100件読み込み、自分なりに分類した。3日かかった。それでも、これが正しいのか、自信が持てない。")

add_body(doc, "次に、教科書のChapter 4で示された『手動の引用確認』を始めた。Perplexity、ChatGPT(検索機能付き)、Google AI Overviews、Geminiの4つのAIを開く。1つの質問につき、それぞれのAIに5回ずつ検索する。回答に自社のURLが出ているかを目視で確認する。")

add_pain_log(doc,
    "【ペイン②: 月25回×4AI×30プロンプト = 物理的限界】",
    "教科書には『月1回、12.5分で済む』と書いてある。確かに30本のプロンプトなら、4AI×5回=600回の検索を1日でやれば、12.5分×30本=6時間ちょっとだ。だが、現実はそうはいかない。Perplexityの回答が表示されるまで5〜15秒、Geminiは認証が切れるたびにログイン、Google AI Overviewsは出たり出なかったり、ChatGPTは月額20ドルの個人アカウントが急にレート制限。途中で集計シートにスクショを貼り付ける作業が入り、1プロンプト30秒どころか実際は2〜3分。気づけば、夜の10時を過ぎていた。")

add_inner(doc, "(これって、ダイエットの体重計を毎日手書きで管理してるのと同じじゃない? なんで2026年にもなってこんなことを ─)")

add_pain_log(doc,
    "【ペイン③: LLMの『揺らぎ』に振り回される】",
    "教科書には書いていない、もう一つの問題があった。同じ質問でも、AIの答えは毎回少しずつ違う。月曜日にPerplexityで「クラウド人事 おすすめ」と打つと自社が3番目に引用されるが、火曜日には引用されなくなる。改善したから消えたのか、それとも単なる確率的揺らぎなのか ─ 美咲には判別する術がない。週次レビューで田島CMOに「先週は出てたんですけど、今週は出てないんです」と説明する自分が、だんだん馬鹿らしく感じてきた。")

add_pain_log(doc,
    "【ペイン④: 改修すべきページが分からない】",
    "Citation Rateが低いプロンプトの原因を分析するため、競合A社の引用ページを開いて、構造を目視で分析する。h1タグの使い方、FAQ構造、Schemaの実装。10ページ分を分解して、自社のどのページをどう書き直すかをWordに落とし込む。それからライターさんへブリーフを書く。さらに、Schema実装はWeb開発チームに依頼するが、彼らの今四半期スプリントは『新機能リリース最優先』で、私の依頼は7月のスプリントへ持ち越しになった。「8週間後ですか…」と返事をしながら、美咲は唇を噛んだ。")

add_scene_break(doc)

add_body(doc, "それでも、地道に続けた。3ヶ月後、ようやく引用率は20%から35%まで伸びた。CMOへの中間報告は、なんとか合格点をもらえた。")

add_dialogue(doc, "「いい仕事だ、山下くん。これを来期、全社展開しよう。多言語対応(英・中・韓・タイ)、北米と欧州のサブブランド10個、競合10社の分析も加えて」")

add_inner(doc, "(……全社展開? 私一人で?)")

doc.add_page_break()


# === 第3話 ===
add_chapter_title(doc, 3, "限界の夜と、出会い")

add_body(doc, "5月末。美咲は精神的に追い詰められていた。手作業の量は3倍に膨れ上がった。プロンプトは300本、対応言語は4ヶ国語、サブブランドは10個、競合は10社。1日の検索回数は、計算上、4,000回を超える。実際にやってみて、3日目で限界が来た。")

add_pain_log(doc,
    "【ペイン⑤: スケールが破綻、検索作業だけで毎日3時間】",
    "PCの前に座り、午前9時から正午まで、ひたすらAIに同じ質問を打ち続けた。スクショをDropboxに保存し、ファイル名は『20260520_perplexity_query01_run3_no.png』のようなルール。500枚を超えたあたりから、自分が何を保存したか分からなくなる。Spreadsheetに貼り付ける作業も、関数のエラーで何度もやり直し。気づけば、肩こりがひどくなり、首が回らない。眼科に行ったら『ドライアイです』と診断された。")

add_pain_log(doc,
    "【ペイン⑥: 多言語対応の人格切替地獄】",
    "英語版のサイトを測るときは、VPNでアメリカ西海岸に接続する必要がある。中国語版は別のVPN、韓国語版もまた別。各AIで言語別アカウントを作り直し、毎回ログインし直す。AI Overviewsはリージョンによって表示挙動が違うので、地域別に再現条件を整えて検索する。Spreadsheetも言語別に4つに分割した。タブが多すぎて、どこに何を書いたか自分でも分からなくなる夜が増えた。")

add_pain_log(doc,
    "【ペイン⑦: SEO・広告との因果関係が説明できない】",
    "田島CMOから月次レビューで質問が来る。「で、結局AI経由で売上はどう推移してるの? 役員会で説明できる3行は?」 美咲は、400行のSpreadsheetを開きながら、口ごもった。引用率の推移は出せる。でも、それが売上にどうつながっているのか、Salesforceとどう紐づければよいのか分からない。広告とSEOとAIO、それぞれが何%寄与しているのか ─ その因果モデルは、世界中のどの教科書にも載っていない。「次回までに整理します」とだけ答えて、会議室を後にした。")

add_pain_log(doc,
    "【ペイン⑧: 家族との時間が消える】",
    "5月23日、土曜日。4歳の娘・葵の保育園の運動会。美咲は朝、葵に「今日はお仕事だから、お父さんと行ってきてね」と言って、リビングのPCの前に座った。来週月曜の経営会議資料を作るため、Spreadsheetから手作業で数値を抜き出し、グラフを6枚、PowerPointに貼り付けた。夕方、夫の健一が葵を連れて帰ってきたとき、葵は「ママ、今日リレーで一等賞だったんだよ」と笑った。美咲は写真を見せてもらい、目を伏せた。「すごいね」と言いながら、心の奥で、何かが崩れていく音がした。")

add_pain_log(doc,
    "【ペイン⑨: 退職を本気で考えた金曜日】",
    "5月29日、午前1時20分。リビングの蛍光灯がチカチカ点滅する中、美咲はSpreadsheetを睨んでいた。Perplexityの認証が切れて、再ログインを促されたとき、ふと、転職サイトのアイコンが目に入った。LinkedInを開く。元同僚が外資系ITに転職して、AIマーケティングのSr. Manager職に就いていた。年収は1.5倍、成功したとSNSに書いていた。美咲は、退職届のテンプレートを検索した。書きかけて、ノートPCを閉じた。涙が出た。")

add_inner(doc, "(……もう辞めたい。マーケティングって、こんなに虚しい仕事だったっけ ─)")

add_scene_break(doc)

add_body(doc, "6月3日、業界カンファレンス『MarkeZine Day 2026』。美咲は半ば義務感で参加していた。最終セッションが終わり、疲れて帰ろうとしたそのとき、ブースの一角で目に飛び込んできたフレーズに、足が止まった。")

add_callout_box(doc, "『AIO最適化を、書く前から書いた後まで自動で回す。』 ─ AIOutmation", color="E0F4E0")

add_body(doc, "ブースの担当者に話を聞いた。AIO教科書のChapter 4で説明されている『手作業のPDCA』を、ソフトウェアで全自動化するプラットフォームだという。")

add_dialogue(doc, "「Spreadsheetの3種類の表は、私たちのダッシュボードに置き換わります。月25回の手動検索は、毎日4つのAIで自動実行されます。多言語・サブブランド・競合分析もすべて並列で走ります。月次レポートは1分でホワイトラベル化されたPDFが出力されます。CMOへの3行サマリも自動生成です」")

add_inner(doc, "(……それ、本当?)")

add_body(doc, "美咲は、半信半疑のまま、その場で2週間の無料トライアルを申し込んだ。家に帰る電車の中、はじめて、明日が楽しみだと思った。")

doc.add_page_break()


# === 第4話 ===
add_chapter_title(doc, 4, "魔法のような半年間")

add_body(doc, "翌週月曜、美咲はAIOutmationにログインした。最初に画面に表示されたのは、たった4つの質問だった。「自社のWebサイトURLは?」「主要競合は?」「対象とするAIはどれ?」「主要顧客の業種は?」")

add_body(doc, "入力すると、たった5分で、AIが300本のプロンプト案をタグ付き(購買ファネル/ペルソナ/プロダクト別)で生成してくれた。これまで3日かけてWin/Loss議事録から手書きで作っていたプロンプトが、5分で全部揃った。しかも、教科書のQuery→Intent→Answerモデルに沿った分類になっている。")

add_inner(doc, "(これだけで、私の3週間が、終わった ─)")

add_scene_break(doc)

add_body(doc, "2週間後。AIOutmationのダッシュボードには、毎日4つのAIで自動実行された結果が、リアルタイムで蓄積されていた。引用率の推移グラフ、競合との比較、ヒートマップ。月に1回しか見えなかった景色が、毎朝コーヒーを飲みながら見られるようになった。スクショを保存することも、ファイル名を考えることも、Spreadsheetを更新することも、もうない。VPN切替も、人格切替も、ない。")

add_pain_log(doc,
    "【ペイン②③⑤⑥が解決された朝】",
    "毎朝9時、美咲はSlackに届くダッシュボード通知を見るだけになった。『今日のCitation Rate: 全体42%(前日比+1.2pt)。注目: 米国ブランドの「prep talent management」プロンプトで競合B社が急伸』。LLMの揺らぎはAIOutmation側で複数回実行+統計処理してノイズが除去されているので、本物の変化だけが報告される。多言語サブブランドも全部1画面で見られる。VPNを切替える日々は、もう過去のものだった。")

add_body(doc, "ある朝、Slackにアラートが届いた。「競合A社が『勤怠管理 おすすめ』のプロンプトで、過去3日で引用シェアを20ポイント伸ばしています」。AIOutmationの分析エンジンによれば、競合A社が新しいFAQページを公開し、構造化データを追加したことが原因だった。改善案までAIが提案してくれた ─ 自社のFAQに『退職時の最終勤怠処理』のQ&Aを追加し、Schema.org/FAQPageを実装すべき、と。")

add_body(doc, "美咲は、その提案を社内のライターさんとWeb開発チームに転送し、その週のうちに更新が完了した。1週間後、Citation Rateは逆転した。")

add_pain_log(doc,
    "【ペイン④が解決された日】",
    "AIOutmationは、競合の被引用ページを構造分析して、自社で作るべき改善案をJSON形式で出力してくれる。ライターさんは、その出力を見るだけでブリーフを理解できた。Web開発チームへの依頼も、Schemaコードがすでに添付されているので、彼らはコピー&ペーストするだけで済んだ。8週間スプリント待ちは、過去の話になった。")

add_scene_break(doc)

add_body(doc, "3ヶ月後の経営会議。CMOへの月次レビューは、もう手作業のスライド作りに数日かけることはなかった。AIOutmationが自動生成した『3行サマリ』を、そのまま使った。")

add_callout_box(doc,
"""【今月の3行サマリ】
・全体引用率: 35% → 58%(前月比+23pt、業界中央値+12pt、北米サブブランドは+31pt)
・カテゴリ別: 勤怠管理が好調(+30pt)、年末調整は競合C社に劣後(-8pt、要追加投資)
・競合差分: 自社は『連携の柔軟性』『日本市場の対応』で優位、『価格透明性』で劣位
・推定: AI経由のMQL寄与は前月比2.1倍、Salesforce実数値で確認済み""", color="E0F4E0")

add_dialogue(doc, "「素晴らしい。今度は取締役会で報告したい。資料を作ってくれないか? 来週の水曜日だ」")

add_body(doc, "美咲は微笑んで答えた。「はい、もうあります」。AIOutmationが自動生成した取締役会用デックを、その場で開いた。Visibility Radar、競合との戦略マップ、四半期ROI、AI流入と売上のアトリビューション分析 ─ すべて、ボタン1つで揃っていた。")

add_pain_log(doc,
    "【ペイン⑦⑧が解決された夜】",
    "10月14日、金曜日の夕方6時。美咲はオフィスを出た。葵を保育園に迎えに行く時間だった。家でハンバーグを作って、3人で夕食を食べた。葵が「ママ、今日も保育園で○○ちゃんが ─」と話しかけてきて、美咲はそれをきちんと聞いた。週末は、家族で公園に行った。残業時間は、月20時間に減った。")

add_scene_break(doc)

add_body(doc, "半年後。AI経由の月間流入は、導入前の3倍に。さらに重要な変化があった。AIから引用された情報経由で、年間契約2億円の大型商談が、3件成約した。営業部が朝会で美咲に感謝を述べた瞬間、彼女は静かに、深く、一息ついた。")

add_inner(doc, "(あのとき、手を挙げて良かった)")

doc.add_page_break()


# === 第5話 ===
add_chapter_title(doc, 5, "新しい役割")

add_body(doc, "12月の取締役会。社長の篠崎が、議題の最後に切り出した。")

add_dialogue(doc, "「来期、当社は『マーケティングDX推進室』を新設する。AIO・SEO・コンテンツを統合した部門だ。室長は ─ 山下美咲。来期からは部長待遇で、君に任せたい」")

add_body(doc, "美咲はマイクの前で言葉に詰まった。隣で田島CMOが満足そうにうなずいている。")

add_dialogue(doc, "「……ありがとうございます。精一杯、務めます」")

add_scene_break(doc)

add_body(doc, "帰り道、美咲はオフィスの自販機の前で立ち止まり、いつものアイスのカフェラテを買った。窓の外には、初冬の夕焼けが広がっている。")

add_body(doc, "5月のあの夜、深夜2時にSpreadsheetと格闘していた自分を思い出した。あの頃、自分は『仕事に潰されている』と感じていた。退職届のテンプレートを検索した夜のことも、葵の運動会に行けなかった土曜日のことも、まだ忘れられない。")

add_body(doc, "でも、本当に変わったのは、ツールではなく、『仕事の意味』だった。手作業に追われていたとき、美咲は『AIに選ばれるための雑用係』だった。AIOutmationが手作業を引き受けてくれた瞬間、美咲は『ブランドの未来を設計する戦略家』に変わった。")

add_callout_box(doc,
"""手作業を、ソフトウェアに。
戦略を、人間に。

それが、AIOutmation が信じている、これからのマーケティングのあり方です。""",
color="E0F4E0")

doc.add_page_break()


# === エピローグ ===
add_chapter_title(doc, "エピローグ", "AIOutmationが解決する 主要なペイン")

add_body(doc, "美咲のストーリーで描かれた9つの具体的なペインが、AIOutmationの主要機能でどう解決されたかを整理します。")

# Pain to feature mapping
pains_features = [
    ("ペイン① プロンプトの正解SOPがない",
     "F2: プロンプト自動生成エンジン",
     "自社URL+競合URL+Win/Lossデータから、ICP×購買ファネル別にAIがMoney Promptを300本提案。3日→5分。"),
    ("ペイン② 月25回×4AI×30プロンプト = 物理的限界",
     "F3: Citation Rate Engine",
     "Perplexity/ChatGPT/Gemini/AI Overviewsで毎日4,000回の検索を自動実行。手動検索ゼロに。"),
    ("ペイン③ LLMの『揺らぎ』に振り回される",
     "F3: Signal/Noise処理(MVPは閾値ベース、V2で統計処理強化)",
     "複数回実行+移動平均+閾値設定で『本物の変化』のみを通知。CMOに馬鹿げた説明をする日々は終了。"),
    ("ペイン④ 改修すべきページが分からない",
     "F4: ダッシュボード+原因分析",
     "競合の被引用ページを自動分析し、自社の改修案をJSON出力。Schemaコードまで自動生成。Web開発のスプリント待ち解消。"),
    ("ペイン⑤ スケール破綻、毎日3時間の検索作業",
     "F1+F3: マルチテナント基盤+Citation Rate Engine",
     "プロンプト数千本、複数ブランド、複数地域を1画面で並列管理。検索作業は完全に消滅。"),
    ("ペイン⑥ 多言語対応の人格切替地獄",
     "F1: マルチリージョン対応",
     "言語別・地域別の設定をWorkspace単位で分離。VPN切替・アカウント切替は不要に。"),
    ("ペイン⑦ SEO・広告との因果関係が説明できない",
     "F6: AIコメンタリー(MVP)+ V2 ROIアトリビューション",
     "Salesforce/CRM連携で『AI経由のMQL/SQL/受注』を自動マッピング。3行サマリで経営に即提示。"),
    ("ペイン⑧ 家族との時間が消える",
     "F1〜F6 全機能の総合効果",
     "残業時間 月80時間 → 月20時間。土曜日は娘の運動会に行ける。最も大切な、人生のペインを解決。"),
    ("ペイン⑨ 退職を考えた金曜日",
     "AIOutmation そのもの",
     "『AIに選ばれるための雑用係』から『ブランドの未来を設計する戦略家』へ。仕事の意味を取り戻す。"),
]

for pain, feature, desc in pains_features:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(pain)
    style_run(r, size=11, bold=True, color="A02020")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(f"→ {feature}")
    style_run(r, size=10, bold=True, color="2E5C8A")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.left_indent = Cm(1.0)
    r = p.add_run(desc)
    style_run(r, size=10, color="333333")


# Final closing message
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(30)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("もし、あなたが今日、")
style_run(r, size=12, color="555555")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("『AIO担当』に任命されたら ─")
style_run(r, size=14, bold=True, color="2E5C8A")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(10)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("最初に、AIOutmationを開いてください。")
style_run(r, size=14, bold=True, color="1F3A5F")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(10)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("あなたの「9つのペイン」は、もう昨日までのものです。")
style_run(r, size=11, color="888888")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(40)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("─ AIOutmation Team ─")
style_run(r, size=10, color="888888")

doc.save("/home/user/AIO/AIOutmation_Story.docx")
print("Saved: AIOutmation_Story.docx")
