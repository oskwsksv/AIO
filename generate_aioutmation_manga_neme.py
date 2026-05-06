"""Generate manga storyboard (ネーム/絵コンテ) for AIOutmation product story.
This is an illustrator-ready document that specifies panel composition,
character expressions, dialogue, and SFX for each panel."""
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


def write_cell(cell, text, size=9, bold=False, bg=None, color=None, italic=False):
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
        run = p.add_run(line)
        style_run(run, size=size, bold=bold, color=color, italic=italic)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level <= 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    sizes = {0: 22, 1: 18, 2: 14, 3: 12, 4: 11}
    colors = {0: "1F3A5F", 1: "1F3A5F", 2: "2E5C8A", 3: "C04000", 4: "555555"}
    style_run(run, size=sizes.get(level, 11), bold=True, color=colors.get(level, "000000"))


def add_paragraph(doc, text, size=10, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    style_run(run, size=size, bold=bold, color=color)


def add_callout_box(doc, body, color="FFF8E0"):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, color)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.4
    r = p.add_run(body)
    style_run(r, size=10)


def setup_doc(doc):
    section = doc.sections[0]
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    style = doc.styles["Normal"]
    style.font.name = "Yu Gothic"
    style.font.size = Pt(10)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "Yu Gothic")


def cover(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(80)
    r = p.add_run("ある日、私は")
    style_run(r, size=18, color="555555")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("「AIOマネージャー」に")
    style_run(r, size=22, bold=True, color="1F3A5F")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("任命された。")
    style_run(r, size=22, bold=True, color="1F3A5F")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(40)
    r = p.add_run("ビジネス漫画ネーム(絵コンテ)")
    style_run(r, size=20, bold=True, color="C04000")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    r = p.add_run("AIOutmation プロダクトストーリー(全5話・約20ページ想定)")
    style_run(r, size=14, color="555555")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(120)
    r = p.add_run("本資料はイラストレーター/漫画家への発注用の指示書です。")
    style_run(r, size=11, color="888888")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("各コマの構図・キャラクター描写・台詞・効果音を詳細に記載しています。")
    style_run(r, size=11, color="888888")

    doc.add_page_break()


def add_panel_table(doc, panels):
    """Render a panel table for one manga page.
    panels: list of dict with keys: num, size, composition, visual, dialogue, sfx"""
    n = len(panels)
    t = doc.add_table(rows=n + 1, cols=6)
    t.autofit = False
    widths = [Cm(0.8), Cm(1.5), Cm(2.2), Cm(7.5), Cm(5.0), Cm(2.0)]
    for r in t.rows:
        for i, c in enumerate(r.cells):
            c.width = widths[i]
    headers = ["コマ", "サイズ", "構図", "描画内容(キャラ・背景・小物)", "台詞・モノローグ", "SFX効果音"]
    for i, h in enumerate(headers):
        write_cell(t.cell(0, i), h, size=9, bold=True, bg="1F3A5F", color="FFFFFF")
    for ri, p in enumerate(panels):
        row = ri + 1
        write_cell(t.cell(row, 0), str(p["num"]), size=10, bold=True, bg="E8EEF7")
        write_cell(t.cell(row, 1), p.get("size", "中"), size=9)
        write_cell(t.cell(row, 2), p.get("composition", ""), size=9)
        write_cell(t.cell(row, 3), p.get("visual", ""), size=9)
        write_cell(t.cell(row, 4), p.get("dialogue", ""), size=9, color="2E5C8A", bold=True)
        write_cell(t.cell(row, 5), p.get("sfx", ""), size=9, color="C04000", italic=True)


# ============== Build Document ==============
doc = Document()
setup_doc(doc)
cover(doc)

# === 制作仕様 ===
add_heading(doc, "1. 制作仕様", level=1)

add_heading(doc, "1.1 形式・体裁", level=2)
add_paragraph(doc, "形式: ビジネス漫画(セミナー漫画/SaaS紹介漫画系)。4コマではなく1ページに4〜6コマ配置の通常漫画フォーマット。", size=11)
add_paragraph(doc, "ページ数: 全20ページ前後(第1〜5話＋エピローグ)", size=11)
add_paragraph(doc, "想定媒体: WebマンガLP / 営業資料 / 採用資料(印刷用にも耐える解像度)", size=11)
add_paragraph(doc, "読み方向: 右開き(伝統的な日本漫画フォーマット)", size=11)
add_paragraph(doc, "想定読了時間: 5〜7分", size=11)

add_heading(doc, "1.2 アートスタイル指示", level=2)
add_paragraph(doc, "・絵柄: ビジネス漫画寄り、リアル系。少年漫画ほどデフォルメせず、青年誌(ビッグコミック等)レベルのリアリティ。", size=10)
add_paragraph(doc, "・線画: クリアな線、描き込みは中程度。背景はオフィス・自宅・カフェなど現代日本の情景を丁寧に描く。", size=10)
add_paragraph(doc, "・スクリーントーン: 心理描写(暗いシーン)に多めに使用。明るいシーンはトーン控えめ。", size=10)
add_paragraph(doc, "・カラー: モノクロが基本。ただし表紙・要所(第4話の解決シーン以降)で一部カラー使用も検討可。", size=10)
add_paragraph(doc, "・参考作品: 『海猿』(リアル系)、『ドラゴン桜』(ビジネス系教育漫画)、『重版出来!』(働く女性主人公)。", size=10)

add_heading(doc, "1.3 キャラクター設定", level=2)

t = doc.add_table(rows=4, cols=2)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(4.0)
    r.cells[1].width = Cm(15.0)

write_cell(t.cell(0, 0), "山下美咲(主人公)", size=11, bold=True, bg="E8EEF7")
write_cell(t.cell(0, 1),
"""年齢: 35歳 / 性別: 女性 / 役職: クラウド人事SaaS『ピープルワン』マーケティング部マネージャー
身長: 162cm / 髪型: 黒髪セミロング(肩より少し下、毛先軽くウェーブ)
服装: 上場企業勤めのきちんとしたオフィスカジュアル(ジャケット+パンツ、白〜ネイビー系)
表情の特徴: 真面目で勝ち気だが、内面はセンシティブ。第2-3話では疲労と眼の下のクマが進行的に深くなる。
小物: 銀縁メガネ(在宅時着用)、アイスカフェラテのプラカップ、ノートPC(ThinkPad想定)、社員証ストラップ""", size=10)

write_cell(t.cell(1, 0), "田島CMO(上司)", size=11, bold=True, bg="E8EEF7")
write_cell(t.cell(1, 1),
"""年齢: 52歳 / 性別: 男性 / 役職: ピープルワン社CMO(最高マーケティング責任者)
体型: 中肉、やや恰幅、鬼ではないが厳格
髪型: ややごま塩。サイドにシャープな白髪
服装: スーツ(濃紺・チャコールグレー)、ノーネクタイ、カフスボタンに少しこだわり
表情: 朝礼ではエネルギッシュ。月次レビューでは鋭い視線。第4話以降は満足げな表情に変化。""", size=10)

write_cell(t.cell(2, 0), "葵(美咲の娘)", size=11, bold=True, bg="E8EEF7")
write_cell(t.cell(2, 1),
"""年齢: 4歳 / 髪型: ショートボブ、可愛い前髪
特徴: 母譲りの大きな目。明るく元気、運動会では一等賞をとる活発な女の子
重要シーン: 第3話で運動会の話、第4話で家族の食卓""", size=10)

write_cell(t.cell(3, 0), "健一(美咲の夫)", size=11, bold=True, bg="E8EEF7")
write_cell(t.cell(3, 1),
"""年齢: 36歳 / 職業: IT企業エンジニア / 髪型: 短めの黒髪
特徴: 優しく協力的だが、美咲が忙しすぎることを密かに心配している。第3話の運動会で葵を連れていく。""", size=10)

add_heading(doc, "1.4 トーン&ムード", level=2)
add_paragraph(doc, "・第1話: 期待と不安が入り混じる。明るめの空気。", size=10)
add_paragraph(doc, "・第2話: 暗くなり始める。教科書を真面目に読む姿は希望的だが、手作業描写でじわじわ疲労感。", size=10)
add_paragraph(doc, "・第3話: 暗黒期。深夜の蛍光灯、ドライアイ、家族との断絶、退職を考える夜。最大の感情ピーク。", size=10)
add_paragraph(doc, "・第4話: 明るい解放感。青空、晴れた朝、家族の笑顔。", size=10)
add_paragraph(doc, "・第5話: 達成感。成熟した美咲の表情、夕焼け、未来への前向きさ。", size=10)

doc.add_page_break()


# === 第1話 ===
add_heading(doc, "2. 第1話「突然の任命」(全3ページ・12コマ想定)", level=1)

add_heading(doc, "P.1 (オープニング・任命前夜の朝)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "大", "composition": "鳥瞰・俯瞰", "visual":
     "高層ビル群の朝の空、ピープルワン社オフィスビル全景。ビルに『PeopleOne』のロゴ。月曜日の通勤ラッシュ。",
     "dialogue": "(タイトル文字)\n第1話「突然の任命」", "sfx": "ザワザワ\n朝のざわめき"},
    {"num": 2, "size": "中", "composition": "横向き・寄り", "visual":
     "オフィスの自席で美咲がアイスカフェラテをストローで吸っている。デスクにはノートPC、観葉植物、社員証。窓の外は晴れた春の朝。",
     "dialogue": "美咲(モノローグ): また月曜の朝礼か。\n今日もどうせCPAの話だろう。", "sfx": "シューッ\n(ストロー音)"},
    {"num": 3, "size": "中", "composition": "中距離・引き", "visual":
     "会議室、長机を囲んで10人ほどのマーケ部メンバー。前方のスクリーンには田島CMOがスライドを示す。",
     "dialogue": "田島CMO: みんな、これを見てくれ。", "sfx": ""},
    {"num": 4, "size": "大", "composition": "スライド大写し", "visual":
     "スクリーンに大きく『AEO/GEO/LLMO ─ 新しい検索の時代へ』の文字。背景にAIアイコン(ChatGPT/Perplexity/Geminiを暗示する象徴)が並ぶ。",
     "dialogue": "田島CMO(吹き出し): これからは『検索順位』じゃない。\n『AIに選ばれるか』が勝負だ。", "sfx": "ザワッ"},
])

add_heading(doc, "P.2 (任命の瞬間)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "CMOアップ", "visual":
     "田島CMOの顔のアップ、まじめで力強い表情、目に力強さ。",
     "dialogue": "田島CMO: AIO ─ AI最適化に本気で取り組む。\nやる人、いない?", "sfx": ""},
    {"num": 2, "size": "大", "composition": "会議室全体俯瞰", "visual":
     "10人のメンバー全員が下を向いて資料を見るふりをしている。シーンと静まり返った空気。",
     "dialogue": "(沈黙)", "sfx": "シーン..."},
    {"num": 3, "size": "中", "composition": "美咲の表情アップ", "visual":
     "美咲の顔のアップ、目に複雑な感情(迷い+好奇心+若干の野心)。眉が少し上がる。",
     "dialogue": "美咲(モノローグ): 誰もやらないなら、私が ─\n…時代を変える仕事になる気がする。", "sfx": "ドキッ"},
    {"num": 4, "size": "中", "composition": "美咲の手元から斜め見上げ", "visual":
     "美咲が手を挙げる。背景の他のメンバーが驚いて美咲を見る。",
     "dialogue": "美咲: …やります。", "sfx": "ピンッ"},
    {"num": 5, "size": "大", "composition": "CMOの満足げな笑顔と美咲の見つめあい", "visual":
     "CMOが嬉しそうに笑う。美咲は曖昧に微笑むが、目の奥には不安。",
     "dialogue": "田島CMO: お、山下くん。任せた!\n3ヶ月後の経営会議でAIO戦略レポートを出してくれ。",
     "sfx": "パチパチ\n(拍手)"},
])

add_heading(doc, "P.3 (帰り道、不安が芽生える)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "美咲の表情アップ・笑顔だが目は動揺", "visual":
     "美咲の顔のアップ。口は微笑んでいるが、目の中に動揺。汗の縦線。",
     "dialogue": "美咲(モノローグ): …手を挙げたはいいが、AEO? GEO? LLMO?\n何ひとつ詳しく知らない。", "sfx": "(汗ダラリ)"},
    {"num": 2, "size": "中", "composition": "デスクに座る美咲の引き", "visual":
     "自席に戻った美咲。肩が落ちている。PCの検索バーに『AEO とは』と入力中。",
     "dialogue": "", "sfx": "カタカタ\n(タイピング)"},
    {"num": 3, "size": "大", "composition": "PC画面とモノローグ", "visual":
     "PC画面に『AEO/GEO/LLMO』の概念説明ページが大量に表示。美咲が頭を抱える後ろ姿。",
     "dialogue": "美咲(モノローグ): これから3ヶ月、\nなにから手をつけよう…",
     "sfx": "(チチチ…時計の音)"},
])

doc.add_page_break()


# === 第2話 ===
add_heading(doc, "3. 第2話「教科書は買った。でも──」(全4ページ・20コマ想定)", level=1)

add_heading(doc, "P.1 (教科書との出会い)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "Amazonの注文画面", "visual":
     "PCのAmazon画面、『これからはじめるAIO AI最適化の教科書』(瀧内賢著、技術評論社、ピンク色のカバー)を注文する。",
     "dialogue": "美咲: 注文…!", "sfx": "ポチッ"},
    {"num": 2, "size": "大", "composition": "美咲の自宅リビング、夜", "visual":
     "週末の夜、リビングのソファで美咲が分厚いピンクの教科書を読んでいる。マグカップ、ノート、付箋。",
     "dialogue": "美咲(モノローグ): 一気読みしよう。\n一晩かけてでも。", "sfx": "(時計が深夜2時を示す)"},
    {"num": 3, "size": "中", "composition": "教科書のページの拡大", "visual":
     "教科書のページに『AIOピラミッド』『PREP→FAQ→HowTo』『引用率(Citation Rate)』の図解。",
     "dialogue": "美咲(モノローグ): 理屈は…\nわかった。", "sfx": "(ページめくる音)サラサラ"},
    {"num": 4, "size": "中", "composition": "美咲の表情・複雑", "visual":
     "美咲の顔、理解と迷いが混在。眉間に微かなしわ。",
     "dialogue": "美咲(モノローグ): でも、これを実際にやるのは別問題だ ─", "sfx": ""},
])

add_heading(doc, "P.2 (Spreadsheet地獄の始まり)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "大", "composition": "オフィスの俯瞰、PC画面に3つのSpreadsheet", "visual":
     "美咲のデスク、画面に『FAQ管理表』『FAQ記録表』『FAQ集計表』の3つのGoogle Sheetsが並ぶ。窓の外は朝。",
     "dialogue": "美咲: 教科書通り、まずは3つのシートを…", "sfx": "カタカタ"},
    {"num": 2, "size": "中", "composition": "PC画面アップ・列名打ち込み", "visual":
     "Spreadsheetに列名を1つ1つ手で打ち込む手元のアップ。タイトル、質問、回答、最終更新日、引用回数、引用率(%)、改善メモ。",
     "dialogue": "美咲: …これだけで、\n午前中が消えた。", "sfx": "ポチポチ\nカチカチ"},
    {"num": 3, "size": "中", "composition": "オフィスの時計", "visual":
     "オフィスの壁の時計、12:30を示す。窓の外の光が変わっている。",
     "dialogue": "", "sfx": "チクタク"},
    {"num": 4, "size": "中", "composition": "美咲がため息", "visual":
     "美咲がデスクで小さくため息。コンビニ弁当の蓋を開けながら。",
     "dialogue": "美咲(モノローグ): 教科書には『月1回12.5分』と書いてあった ─\nでも私は30本ある。", "sfx": "はぁ…"},
])

add_heading(doc, "P.3 (ペイン①:プロンプト選定の難しさ)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "営業の田中さんと立ち話", "visual":
     "オフィスの一角で営業の田中(30代男性)に話を聞く美咲。",
     "dialogue": "美咲: 田中さん、お客さんが最初によく聞く質問って?\n田中: えっと…料金とか? あとサポート体制かな…?", "sfx": ""},
    {"num": 2, "size": "中", "composition": "Notionで議事録を読み込む美咲", "visual":
     "PCにNotion画面、Win/Lossミーティング議事録100件のリスト。美咲がスクロールしている。",
     "dialogue": "美咲(モノローグ): 100件全部読むしかない。", "sfx": "スクロール"},
    {"num": 3, "size": "中", "composition": "夜のオフィス、デスクに付箋の山", "visual":
     "夜、デスクに大量の付箋(青・黄・ピンク)が貼られている。美咲は疲れた顔でノートに分類を書き出す。",
     "dialogue": "美咲(モノローグ): 3日かかった。\nこれが正しいのか、自信が持てない。", "sfx": ""},
    {"num": 4, "size": "中", "composition": "ペイン①強調枠", "visual":
     "コマ全面に赤枠の【ペイン①: プロンプトの正解SOPがどこにもない】の文字。背景に頭を抱える美咲のシルエット。",
     "dialogue": "(キャプション枠)\nペイン①: プロンプトの正解SOPがどこにもない", "sfx": ""},
])

add_heading(doc, "P.4 (ペイン②③:手動検索とLLM揺らぎ)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "大", "composition": "PC画面に4つのAIブラウザを並べた俯瞰", "visual":
     "Perplexity / ChatGPT / Google AI Overviews / Gemini の4つのタブを並列で開いている画面。",
     "dialogue": "美咲: 4つのAI、それぞれに30本×5回検索…\n4×30×5=600回。", "sfx": "(冷や汗)"},
    {"num": 2, "size": "中", "composition": "アップ・スクリーンショットを撮る画面", "visual":
     "PC画面にPerplexityの応答、Cmd+Shift+4のスクショキー、ファイル名『20260520_perplexity_query01_run3_no.png』",
     "dialogue": "", "sfx": "カシャッ"},
    {"num": 3, "size": "中", "composition": "オフィスの夜10時", "visual":
     "オフィスの暗い窓、時計が22:00を示す。美咲一人が机に向かう後ろ姿、PCの光が顔を照らす。",
     "dialogue": "美咲(モノローグ): これって、ダイエットの体重計を\n毎日手書きで管理してるのと同じじゃない?", "sfx": "(チカチカ蛍光灯)"},
    {"num": 4, "size": "中", "composition": "ペイン②強調枠", "visual":
     "赤枠の【ペイン②: 月25回×4AI×30プロンプト = 6時間超】",
     "dialogue": "(キャプション)\nペイン②: 物理的限界、毎晩夜10時", "sfx": ""},
    {"num": 5, "size": "中", "composition": "Perplexityの画面、月曜と火曜の比較", "visual":
     "上下2分割のコマ。上は月曜日のPerplexity画面で自社が3番目に引用される。下は火曜日に引用が消えている。",
     "dialogue": "", "sfx": "ガーン"},
    {"num": 6, "size": "中", "composition": "CMOへ報告で口ごもる美咲", "visual":
     "会議室で田島CMOに説明する美咲。CMOは怪訝な顔。",
     "dialogue": "美咲: 先週は出てたんですけど、今週は ─\nCMO: ……それ、どういうこと?", "sfx": "シーン"},
])

doc.add_page_break()


# === 第3話 ===
add_heading(doc, "4. 第3話「限界の夜と、出会い」(全4ページ・21コマ想定)", level=1)

add_heading(doc, "P.1 (スケール破綻)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "CMO会議室", "visual":
     "田島CMOが嬉しそうに紙を渡す。",
     "dialogue": "田島CMO: いい仕事だ! 来期、全社展開しよう。\n多言語(英・中・韓・タイ)、サブブランド10個、競合10社の分析も加えて。",
     "sfx": "(ガーン)"},
    {"num": 2, "size": "大", "composition": "美咲の机、モニター3台に膨大なSpreadsheet", "visual":
     "美咲のデスクに3つのモニターが並び、それぞれに大量のSpreadsheet。机の上はスクショの印刷物の山。",
     "dialogue": "美咲(モノローグ): 5月。プロンプトは300本、\n4ヶ国語、サブブランド10個、競合10社。\n1日4,000回以上の検索…",
     "sfx": "ガサガサ\n(紙の山)"},
    {"num": 3, "size": "中", "composition": "眼科の診察室", "visual":
     "眼科で診察を受ける美咲。眼科医が「ドライアイです」と告げる。",
     "dialogue": "眼科医: 山下さん、これはドライアイですね。\n美咲: …やっぱり。", "sfx": ""},
    {"num": 4, "size": "中", "composition": "ペイン⑤強調枠", "visual":
     "赤枠の【ペイン⑤: スケール破綻、毎日3時間の検索作業】",
     "dialogue": "(キャプション)\n肩こり・首痛・ドライアイ", "sfx": ""},
])

add_heading(doc, "P.2 (家族との時間が消える・最大の感情ピーク)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "玄関、葵が運動会の準備", "visual":
     "5月23日土曜日朝、4歳の葵が運動会用のリュックを背負っている。父・健一が一緒。美咲は仕事着姿でPC前。",
     "dialogue": "葵: ママ、運動会くるよね?\n美咲: ごめん…今日はお仕事。お父さんと行ってきて。", "sfx": ""},
    {"num": 2, "size": "中", "composition": "葵の泣きそうな顔アップ", "visual":
     "葵の小さい顔、目に涙がたまる。手にハンカチを握りしめる。",
     "dialogue": "葵: …うん。", "sfx": "(しーん)"},
    {"num": 3, "size": "大", "composition": "リビングのPC前、ひとり美咲", "visual":
     "玄関のドアが閉まる音。一人になったリビングで、美咲がPC前に座り直す。窓の外から運動会の応援の音。",
     "dialogue": "美咲(モノローグ): あと数時間で経営会議資料を作らないと…",
     "sfx": "バタンッ(ドア)\n遠くで「フレー、フレー」"},
    {"num": 4, "size": "中", "composition": "葵が運動会で1等賞", "visual":
     "夕方、健一が葵を連れて帰ってくる。葵は手にメダル。",
     "dialogue": "葵: ママ、今日リレーで1等賞だったよ!\n健一: 写真、撮ったよ。", "sfx": ""},
    {"num": 5, "size": "中", "composition": "美咲、写真を見つめる、目を伏せる", "visual":
     "美咲がスマホで写真を見る。葵の笑顔。美咲は微笑むが目を伏せる。心の中で何かが崩れる音。",
     "dialogue": "美咲: …すごいね。\n美咲(モノローグ): 私は、何やってるんだろう。",
     "sfx": "(コトン...心の音)"},
    {"num": 6, "size": "中", "composition": "ペイン⑧強調枠", "visual":
     "赤枠の【ペイン⑧: 家族との時間が消える】",
     "dialogue": "(キャプション)\n5月23日 土曜日", "sfx": ""},
])

add_heading(doc, "P.3 (退職を考えた金曜日・最大のクライマックス)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "大", "composition": "リビング、深夜1時20分", "visual":
     "5月29日金曜深夜1時20分。蛍光灯の下、PC前の美咲。一人。家族は寝ている。",
     "dialogue": "(時計のアップ: 01:20)", "sfx": "チカチカ\n(蛍光灯)"},
    {"num": 2, "size": "中", "composition": "PC画面・Perplexity認証エラー", "visual":
     "PC画面に『再ログインしてください』のエラー。美咲は深いため息。",
     "dialogue": "美咲: …また?", "sfx": "はぁ……"},
    {"num": 3, "size": "中", "composition": "LinkedInを開く美咲", "visual":
     "PC画面にLinkedIn、元同僚の昇進記事『AIマーケティング Sr. Manager 外資系IT』。",
     "dialogue": "美咲(モノローグ): あの子、転職して年収1.5倍…", "sfx": ""},
    {"num": 4, "size": "中", "composition": "検索バーに『退職届テンプレート』", "visual":
     "PC画面に『退職届 テンプレート』の検索結果。美咲の手がマウスを握る。",
     "dialogue": "", "sfx": "(マウスのクリック準備)"},
    {"num": 5, "size": "大", "composition": "美咲、涙", "visual":
     "美咲の顔のアップ、目に涙が浮かぶ。ノートPCを静かに閉じる。",
     "dialogue": "美咲(モノローグ): もう辞めたい…\nマーケティングって、こんなに虚しい仕事だったっけ ─", "sfx": "ポロッ\n(涙)"},
    {"num": 6, "size": "中", "composition": "ペイン⑨強調枠", "visual":
     "赤枠の【ペイン⑨: 退職を本気で考えた金曜日】",
     "dialogue": "(キャプション)\n5月29日 午前1時20分", "sfx": ""},
])

add_heading(doc, "P.4 (運命の出会い・転換点)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "大", "composition": "MarkeZine Day 2026会場", "visual":
     "業界カンファレンス『MarkeZine Day 2026』のメイン会場。多くの参加者。看板に大きくロゴ。",
     "dialogue": "(キャプション)\n6月3日 MarkeZine Day 2026", "sfx": "ザワザワ\n(会場の喧騒)"},
    {"num": 2, "size": "中", "composition": "美咲が疲れて帰る後ろ姿", "visual":
     "セッション終了後、美咲が会場出口へ向かう疲れた後ろ姿。",
     "dialogue": "美咲(モノローグ): もう帰ろう…", "sfx": ""},
    {"num": 3, "size": "大", "composition": "AIOutmationブースの看板アップ", "visual":
     "ブースの大型ポスター。『AIO最適化を、書く前から書いた後まで自動で回す。─ AIOutmation』の文字が大きく。",
     "dialogue": "(ポスターの文字)\nAIO最適化を、書く前から書いた後まで自動で回す。",
     "sfx": "ピカッ\n(光の効果)"},
    {"num": 4, "size": "中", "composition": "美咲、足が止まる", "visual":
     "美咲が立ち止まり、ポスターを見つめる横顔。目に光が戻る。",
     "dialogue": "美咲(モノローグ): …自動で?", "sfx": "ピタッ"},
    {"num": 5, "size": "中", "composition": "ブースの担当者と立ち話", "visual":
     "ブースの担当者(30代男性、ジャケット姿)が美咲に説明している。",
     "dialogue": "担当者: 月25回の手動検索は毎日4つのAIで自動実行されます。\n月次レポートは1分でホワイトラベル化されたPDFが出力されます。",
     "sfx": ""},
    {"num": 6, "size": "中", "composition": "美咲の表情・期待", "visual":
     "美咲の顔のアップ、目が輝く。久しぶりの希望の表情。",
     "dialogue": "美咲(モノローグ): …それ、本当?\n試してみよう。", "sfx": "キラッ"},
])

doc.add_page_break()


# === 第4話 ===
add_heading(doc, "5. 第4話「魔法のような半年間」(全3ページ・15コマ想定)", level=1)

add_heading(doc, "P.1 (オンボーディング体験)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "月曜の朝、美咲のデスク", "visual":
     "週明け月曜、美咲がオフィスでPCを開く。AIOutmationのログイン画面。窓の外は晴れ。",
     "dialogue": "美咲: …いざ。", "sfx": "カチッ"},
    {"num": 2, "size": "中", "composition": "オンボーディング画面・4つの質問", "visual":
     "PC画面に『1.自社URLは? 2.主要競合は? 3.対象AI? 4.顧客業種?』の4つだけ。",
     "dialogue": "美咲: たった4つの質問だけ?", "sfx": ""},
    {"num": 3, "size": "中", "composition": "5分後、画面に300本のプロンプト", "visual":
     "プロンプト一覧画面、タグ付き(購買ファネル/ペルソナ/プロダクト別)で300本がリスト。",
     "dialogue": "(画面: 300本のプロンプトが生成されました)", "sfx": "テン↑"},
    {"num": 4, "size": "大", "composition": "美咲の驚きの表情アップ", "visual":
     "美咲の顔のアップ、目を見開く、口が開く。",
     "dialogue": "美咲: …5分?\n3週間悩んでたものが、5分で…!?", "sfx": "ガーン!\n(良い意味で)"},
    {"num": 5, "size": "中", "composition": "心の声", "visual":
     "美咲がPC画面に手を当てる、感動。",
     "dialogue": "美咲(モノローグ): 私の3週間が、終わった ─",
     "sfx": "ジーン……"},
])

add_heading(doc, "P.2 (日常運用の風景)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "朝のSlack通知", "visual":
     "美咲のスマホ画面にSlack通知。『今日のCitation Rate: 全体42% (前日比+1.2pt)』",
     "dialogue": "美咲: 毎朝、ダッシュボードがコーヒー1杯で見られる。", "sfx": "ピコン!"},
    {"num": 2, "size": "中", "composition": "オフィスのデスク、リラックスした朝", "visual":
     "美咲がアイスカフェラテを片手にダッシュボードをチェック。穏やかな表情。",
     "dialogue": "美咲: VPNも、人格切替も、もういらない…", "sfx": "(穏やかなBGM)"},
    {"num": 3, "size": "中", "composition": "アラート通知の画面", "visual":
     "Slack画面に赤いアラート: 『競合A社が「勤怠管理 おすすめ」で過去3日に+20pt引用シェア急伸』。AIによる改善案が下に表示。",
     "dialogue": "美咲: 競合A社が動いた!", "sfx": "ピロン!"},
    {"num": 4, "size": "中", "composition": "ライターさんと協力", "visual":
     "美咲がライターのスタッフ(女性)と画面を見ながら相談。AIOutmationの改善案を一緒に確認。",
     "dialogue": "美咲: AIが改善案出してくれた! Schemaコードも添付。\nライター: これなら今週中に対応できますね!", "sfx": ""},
    {"num": 5, "size": "中", "composition": "1週間後、引用率逆転のグラフ", "visual":
     "PC画面にグラフ、競合と自社の引用率が逆転する。",
     "dialogue": "美咲: 逆転した!", "sfx": "わーい!"},
])

add_heading(doc, "P.3 (経営会議でCMOを驚かせる)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "経営会議室", "visual":
     "3ヶ月後の月次レビュー。CMOと美咲が向かい合う。CMOが3行サマリを見る。",
     "dialogue": "田島CMO: …今月の3行サマリ、いいね。\n来週、取締役会で報告したい。資料作ってくれ!", "sfx": ""},
    {"num": 2, "size": "大", "composition": "美咲がスマホで微笑む", "visual":
     "美咲がスマホでAIOutmationのレポートをすぐに開く、自信ある微笑み。",
     "dialogue": "美咲: …もうあります。", "sfx": "ニコッ"},
    {"num": 3, "size": "中", "composition": "CMOの驚き", "visual":
     "田島CMOが目を丸くする。スマホ画面を覗き込む。",
     "dialogue": "田島CMO: えっ…!?\n…素晴らしい!", "sfx": "ガビーン!"},
    {"num": 4, "size": "中", "composition": "10月14日、定時退社の美咲", "visual":
     "夕方6時の正面玄関。美咲が定時で帰る。スーツの軽快な歩き。",
     "dialogue": "美咲(モノローグ): 残業時間 月80時間 → 月20時間。\n久々に、葵にハンバーグ作れる。", "sfx": "コツコツ"},
    {"num": 5, "size": "大", "composition": "家族の食卓", "visual":
     "夕食、健一・美咲・葵の3人。葵がハンバーグを食べながら笑う。美咲も笑顔。",
     "dialogue": "葵: ママのハンバーグ、おいしい!\n美咲: ありがとう、葵。", "sfx": "(明るいBGM)"},
])

doc.add_page_break()


# === 第5話 ===
add_heading(doc, "6. 第5話「新しい役割」(全2ページ・8コマ想定)", level=1)

add_heading(doc, "P.1 (取締役会・昇進の瞬間)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "大", "composition": "12月の取締役会", "visual":
     "重厚な取締役会室、長机を囲む7-8人の役員。社長の篠崎(60代男性)が壇上で議題を読み上げる。",
     "dialogue": "(キャプション)\n12月、取締役会", "sfx": ""},
    {"num": 2, "size": "中", "composition": "社長の発表", "visual":
     "篠崎社長がマイクを手に発表。",
     "dialogue": "社長: 来期、当社は『マーケティングDX推進室』を新設する。\n室長は ─ 山下美咲。部長待遇だ。", "sfx": "ピンッ"},
    {"num": 3, "size": "中", "composition": "美咲、言葉に詰まる", "visual":
     "美咲のアップ、目が大きく見開かれ、口が少し開く。",
     "dialogue": "美咲: …えっ。", "sfx": "(ハッ)"},
    {"num": 4, "size": "中", "composition": "田島CMOが満足げにうなずく", "visual":
     "隣で田島CMOが満足げにうなずく。",
     "dialogue": "", "sfx": "(うんうん)"},
    {"num": 5, "size": "大", "composition": "美咲の感謝の表情", "visual":
     "美咲が立ち上がってマイクの前に。深く頭を下げる。",
     "dialogue": "美咲: …ありがとうございます。\n精一杯、務めます。", "sfx": "パチパチ\n(拍手)"},
])

add_heading(doc, "P.2 (帰り道・回想と決意)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "中", "composition": "オフィスの自販機の前", "visual":
     "美咲が自販機の前で、いつものアイスカフェラテを買う。窓の外は冬の夕焼け。",
     "dialogue": "", "sfx": "(ガッシャン!)"},
    {"num": 2, "size": "中", "composition": "回想・5月深夜の自分", "visual":
     "コマの一部に5月深夜のリビング、PC前で泣いていた自分の姿が回想として描かれる。",
     "dialogue": "美咲(モノローグ): …5月のあの夜。\n本当に辞めようと思った。", "sfx": "(モノクロでぼかし)"},
    {"num": 3, "size": "大", "composition": "夕焼けの中、美咲が前を向く", "visual":
     "オフィスビルのロビーから外に出る美咲。背景に冬の夕焼け。胸を張って前を向く。",
     "dialogue": "美咲(モノローグ): 本当に変わったのは、ツールじゃない。\n仕事の意味だった。", "sfx": ""},
])

doc.add_page_break()


# === エピローグ ===
add_heading(doc, "7. エピローグ「これからのマーケティング」(全1ページ・3コマ想定)", level=1)

add_heading(doc, "P.1 (最終メッセージ)", level=3)
add_panel_table(doc, [
    {"num": 1, "size": "大", "composition": "見開き全面・キャッチコピー", "visual":
     "白抜き背景に大きな文字。背景にうっすら美咲が空を見上げるシルエット。",
     "dialogue": "(大文字キャッチコピー)\n『手作業を、ソフトウェアに。\n戦略を、人間に。』",
     "sfx": ""},
    {"num": 2, "size": "中", "composition": "未来の美咲の後ろ姿", "visual":
     "オフィスで部下数名を率いる美咲の後ろ姿。マーケティングDX推進室の銘板。",
     "dialogue": "(キャプション)\nそれが、AIOutmationが信じている、これからのマーケティングのあり方です。",
     "sfx": ""},
    {"num": 3, "size": "中", "composition": "ロゴと締め", "visual":
     "AIOutmationのロゴ、Webサイト/QRコード、お問い合わせ情報。",
     "dialogue": "もし、あなたが今日『AIO担当』に任命されたら ─\n最初に、AIOutmationを開いてください。",
     "sfx": ""},
])

doc.add_page_break()


# === 制作Tips ===
add_heading(doc, "8. 漫画家・イラストレーターへの指示Tips", level=1)

add_heading(doc, "8.1 各話のキー感情(描き分けのポイント)", level=2)
t = doc.add_table(rows=7, cols=3)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(2.5)
    r.cells[1].width = Cm(7.5)
    r.cells[2].width = Cm(8.0)

write_cell(t.cell(0, 0), "話", size=10, bold=True, bg="1F3A5F", color="FFFFFF")
write_cell(t.cell(0, 1), "美咲のキー感情", size=10, bold=True, bg="1F3A5F", color="FFFFFF")
write_cell(t.cell(0, 2), "描画ポイント", size=10, bold=True, bg="1F3A5F", color="FFFFFF")

emotions = [
    ("第1話", "好奇心+不安(60:40)", "明るい瞳、口元微笑み、しかし眉は少し下がっている。CMOの発表→沈黙→挙手の流れにメリハリ。"),
    ("第2話", "戸惑い→疲労(段々暗くなる)", "前半は集中した真剣な目。後半に向けて目の下にうっすらクマ、肩が落ちる。"),
    ("第3話", "絶望→希望(後半に転換)", "前半は深い疲労、目元の暗さ、涙、家族の写真を見つめる悲しみ。後半でブースを見つけたとき目に光が戻る。"),
    ("第4話", "驚き→喜び→自信", "目を大きく見開く、笑顔が増える、姿勢が良くなる、家族の食卓では完全な笑顔。"),
    ("第5話", "感慨深さ・成熟", "穏やかで深い表情、目に涙、しかし前を向く。回想シーンとの対比が重要。"),
    ("エピローグ", "決意・前向き", "強い視線、空を見上げる構図、希望の象徴(夕焼け・朝日)。"),
]
for i, (ep, emo, point) in enumerate(emotions):
    write_cell(t.cell(i+1, 0), ep, size=10, bold=True, bg="E8EEF7")
    write_cell(t.cell(i+1, 1), emo, size=10)
    write_cell(t.cell(i+1, 2), point, size=10)

add_heading(doc, "8.2 視覚的テーマ(チャプター別の色温度)", level=2)
add_paragraph(doc, "・第1話: ニュートラル(普通の朝)。トーン控えめ。", size=10)
add_paragraph(doc, "・第2話: 徐々に冷たい(青み)。蛍光灯の硬い光。", size=10)
add_paragraph(doc, "・第3話: 寒色全開、夜の暗さ。蛍光灯チカチカ。スクリーントーン多用。", size=10)
add_paragraph(doc, "・第4話: 暖色に転換。朝日、自然光、笑顔。トーン控えめで明るく。", size=10)
add_paragraph(doc, "・第5話: 夕焼けのオレンジ。成熟した暖かい色味。", size=10)

add_heading(doc, "8.3 重要な視覚的シンボル", level=2)
add_paragraph(doc, "・アイスカフェラテのプラカップ: 美咲の習慣的アイテム。第1話と第5話で同じカップが映ることで時間経過を表現。", size=10)
add_paragraph(doc, "・教科書(ピンクのカバー): 第2話で初登場、第5話で本棚に整然と並ぶ姿で再登場(成熟の暗示)。", size=10)
add_paragraph(doc, "・蛍光灯のチカチカ: 第3話の絶望シーンで使用、第4話以降は登場しない。", size=10)
add_paragraph(doc, "・葵のメダル: 運動会の1等賞メダルが第4話で家族の食卓に飾られている描写があると◎", size=10)
add_paragraph(doc, "・PC/スマホ画面: AIOutmationのUIは実際のSaaSダッシュボードを意識(濃紺ベース、緑のCitation Rateグラフ)。", size=10)

add_heading(doc, "8.4 NG/避けたい表現", level=2)
add_paragraph(doc, "× 過度なデフォルメ: SDキャラやちびキャラ化はビジネス漫画には不適。", size=10)
add_paragraph(doc, "× 過剰な美化: 美咲は普通の働く女性。アイドル系の描写は避ける。", size=10)
add_paragraph(doc, "× 競合製品の実名描写: ChatGPT等は『AI検索ツール』『AIアシスタント』として描き、ロゴ無断使用は避ける。", size=10)
add_paragraph(doc, "× 過度なPR感: 第4話以降のAIOutmation登場場面でも『広告色』を強調しすぎず、自然なソリューション発見の流れに。", size=10)

add_heading(doc, "8.5 発注時の追加情報リクエスト", level=2)
add_paragraph(doc, "イラストレーターさんへの発注時、追加で以下を提示推奨:", size=10)
add_paragraph(doc, "・実際の前作(過去のWord版『AIOutmation_Story.docx』)を共有して文脈理解を深める", size=10)
add_paragraph(doc, "・AIOutmationの実画面イメージ(リーンキャンバス、MVP仕様書からダッシュボードイメージ)", size=10)
add_paragraph(doc, "・参考になる既存ビジネス漫画のサンプル(『ドラゴン桜』『重版出来!』等)", size=10)
add_paragraph(doc, "・キャラクター3面図(正面・側面・後ろ姿)を初稿で受領", size=10)
add_paragraph(doc, "・第1話の3ページネーム化を初回納品とし、トーン確認後に残りへ進める段階発注", size=10)

add_heading(doc, "8.6 想定発注規模・予算目安", level=2)
add_paragraph(doc, "・ページ単価: 漫画家のレベルにより1ページ3〜10万円(プロ級は10〜30万円)", size=10)
add_paragraph(doc, "・全20ページ想定: 60万円〜600万円のレンジ", size=10)
add_paragraph(doc, "・推奨アプローチ: まず第1話(3ページ)を試し発注し、品質確認後に残り発注", size=10)
add_paragraph(doc, "・期間: フルカラー20ページで2〜4ヶ月。モノクロなら1〜2ヶ月。", size=10)

add_heading(doc, "8.7 発注先候補(参考)", level=2)
add_paragraph(doc, "・ココナラ・ランサーズ等のクラウドソーシング: 1ページ1〜5万円、品質バラつき", size=10)
add_paragraph(doc, "・漫画制作会社(マンガぼっくす、コルク等): プロ品質、1ページ10万円〜", size=10)
add_paragraph(doc, "・個人漫画家への直接依頼: TwitterやpixivでBtoB漫画実績ある人を探す", size=10)
add_paragraph(doc, "・営業漫画専門会社(ストーリーフィット等): ビジネス漫画特化、企画から相談可", size=10)


add_paragraph(doc, "", size=8)
add_paragraph(doc, "─ 本資料はイラストレーター/漫画家への発注用指示書です。実際の作画は受注先にお任せください。", size=9, color="888888")

doc.save("/home/user/AIO/AIOutmation_Manga_Storyboard.docx")
print("Saved: AIOutmation_Manga_Storyboard.docx")
