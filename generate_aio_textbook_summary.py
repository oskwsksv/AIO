"""Generate Word summaries (50p detailed + 10p brief) of the AIO textbook.

Source: 『これからはじめるAIO AI最適化の教科書』瀧内賢著(技術評論社, 2026年3月)
Structure: Problem → Solution / Goal → Means logic mapping
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============== Helpers ==============

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


def write_cell(cell, text, size=9, bold=False, bg=None, color=None):
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
        style_run(run, size=size, bold=bold, color=color)


def add_heading(doc, text, level=1, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level <= 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    sizes = {0: 22, 1: 18, 2: 14, 3: 12, 4: 11}
    default_colors = {0: "1F3A5F", 1: "1F3A5F", 2: "2E5C8A", 3: "444444", 4: "555555"}
    style_run(run, size=sizes.get(level, 11), bold=True, color=color or default_colors.get(level, "000000"))


def add_paragraph(doc, text, size=10, bold=False, color=None, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(indent * 0.5) if indent else None
    run = p.add_run(text)
    style_run(run, size=size, bold=bold, color=color)
    return p


def add_bullet(doc, text, size=10, indent=0):
    p = doc.add_paragraph(style="List Bullet")
    if indent:
        p.paragraph_format.left_indent = Cm(0.5 + indent * 0.5)
    run = p.add_run(text)
    style_run(run, size=size)


def add_callout(doc, title, body, color="FFF4E0"):
    """Create a single-cell highlighted box."""
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, color)
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(3)
    r = p1.add_run(title)
    style_run(r, size=11, bold=True, color="C04000")
    p2 = cell.add_paragraph()
    r = p2.add_run(body)
    style_run(r, size=10)


def add_two_col_table(doc, header_left, header_right, rows, widths=(5.0, 12.0)):
    """Two-column table with headers."""
    t = doc.add_table(rows=len(rows) + 1, cols=2)
    t.autofit = False
    for r in t.rows:
        r.cells[0].width = Cm(widths[0])
        r.cells[1].width = Cm(widths[1])
    write_cell(t.cell(0, 0), header_left, size=10, bold=True, bg="1F3A5F", color="FFFFFF")
    write_cell(t.cell(0, 1), header_right, size=10, bold=True, bg="1F3A5F", color="FFFFFF")
    for i, (l, r) in enumerate(rows):
        write_cell(t.cell(i + 1, 0), l, size=10, bold=True, bg="E8EEF7")
        write_cell(t.cell(i + 1, 1), r, size=10)


def add_n_col_table(doc, headers, rows, widths):
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


def setup_doc(doc, landscape=False):
    section = doc.sections[0]
    if landscape:
        new_w, new_h = section.page_height, section.page_width
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = new_w
        section.page_height = new_h
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


# ============== Content Components ==============

def cover_page(doc, title, subtitle, version_label):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(120)
    r = p.add_run(title)
    style_run(r, size=28, bold=True, color="1F3A5F")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    r = p.add_run(subtitle)
    style_run(r, size=16, bold=True, color="2E5C8A")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(50)
    r = p.add_run("出典:『これからはじめる AIO AI最適化の教科書』")
    style_run(r, size=13, color="555555")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("瀧内賢 著／技術評論社／2026年3月発行/全464ページ")
    style_run(r, size=11, color="555555")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(40)
    r = p.add_run(f"【{version_label}】")
    style_run(r, size=14, bold=True, color="C04000")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(180)
    r = p.add_run("作成: 2026年5月")
    style_run(r, size=10, color="888888")

    doc.add_page_break()


# ============== Shared content data ==============

EXEC_SUMMARY = """検索の主役が"人"から"AI"へ移行する中、Webサイトの役割は「見つけてもらう」から「AIに選ばれる」へと根本的に変わりつつある。本書は、SEOの延長線上にある単一の技術論ではなく、AEO(回答最適化)・GEO(引用最適化)・LLMO(学習対象化最適化)を統合した設計思想=AIOを、実装手順とともに体系化した実践書である。

著者・瀧内賢氏(SEO書籍7作・生成AI書籍8作の実績)は、SEO/AIOの両領域を長く実務で扱ってきた専門家として、抽象的な概念論ではなく、PREP法・FAQ構文・JSON-LDスキーマ・PDCA運用といった日々の業務に落とし込める形で示している。本サマリーでは、書籍全464ページの「課題→打ち手」構造を抽出し、実装ロードマップとして整理する。"""

KEY_FRAMEWORKS = """本書の中核となる4つのフレームワーク:

1. AIO ピラミッド: SEO(基礎構造)→AEO(理解)→GEO(引用)→LLMO(学習対象化)→運用(PDCA)を「土台→4層→循環」として捉える設計モデル。家屋建築のメタファーで、土台から順に積み上げる構造。

2. AEO/GEO/LLMO 三層構造: AEOは「AIに理解させる」、GEOは「AIに引用させる」、LLMOは「AIに学習対象として採用させる」。短期(AEO/GEO 3〜12ヶ月)→長期(LLMO 1年〜)の時間軸を持つ連動施策。

3. Query→Intent→Answer モデル: AIが質問を理解する内部プロセス。記事側でも見出し(Query)で意図を明示し、回答(Answer)を三層構造(PREP原理層→FAQ深化層→HowTo具体層)で提供する。

4. E-T-R 再定義: 従来のE-E-A-T(経験・専門性・権威性・信頼性)を、AI時代向けに「Evidence(根拠)×Traceability(追跡性)×Retention(定着性)」の3軸へ再構成。AIが信頼を判定するための新しい基準。"""

CHAPTER_STRUCTURE = [
    ("Chapter 0", "本書を読みはじめる前に", "AIO学習の前提知識・背景"),
    ("Chapter 1", "AIO「AI最適化」の基本を知る", "SEO→AIOの変化、三層構造、Query→Intent→Answer、E-T-R、AIOピラミッド"),
    ("Chapter 2", "AIが理解しやすい「文章構成」を作る", "PREP法、FAQ構文(Q&A)、HowTo構文、meta+llms.txt+著者情報、E-E-A-T可視化"),
    ("Chapter 3", "AIが引用しやすい「文書構造」を設計する", "HTML階層、FAQPageスキーマ、Evidence+Reference、SEO/AIO連携、HowToスキーマ"),
    ("Chapter 4", "AIに選ばれ続けるためのPDCAサイクル", "PDCA全体像と引用率、定期確認・スプレッドシート管理、ページ別分析、運用スケジュール、トラブル対処"),
    ("Chapter 5", "AIに学ばれる情報資産を育てる", "Web全体検索型AIの理解、llms.txt、AI別最適化、更新頻度・一貫性、Evidence/Reference深化"),
]

# ============== Chapter contents (problem-solution structured) ==============

CH1_CONTENT = {
    "title": "Chapter 1: AIO「AI最適化」の基本を知る",
    "problem": "Googleで1位を目指してきたWebサイトが、AIに引用されないと「存在しない」のと同じになる時代に、何をどう対策すればよいか不明瞭。AEO/GEO/LLMOの個別解説はあっても、統合された設計思想がない。",
    "solution": "AIO=AEO+GEO+LLMO+SEOの統合的設計思想として再定義し、AIOピラミッドという階層モデル(土台→4層→循環)で全体像を提示する。",
    "sections": [
        {
            "title": "1-1: SEOからAIOへの変化を理解する",
            "goal": "「検索順位」ではなく「AIに選ばれる」を新しいKPIとして認識する",
            "means": "SEO(検索エンジン向け)とAIO(AI向け)の評価軸の違いを整理。AIは『信頼できる理由』(Evidence)を重視し、ユーザーが自分で探す→AIが選んでくれるへの構造変化を理解する。",
        },
        {
            "title": "1-2: AEO/GEO/LLMOの三層構造モデリング",
            "goal": "AI最適化の3つの異なる施策の役割と関係性を明確に把握する",
            "means": """AEO(Answer Engine Optimization): AIにサイト内容を理解させ、回答パーツとして抽出しやすくする最適化。FAQ構文/Q&A構造、PREP法、見出し設計が中心。
GEO(Generative Engine Optimization): AIに信頼できる情報源として引用させ、高品質な回答生成に貢献させる最適化。QCEP構文(Question/Conclusion/Evidence/Publisher)、根拠と出典の明示が中心。短中期施策(3-12ヶ月)。
LLMO(Large Language Model Optimization): AIモデルの将来更新時に学習対象として採用される可能性を高める長期施策。継続的な引用蓄積、更新頻度、構造化された知識資産化が中心。長期施策(1年〜)。""",
        },
        {
            "title": "1-3: Query→Intent→Answer がAIの理解構造",
            "goal": "AIがコンテンツを理解する内部プロセスに合わせて、コンテンツ側を設計する",
            "means": """Query(質問)→Intent(意図)→Answer(回答)という3ステップでAIは情報を解釈する。コンテンツ側もこの流れに合わせ:
・見出し: 質問形式(Query)で意図を明示
・回答: 三層構造で提供
  レイヤー1(PREP法): 原理・本質層 - 結論+理由を簡潔に
  レイヤー2(FAQ): 例外・補足層 - Q&A形式で深化
  レイヤー3(HowTo): 実践・手順層 - 具体的アクションへ
・原則: 下層は上層を前提とする、各層は独立した価値、必要な層のみ生成""",
        },
        {
            "title": "1-4: E-T-R再定義 - AI時代の新しい信頼軸",
            "goal": "従来E-E-A-T(経験・専門性・権威性・信頼性)に加え、AIが評価する新軸を実装する",
            "means": """E-T-Rの3軸:
・E(Evidence/根拠): 数値データ・調査結果・実績を具体的に提示
・T(Traceability/追跡性): 出典URL・更新日・著者情報で追跡可能にする
・R(Retention/定着性): 継続的な引用と長期的な信頼の蓄積で「定着」する状態
E-E-A-Tは「信頼を伝える表現方法」、E-T-Rは「AIが信頼を判断する基準」。両者は補完関係。""",
        },
        {
            "title": "1-5: AIOピラミッドの循環構造",
            "goal": "施策の優先順位と相互関係を1枚の図で把握する",
            "means": """ピラミッド構造(下から順):
・土台層(SEO): タイトル最適化、内部リンク、構造化データ、E-E-A-T
・1層目(AEO/理解層): PREP構成、FAQ/Q&A構造、見出し構造
・2層目(GEO/引用層): QCEP構成、比較構文、統計データ引用、権威ある出典
・3層目(LLMO/学習対象化層): 独立した意味単位の段落、HTML構造化、更新日・著者情報・出典URL
・最上層(運用): PDCA、アクセス解析、AI引用率測定、定期更新
特徴: 下が無いと上は機能しない/各層に主担当技術/サポート技術が補完/最上層から再び1層目に戻る循環""",
        },
    ],
}

CH2_CONTENT = {
    "title": "Chapter 2: AIが理解しやすい「文章構成」を作る",
    "problem": "AIに「読まれる」だけでは引用されない。AIは構造を読み取って意味を理解するため、ただの良い文章ではなく『AIに役割が認識される構造化された文章』が必要。",
    "solution": "PREP法→FAQ構文→HowTo構文の3つの『型』を1ページ内で組み合わせる。さらにmeta情報+llms.txt+著者情報で入口の文脈を与え、E-E-A-Tテンプレートで信頼性を可視化する。",
    "sections": [
        {
            "title": "2-1: PREP→FAQ→HowToで文章構成を整える",
            "goal": "1ページ内でAIに『理解→引用→学習対象化』の流れを提供する",
            "means": """PREP法: 段落全体の主張を論理的に説明(AEO主担当)
FAQ構文: 想定される疑問を軸に情報を配置(AEO+GEO+LLMO)
HowTo構文: 行動の手順を順序立てて説明(AEO+GEO+LLMO)
配置順: PREP(全体像)→FAQ(個別の疑問)→HowTo(手順)
※HowToは「やり方・手順」を求められる場合のみ。「なぜ?」「どれがいい?」「違いは?」「意味は?」にはPREP+FAQで十分""",
        },
        {
            "title": "2-2: PREP+FAQ(Q&A)で引用される最小単位を設計する",
            "goal": "AIが部分引用しても文脈が崩れない『独立した意味単位』を作る",
            "means": """PREP法の4段階:
・P(Point/結論): 「それは○○です」 - 冒頭で結論を明示しAIの記憶に残す
・R(Reason/理由): 「なぜなら○○だからです」 - 因果関係を明確化
・E(Example/具体例): 「例えば○○です」 - 数値や例で抽象を具体化
・P(Point/結論再強調): 「だから○○なのです」 - 末尾で再度結論
FAQ構文(Q&A構造):
・Question: 読者が抱く疑問を質問形式で明示
・Answer: 簡潔で完結した回答(必要に応じてPREP法で構成)
・補足(任意): 数値データや具体例で補強
HTML実装: dl/dt/dd タグで質問と回答の対応関係を構造的に明示""",
        },
        {
            "title": "2-3: FAQ+HowToで知識として採用してもらう",
            "goal": "実用的な手順情報を、AIが学習対象として採用しやすい形式で提供",
            "means": """HowTo構文: 「何かを実行する具体的手順」を段階的に記述
・Step1, Step2, Step3...と番号付きで明確化
・各ステップに必要な前提条件を記載
・成果物・完了状態を明示
FAQと組み合わせる効果:
・FAQ「どうすれば○○できますか?」→ HowToへリンクすることで、質問に対する完全な回答セットを提供
・AIは「このページは体系的に手順を持つ知識源である」と認識し、学習対象として優先的に採用""",
        },
        {
            "title": "2-4: meta情報+llms.txt+著者情報で最初の文脈を与える",
            "goal": "本文に入る前の『入口設計』でAIに文脈を与える",
            "means": """meta情報(ページ単位の名札):
・title: 25文字前後、ブランド名+地域+業種+強みを含める
・description: 100-110文字、結論→理由→効果の順
・OGP(og:title/og:description/og:image): SNS/AIが主題をつかみやすく
llms.txt(サイト全体の取り扱いルール):
・優先的に読んで欲しいページのリスト
・ルートディレクトリに配置 (https://example.com/llms.txt)
・Markdown形式で記述、AIに重要ページを案内
著者情報(知識の責任主体):
・著者の経歴・専門分野・肩書を構造化データとして整理
・schema.org Person型で記述""",
        },
        {
            "title": "2-5: E-E-A-Tで信頼を可視化する",
            "goal": "AIが『信頼』を機械的に読み取れる形にE-E-A-Tを変換する",
            "means": """個人署名(著者の肩書・経歴・専門分野・更新日):
・footer内に画像+経歴+組織+更新日を記載
・写真があることで「実在する人物が書いている」と判断
構造署名(FAQ内の監修者・出典明記):
・FAQ回答内に「監修: 山本花菜」「出典: 公式サイト」を含める
・citeタグで「引用・出典」のHTML意味構造を伝達
E-E-A-Tテンプレート整備:
・経験(Experience): 実績データ、一次情報の痕跡、定型句「〜してきた経験から」
・専門性(Expertise): 資格・専門用語+平易な説明
・権威性(Authoritativeness): 社名・メディア掲載リンク・外部評価
・信頼性(Trust): 連絡先・更新日・監修者・出典タグ
文体統一: 同サイト内で文末・敬体・単語を統一しないとAIは『運営方針が一貫していない』と低評価""",
        },
    ],
}

CH3_CONTENT = {
    "title": "Chapter 3: AIが引用しやすい「文書構造」を設計する",
    "problem": "Chapter 2で作った文章をAIが正しく解釈できなければ引用に至らない。HTMLレベルの構造設計が無いと、同じ内容でもAIの拾われ方が変動する。",
    "solution": "HTML階層+JSON-LDスキーマ+Evidence/Reference+SEO/AIO連携の4階層で文書構造を設計し、AIの解釈を「推論依存」から「定義情報」へ転換する。",
    "sections": [
        {
            "title": "3-1: HTML構造を整理する",
            "goal": "AIがページの論理構造を正確に把握できるようにする",
            "means": """h-タグ階層の基本ルール:
・h1: ページ全体のメインテーマ(原則1ページに1つ)
・h2: 大見出し(主要セクションの区切り)
・h3: 中見出し(h2配下の詳細項目)
・原則: h1→h2→h3 の順序を守る、階層を飛ばさない、見出しの直後に必要に応じてpタグで補足を加える
dlタグでFAQ構文を実装:
・dl(コンテナ)/dt(質問)/dd(回答)
sectionタグで意味的区切り:
・<section class="faq">, <section class="evidence"> など
・AIはセクションごとの目的を理解できる""",
        },
        {
            "title": "3-2: FAQPageスキーマ(JSON-LD)を実装する",
            "goal": "FAQ情報を「定義情報」としてAIに明示する",
            "means": """JSON-LDによる構造化データ:
・schema.orgのFAQPageスキーマを使用
・mainEntity, name(質問), acceptedAnswer, text(回答)の階層を実装
・HTMLで設計した「意味」を、解釈不要な形でAIに伝達
効果:
・検索結果でのリッチスニペット表示(SEO効果)
・AIによる引用の優先度向上(GEO効果)
・「このページは構造化された知識源」と認識(LLMO効果)
注意点: HTMLとJSON-LDの内容は完全一致させる(齟齬があるとペナルティ)""",
        },
        {
            "title": "3-3: Evidence+Referenceを追加する",
            "goal": "AIが「信頼できる」と判断する根拠を明示する",
            "means": """Evidence層(自社の根拠):
・数値データ・調査結果・実績(自社調査・公式発表)
・section class="evidence"で囲み、構造的に分離
・例: 「顧客満足度96% (2025年10月当社調べ)」
Reference層(外部の権威):
・第三者機関・公的データ・業界団体への外部リンク
・citeタグで引用・出典であることを明示
・例: <cite>総務省 情報通信白書(2025年版)</cite>
両者の組み合わせ:
・自社Evidence(独自性) + 外部Reference(信頼性)で、AIは「客観性のある一次情報源」と判断""",
        },
        {
            "title": "3-4: SEOとAIOを連携させる",
            "goal": "検索エンジンとAIの両方に発見・評価される状態を作る",
            "means": """canonicalタグ:
・重複コンテンツを防ぎ、正規版URLを明示
・AIは「このページの正式版はどこか」を判別
内部リンク戦略:
・関連FAQをつなげ、トピッククラスター化
・Evidence/Reference層への内部リンクで信頼性ネットワークを形成
title/meta description/h1の整合:
・タイトルとh1とmeta descriptionの内容を揃える
・AIが「タイトルと内容に矛盾がない」と判断""",
        },
        {
            "title": "3-5: HowToスキーマ(JSON-LD)を実装する",
            "goal": "手順情報を構造化し、AIが手順として認識できるようにする",
            "means": """schema.orgのHowToスキーマ:
・name(手順名), step(各ステップ), text(説明)の階層
・各stepにposition番号を付与し順序を明示
効果:
・「Aするには?」のような手順クエリで優先引用
・Google検索結果のリッチスニペットでステップ表示
完成形: FAQPage + HowTo の両スキーマで、ページが「Q&A+手順」の完全な知識源として認識される""",
        },
    ],
}

CH4_CONTENT = {
    "title": "Chapter 4: AIに選ばれ続けるためのPDCAサイクル",
    "problem": "AIに引用される文書を一度作っても、AIアルゴリズムや競合変動で評価は変動する。継続的な測定・改善が無いと「作って放置」になり、引用率が低下する。",
    "solution": "Citation Rate(引用率)という独自指標を定義し、PDCAサイクルをGoogle Sheetsで運用管理する。月10-15分の作業で継続的にAI引用を最適化する。",
    "sections": [
        {
            "title": "4-1: 運用管理の全体像とFAQのPDCAサイクル",
            "goal": "AIに「選ばれ続ける」状態を継続運用で実現する",
            "means": """PDCAの4段階:
・Plan(計画): FAQ管理表で「今あるFAQ」を見える化、優先順位整理
・Do(実行): 新規FAQの構造化、または前回Actで計画した改善を実装
・Check(確認): AI引用状況を確認、引用率を測定
・Act(改善): Checkで発見した問題を分析、次回Doへの改善計画を立案
引用率(Citation Rate): 本書独自指標
・計算式: (引用回数 ÷ チェック総数) × 100
・目標値: 維持目標60%、理想目標70%、優先改善対象50%未満
・例: FAQ1: 引用回数5回、チェック総数7回 → 引用率71%
3つの成功要素:
・定期的な測定の習慣化(月10-15分)
・データの可視化(色分けで改善対象を明確化)
・小さな改善の積み重ね(数値追加、更新日明記、画像追加)""",
        },
        {
            "title": "4-2: AI引用状況の定期確認とスプレッドシート管理",
            "goal": "毎月10-15分で継続可能な確認・記録の仕組みを構築する",
            "means": """確認に使うAI 3種:
・Perplexity(最優先): 出典URLが番号付きで明確に表示される
・Google AI Overviews: Google検索結果上部に表示
・ChatGPT(検索機能付き): 参考程度
確認頻度の推奨:
・初期(公開後1-2ヶ月): 週1回×各FAQ3回ずつ検索
・定期運用(3ヶ月以降): 月1回×各FAQ5回ずつ検索
・1日に集中して実行(条件統一のため)
2種類のスプレッドシート:
・FAQ記録表(生データ): 確認のたびに1行追加。確認日/AI/検索クエリ/FAQNo./引用された?/メモ
・FAQ集計表(月末作成): 1FAQに1行のサマリー。引用回数/チェック総数/引用率(%)/判定
「引用」の定義: AIの回答画面に出典(Sources)として自社URLが表示されている状態。回答内容が似ていても出典表示が無ければ「引用された」とは数えない。""",
        },
        {
            "title": "4-3: ページ別・セクション別パフォーマンス分析",
            "goal": "FAQ単位、HowTo単位、Evidence単位で性能を分析し改善対象を特定",
            "means": """ページ別ヒートマップ作成:
・縦軸: ページタイトル、横軸: 月別引用率
・色分け: 70%↑緑 / 60-70%黄 / 50-60%橙 / 50%↓赤
セクション別分析:
・FAQ単位: どのQ&Aが引用率高いか
・HowTo単位: 手順情報が引用されているか
・Evidence単位: 数値根拠の引用パターンを把握
原因切り分け:
・AEO問題(理解されていない): 構造改修(h階層、PREP、FAQ構文)
・GEO問題(信頼度不足): Evidence追加、出典明記
・LLMO問題(継続性不足): 更新頻度、関連ページ整備""",
        },
        {
            "title": "4-4: 運用スケジュールとチェックリスト管理",
            "goal": "週次・月次・四半期・年次の確認フローをテンプレート化",
            "means": """週次確認(初期2ヶ月のみ):
・Top10プロンプトのみ確認
・15分程度
月次確認(定期運用):
・全FAQ×5回検索 = 約25回×30秒 = 12.5分
・FAQ集計表を更新、ヒートマップ更新
・引用率50%未満のFAQを特定し改善計画
四半期レビュー:
・Chapter1で学んだAI引用全体像と照合
・新規プロンプト・FAQの追加、不要なFAQの統合・削除
・競合の被引用ページ調査
年次戦略会議:
・AEO/GEO/LLMOの再評価
・llms.txt見直し、E-E-A-Tテンプレ更新""",
        },
        {
            "title": "4-5: 運用でよくあるトラブルと対処法",
            "goal": "頻出トラブルを事前に把握し、迅速に対応できるようにする",
            "means": """トラブル例と対処:
・引用率が50%未満で改善しない: 文章構成(PREP/FAQ)が壊れている可能性。Chapter2に戻って構造再確認
・特定モデルで引用されない(例: Geminiだけ低い): モデル別最適化(Chapter5の5-3参照)
・引用が消えた: 競合がより強い情報を出した、自社情報が古くなった、HTMLが破壊された等を順に確認
・スプレッドシート運用が続かない: Google Apps Scriptで一部自動化、または専用ツール導入を検討""",
        },
    ],
}

CH5_CONTENT = {
    "title": "Chapter 5: AIに学ばれる情報資産を育てる",
    "problem": "短期の引用獲得だけでは、将来のAIモデル更新時に淘汰される可能性がある。「一時的な引用」ではなく「長期的な学習対象」になるための設計が必要。",
    "solution": "Web全体検索型AIの内部処理を理解し、llms.txtで重要ページを案内、AI別(Perplexity/Gemini/ChatGPT)に最適化、Evidence/Referenceを継続深化することで知識資産化を実現。",
    "sections": [
        {
            "title": "5-1: Web全体検索型AIの処理理解と引用構造分析",
            "goal": "AIが回答を生成する内部プロセス(RAG)を理解する",
            "means": """検索+生成の4ステップ:
・Retrieve(検索): キーワード抽出、関連ページ収集、robots.txt/llms.txt参照
・Rank(順位付け): 信頼性・鮮度・関連度で評価し上位コンテキスト選定
・Summarize(要約・統合): 選ばれた情報を元に回答文を生成・要約
・Reference(出典化): 各文の根拠ページURLを出典として紐づけ
従来AI vs Web全体検索型AI:
・従来: 学習データのみ・最新情報対応不可・出典なし
・Web検索型: 事前学習+リアルタイムWeb・最新対応・出典あり
発見条件3つ:
・条件1: Googleインデックスへの登録(土台)
・条件2: llms.txtの設置(重要ページ案内)
・条件3: robots.txtでAIクローラー(GPTBot,Google-Extended等)を許可""",
        },
        {
            "title": "5-2: llms.txtの設定",
            "goal": "AIに「このサイトのどのページを優先的に参照してほしいか」を明示",
            "means": """llms.txt配置:
・サイトルート(https://example.com/llms.txt)に配置
・Markdown形式で記述
基本フォーマット:
# サイト名
> サイトの簡潔な説明文(AI向け)
## メインコンテンツ
- [ページ名](URL): ページの説明
## サブコンテンツ
- [ページ名](URL): ページの説明
ベストプラクティス:
・FAQページや主要なナレッジページを優先記載
・更新頻度の高いページを明示
・robots.txtでクロール許可も併せて設定""",
        },
        {
            "title": "5-3: AI別の最適化戦略",
            "goal": "Perplexity, Gemini, ChatGPTそれぞれの特性に合わせて最適化",
            "means": """Perplexity:
・自社インデックス+外部検索API(Google/Bing)
・出典URLを必ず番号付き表示
・対策: Evidence重視、最新情報、独自データを優先
Gemini(Google AI Overviews):
・Googleインデックスから候補取得、複数ソースを要約
・対策: SEO強化、JSON-LD実装、構造化データ重視
ChatGPT(検索機能付き):
・Bing検索API+モデル訓練データ組み合わせ
・対策: ChatGPTの過去学習データに含まれる情報源として確立を狙う(LLMO的視点)
共通対策:
・FAQPage/HowToスキーマで構造化
・Author情報の構造化(Person型)
・更新日の明示(dateModified)""",
        },
        {
            "title": "5-4: 更新頻度と一貫性の維持戦略",
            "goal": "AIに「定期的に更新される信頼できる知識源」と認識させる",
            "means": """更新の鉄則:
・古い情報の単純削除はせず、追記更新で履歴を残す
・dateModifiedで最終更新日を明示(JSON-LD)
・同じテーマで時系列の記事を蓄積しトピッククラスタ化
一貫性の維持:
・文体・敬体・専門用語の統一(2-5参照)
・著者署名の統一フォーマット
・関連ページへの内部リンク継続
更新頻度の目安:
・FAQ: 四半期に1回見直し
・HowTo: 半年に1回見直し
・Evidence(数値データ): 年1回最新版に更新""",
        },
        {
            "title": "5-5: Evidence層・Reference層の深化実装",
            "goal": "知識資産としての厚みを継続的に増やす",
            "means": """Evidence深化:
・自社調査の継続実施(年1〜2回)
・施工事例・顧客事例の蓄積
・数値データの定期更新
・自社実験・検証結果の公開
Reference深化:
・第三者機関データへの引用拡大(政府統計、業界団体、学術論文)
・citeタグで明確に出典を区別
・外部リンクの定期的な存続確認(リンク切れ防止)
両層の効果:
・Evidence: 独自性・差別化(GEO競争力)
・Reference: 客観性・社会的信用(LLMO学習対象化)
両者を組み合わせ、独立段落で記述することで、AIは「客観性のある一次情報源」と認識し、長期的に学習対象として扱う""",
        },
    ],
}

ALL_CHAPTERS = [CH1_CONTENT, CH2_CONTENT, CH3_CONTENT, CH4_CONTENT, CH5_CONTENT]


# ============== 50-page detailed version ==============

def build_50p_doc():
    doc = Document()
    setup_doc(doc)

    cover_page(doc, "AIO最適化の教科書", "構造化サマリー", "詳細版/50ページ目安")

    # ----- Section: Executive Summary -----
    add_heading(doc, "エグゼクティブサマリー", level=1)
    add_paragraph(doc, EXEC_SUMMARY, size=11)

    add_heading(doc, "本書の中核フレームワーク", level=2)
    add_paragraph(doc, KEY_FRAMEWORKS, size=10)

    add_callout(doc, "本サマリーの構造", "全章を「課題(Problem) → 打ち手(Solution) → 各節の目的(Goal) → 手段(Means)」の順で整理。日々の業務に落とし込める粒度で記述。")

    doc.add_page_break()

    # ----- Section: 全体構造 -----
    add_heading(doc, "1. 全体構造マップ", level=1)
    add_paragraph(doc, "本書は、表面的なテクニック集ではなく、「なぜAIに選ばれない状態が起きているか」「どうすれば解決できるか」を体系化した設計書である。各章は独立したテーマを扱いながら、AIOピラミッドの階層構造に対応している。", size=10)

    add_heading(doc, "1.1 章構成と扱うテーマ", level=2)
    add_n_col_table(doc,
                    ["章", "タイトル", "扱うテーマ"],
                    CHAPTER_STRUCTURE,
                    widths=[2.5, 5.0, 9.0])

    add_heading(doc, "1.2 AIOピラミッド(本書の中核モデル)", level=2)
    add_paragraph(doc, "家屋建築のメタファー: 土台→1F→2F→3F→屋根+メンテナンス。下層が無いと上層は機能しない。", size=10)
    pyramid_rows = [
        ("最上層: 運用層", "PDCA, アクセス解析, AI引用率測定, 定期更新", "全体"),
        ("3層目: 学習対象化層", "独立した意味単位の段落, HTMLタグ構造化, 更新日・著者情報・出典URL", "LLMO"),
        ("2層目: 引用層", "QCEP構成, 比較構文, 統計データの引用, 権威ある出典の明示", "GEO"),
        ("1層目: 理解層", "PREP構成, FAQ構文/Q&A構造, 見出し構造", "AEO"),
        ("土台: 基礎構造層", "タイトル最適化, 内部リンク, 構造化データ, E-E-A-T", "SEO"),
    ]
    add_n_col_table(doc, ["階層", "具体的施策", "主担当技術"], pyramid_rows, widths=[4.5, 8.5, 3.5])
    add_paragraph(doc, "各層は「主担当技術」+「サポート技術」で構成され、最上層から1層目に戻る循環構造になっている。一度作って終わりではなく、継続運用で「AIに選ばれ続ける」状態を実現する。", size=10)

    add_heading(doc, "1.3 三層構造(AEO/GEO/LLMO)の時間軸", level=2)
    add_n_col_table(doc,
                    ["技術", "目的", "時間軸"],
                    [
                        ("AEO", "AIにサイト内容を理解させ、回答パーツとして抽出しやすくする", "短期(1-3ヶ月)"),
                        ("GEO", "AIに信頼できる情報源として引用させる", "短中期(3-12ヶ月)"),
                        ("LLMO", "将来のAIモデル更新時に学習対象として採用される", "長期(1年〜)"),
                    ],
                    widths=[2.5, 9.5, 4.0])

    add_heading(doc, "1.4 Query→Intent→Answerモデル", level=2)
    add_paragraph(doc, "AIが情報を理解する内部プロセス。コンテンツ側もこの流れに合わせる:", size=10)
    add_bullet(doc, "Query(質問): 見出しは質問形式で意図を明示")
    add_bullet(doc, "Intent(意図): 「なぜその質問をしているか」を読み取り回答を生成")
    add_bullet(doc, "Answer(回答): 三層構造で提供")
    add_bullet(doc, "  レイヤー1(PREP): 原理・本質層", indent=1)
    add_bullet(doc, "  レイヤー2(FAQ): 例外・補足層", indent=1)
    add_bullet(doc, "  レイヤー3(HowTo): 実践・手順層", indent=1)

    add_heading(doc, "1.5 E-T-R再定義(AI時代の信頼軸)", level=2)
    add_n_col_table(doc,
                    ["軸", "意味", "具体的実装"],
                    [
                        ("E (Evidence/根拠)", "数値データ・調査結果・実績を具体的に提示", "「顧客満足度96% (2025年10月当社調べ)」"),
                        ("T (Traceability/追跡性)", "出典URL・更新日・著者情報で追跡可能にする", "footer署名、citeタグ、最終更新日"),
                        ("R (Retention/定着性)", "継続的な引用と長期的な信頼の蓄積", "更新頻度、関連ページ拡充、Evidence蓄積"),
                    ],
                    widths=[3.5, 6.0, 6.5])
    add_paragraph(doc, "従来のE-E-A-Tと対立せず補完: E-E-A-Tは「信頼を伝える表現方法」、E-T-Rは「AIが信頼を判断する基準」。", size=10)

    doc.add_page_break()

    # ----- Sections for each chapter -----
    for ci, ch in enumerate(ALL_CHAPTERS):
        add_heading(doc, f"{ci+2}. {ch['title']}", level=1)

        add_heading(doc, "課題(Problem)", level=3)
        add_callout(doc, "この章が解決する大課題", ch["problem"], color="FFE0E0")

        add_heading(doc, "打ち手(Solution)", level=3)
        add_callout(doc, "本章の中核ソリューション", ch["solution"], color="E0F4E0")

        add_heading(doc, "各節の目的と手段", level=2)
        for sec in ch["sections"]:
            add_heading(doc, sec["title"], level=3)
            add_two_col_table(doc, "目的(Goal)", "手段(Means)",
                              [(sec["goal"], sec["means"])], widths=(4.0, 13.0))

        doc.add_page_break()

    # ----- Section: Implementation Roadmap -----
    add_heading(doc, "7. 実装ロードマップ(本書を実務に落とすための90日プラン)", level=1)
    add_n_col_table(doc,
                    ["フェーズ", "期間", "アクション", "達成基準"],
                    [
                        ("Phase 1: 診断", "Week 1-2", "現状監査(HTML構造、JSON-LD有無、E-E-A-T、AI引用ベースライン)", "監査レポート完成、改善対象リスト化"),
                        ("Phase 2: 基盤整備", "Week 3-6", "重要ページのHTML階層整備、FAQ構文実装、PREP法でリライト、E-E-A-Tテンプレ整備", "Top10ページがChapter2-3要件を満たす"),
                        ("Phase 3: 構造化", "Week 7-9", "FAQPageスキーマ・HowToスキーマ実装、Evidence+Reference追加、llms.txt配置", "JSON-LDが正しく動作、リッチスニペット表示"),
                        ("Phase 4: 計測開始", "Week 10-12", "Citation Rate計測開始、FAQ管理表作成、月次PDCA確立", "引用率ベースライン確定、運用ルーティン化"),
                        ("Phase 5: 継続改善", "Week 13〜", "PDCA運用、AI別最適化、Evidence層・Reference層の深化", "引用率60%以上を維持"),
                    ],
                    widths=[3.0, 2.5, 7.5, 4.0])

    # ----- Section: Glossary -----
    add_heading(doc, "8. 用語集", level=1)
    add_two_col_table(doc, "用語", "定義",
                      [
                          ("AIO", "AI Optimization。AEO+GEO+LLMO+SEOを統合した設計思想"),
                          ("AEO", "Answer Engine Optimization。AIに理解させ回答に使われやすくする最適化"),
                          ("GEO", "Generative Engine Optimization。AIに引用される最適化"),
                          ("LLMO", "Large Language Model Optimization。将来のAI学習対象になる最適化"),
                          ("PREP法", "Point/Reason/Example/Pointの4段階で文章を構成する方法"),
                          ("FAQ構文", "Question/Answer形式で疑問と回答を整理する構文"),
                          ("HowTo構文", "段階的な手順を説明する構文"),
                          ("QCEP構文", "Question/Conclusion/Evidence/Publisher。GEO向け文章構造(本書独自)"),
                          ("E-T-R", "Evidence/Traceability/Retention。AI時代の信頼3軸(本書独自)"),
                          ("引用率(Citation Rate)", "(引用回数 ÷ チェック総数) × 100。本書独自の運用指標"),
                          ("JSON-LD", "JSON for Linked Data。構造化データのフォーマット"),
                          ("FAQPageスキーマ", "schema.orgのFAQページ用構造化データ仕様"),
                          ("HowToスキーマ", "schema.orgのHowTo手順用構造化データ仕様"),
                          ("Evidence層", "自社の根拠(数値・実績・調査)を示すセクション"),
                          ("Reference層", "外部の権威(第三者機関、公的データ)への参照セクション"),
                          ("llms.txt", "AI向けサイト案内ファイル。優先参照ページを明示"),
                          ("AIOピラミッド", "SEO土台+AEO+GEO+LLMO+運用の5階層モデル(本書中核)"),
                      ],
                      widths=(4.0, 13.0))

    add_paragraph(doc, "出典: 瀧内賢『これからはじめるAIO AI最適化の教科書』技術評論社, 2026年3月発行, 全464ページ", size=9, color="888888")

    return doc


# ============== 10-page brief version ==============

def build_10p_doc():
    doc = Document()
    setup_doc(doc)

    cover_page(doc, "AIO最適化の教科書", "エグゼクティブサマリー", "要旨版/10ページ目安")

    # Page 1: Exec Summary
    add_heading(doc, "1. エグゼクティブサマリー", level=1)
    add_paragraph(doc, EXEC_SUMMARY, size=11)
    add_callout(doc, "本書の結論", "AIに「選ばれ続ける」ためには、AEO+GEO+LLMOを統合したAIO設計と、PDCAによる継続運用が必須。一度作って終わりのSEO的発想から、長期的な知識資産育成へのパラダイム転換が求められる。")

    doc.add_page_break()

    # Page 2: Big Picture
    add_heading(doc, "2. AIO全体像 - AIOピラミッド", level=1)
    add_paragraph(doc, "本書の中核モデル。家屋建築のメタファーで、土台→4階層→屋根(運用)を順に積み上げる。", size=10)
    pyramid_rows = [
        ("最上層: 運用", "PDCA, AI引用率測定, 定期更新", "全体"),
        ("3層目: 学習対象化", "段落構造化, 更新日, 著者情報, 出典URL", "LLMO"),
        ("2層目: 引用層", "QCEP構成, 統計データ, 権威ある出典", "GEO"),
        ("1層目: 理解層", "PREP構成, FAQ/Q&A構造, 見出し設計", "AEO"),
        ("土台: 基礎構造", "タイトル, 内部リンク, 構造化データ, E-E-A-T", "SEO"),
    ]
    add_n_col_table(doc, ["階層", "具体的施策", "主担当"], pyramid_rows, widths=[4.0, 9.0, 3.5])

    add_heading(doc, "三層構造(AEO/GEO/LLMO)の時間軸", level=2)
    add_n_col_table(doc, ["技術", "目的", "時間軸"],
                    [
                        ("AEO", "AIに理解させ、回答に使われやすく", "短期 1-3ヶ月"),
                        ("GEO", "AIに信頼できる情報源として引用させる", "短中期 3-12ヶ月"),
                        ("LLMO", "将来のAI学習対象になる", "長期 1年〜"),
                    ], widths=[2.5, 9.5, 4.0])

    doc.add_page_break()

    # Page 3: Chapter 1
    add_heading(doc, "3. Chapter 1: AIOの基本", level=1)
    add_callout(doc, "課題", "Googleで1位を目指してきたWebサイトが、AI時代には「存在しない」のと同じになる。AEO/GEO/LLMOの統合的設計思想がない。", color="FFE0E0")
    add_callout(doc, "打ち手", "AIO=AEO+GEO+LLMO+SEOの統合設計思想。AIOピラミッドという階層モデルで全体像を提示。", color="E0F4E0")
    add_heading(doc, "重要フレームワーク", level=2)
    add_bullet(doc, "Query→Intent→Answer: AIの理解プロセスに合わせ、見出し(Query形式)→PREP原理層→FAQ深化層→HowTo具体層")
    add_bullet(doc, "E-T-R: Evidence(根拠)×Traceability(追跡性)×Retention(定着性)。AIが信頼を判断する3軸")
    add_bullet(doc, "AEO/GEO/LLMO三層: 短期(AEO/GEO 3-12ヶ月)→長期(LLMO 1年〜)の連動施策")

    doc.add_page_break()

    # Page 4-5: Chapter 2
    add_heading(doc, "4. Chapter 2: 文章構成", level=1)
    add_callout(doc, "課題", "AIに『読まれる』だけでは引用されない。AIは構造を読み取って意味を理解する。", color="FFE0E0")
    add_callout(doc, "打ち手", "PREP法→FAQ構文→HowTo構文の3つの『型』を1ページ内で組み合わせる。+ meta+llms.txt+著者情報で入口設計、E-E-A-Tで信頼可視化。", color="E0F4E0")
    add_heading(doc, "PREP法", level=2)
    add_paragraph(doc, "Point(結論)→Reason(理由)→Example(具体例)→Point(結論再強調)。AIに「主張→根拠→裏付け→結論」を明示し、部分引用しても文脈が崩れない最小単位を作る。", size=10)
    add_heading(doc, "FAQ構文(Q&A構造)", level=2)
    add_paragraph(doc, "dl/dt/ddタグで質問と回答の対応を構造的に明示。AIが「この部分は○○という質問への回答」と認識しやすくなり、引用候補として優先される。", size=10)
    add_heading(doc, "HowTo構文", level=2)
    add_paragraph(doc, "「やり方・手順」が必要な場合のみ。Step1, Step2...と番号付きで明確化。AIが学習対象として優先採用しやすくなる。", size=10)
    add_heading(doc, "入口設計と信頼設計", level=2)
    add_bullet(doc, "meta(title 25字+description 100字)+OGP: ページ単位の名札")
    add_bullet(doc, "llms.txt: サイトルートに配置、AIに優先参照ページを案内")
    add_bullet(doc, "著者情報(schema.org Person): 経歴・専門分野・組織を構造化")
    add_bullet(doc, "個人署名(footer)+構造署名(FAQ内cite)で信頼を可視化")

    doc.add_page_break()

    # Page 6-7: Chapter 3
    add_heading(doc, "5. Chapter 3: 文書構造の設計", level=1)
    add_callout(doc, "課題", "Chapter 2の文章をAIが正しく解釈できないと引用に至らない。HTMLレベルの構造設計が無いと、同じ内容でも拾われ方が変動する。", color="FFE0E0")
    add_callout(doc, "打ち手", "HTML階層+JSON-LDスキーマ+Evidence/Reference+SEO/AIO連携で、AIの解釈を「推論依存」から「定義情報」へ転換。", color="E0F4E0")
    add_heading(doc, "実装の4階層", level=2)
    add_n_col_table(doc, ["階層", "施策", "効果"],
                    [
                        ("HTML構造", "h1→h2→h3階層、dl/dt/dd、section", "AIがページの論理構造を把握"),
                        ("構造化データ", "FAQPageスキーマ、HowToスキーマ(JSON-LD)", "AIに「定義情報」として明示"),
                        ("信頼性層", "Evidence(自社根拠)+Reference(外部権威)", "AIが「客観性ある一次情報源」と判断"),
                        ("SEO/AIO連携", "canonical、内部リンク、title/h1/meta整合", "検索とAI両方に発見・評価される"),
                    ], widths=[3.5, 6.5, 7.0])

    doc.add_page_break()

    # Page 8-9: Chapter 4 (most important per user's emphasis)
    add_heading(doc, "6. Chapter 4: PDCA運用 ★最重要章", level=1)
    add_callout(doc, "課題", "AIに引用される文書を作っても、AIアルゴリズム/競合変動で評価は変動する。継続的測定・改善が無いと「作って放置」で引用率低下。", color="FFE0E0")
    add_callout(doc, "打ち手", "Citation Rate(引用率)という独自指標を定義し、PDCAサイクルをGoogle Sheetsで運用管理。月10-15分の作業で継続最適化。", color="E0F4E0")
    add_heading(doc, "PDCAの4段階", level=2)
    add_n_col_table(doc, ["段階", "アクション"],
                    [
                        ("Plan(計画)", "FAQ管理表で「今あるFAQ」を見える化、改善優先順位を整理"),
                        ("Do(実行)", "新規FAQの構造化、または前回Actで計画した改善を実装"),
                        ("Check(確認)", "Perplexity/Google AI Overviews/ChatGPTでAI引用状況を確認、引用率を測定"),
                        ("Act(改善)", "Checkで発見した問題を分析、次回Doへの改善計画を立案"),
                    ], widths=[3.0, 14.0])

    add_heading(doc, "引用率(Citation Rate)の計算と目標", level=2)
    add_paragraph(doc, "計算式: (引用回数 ÷ チェック総数) × 100", size=10, bold=True)
    add_bullet(doc, "目標値: 維持60% / 理想70% / 改善対象50%未満")
    add_bullet(doc, "「引用」の定義: AI回答画面の『出典(Sources)』に自社URLが表示されている状態。回答内容が似ていても出典表示が無ければカウントしない")

    add_heading(doc, "確認頻度と運用工数", level=2)
    add_n_col_table(doc, ["時期", "頻度", "工数"],
                    [
                        ("初期(公開後1-2ヶ月)", "週1回×各FAQ3回ずつ検索", "約15分/週"),
                        ("定期運用(3ヶ月以降)", "月1回×各FAQ5回ずつ検索", "約12.5分(=25回×30秒)+記録 = 計10-15分/月"),
                    ], widths=[5.0, 7.0, 5.0])

    add_heading(doc, "管理する2種類のスプレッドシート", level=2)
    add_bullet(doc, "FAQ記録表(生データ): 確認のたびに1行追加。確認日/AI/検索クエリ/FAQNo./引用?/メモ")
    add_bullet(doc, "FAQ集計表(月末作成): 1FAQに1行サマリー。引用回数/チェック総数/引用率%/判定")

    doc.add_page_break()

    # Page 10: Chapter 5 + Roadmap
    add_heading(doc, "7. Chapter 5: 知識資産化(LLMO)", level=1)
    add_callout(doc, "課題", "短期の引用獲得だけでは、将来のAIモデル更新時に淘汰される可能性。", color="FFE0E0")
    add_callout(doc, "打ち手", "Web全体検索型AIの内部処理を理解→llms.txt配置→AI別最適化→Evidence/Reference深化→更新頻度・一貫性維持。", color="E0F4E0")
    add_heading(doc, "重要施策", level=2)
    add_bullet(doc, "発見条件3つ: Googleインデックス登録 / llms.txt設置 / robots.txtでAIクローラー許可")
    add_bullet(doc, "AI別最適化: Perplexity(Evidence重視)、Gemini(SEO+JSON-LD強化)、ChatGPT(過去学習データへの定着)")
    add_bullet(doc, "更新の鉄則: 単純削除せず追記更新、dateModified明示、トピッククラスタ化")

    add_heading(doc, "8. 90日実装ロードマップ", level=1)
    add_n_col_table(doc, ["フェーズ", "期間", "アクション"],
                    [
                        ("診断", "Week 1-2", "現状監査、改善対象リスト化"),
                        ("基盤整備", "Week 3-6", "HTML階層+FAQ構文+E-E-A-Tテンプレ"),
                        ("構造化", "Week 7-9", "JSON-LD実装、Evidence+Reference、llms.txt"),
                        ("計測開始", "Week 10-12", "Citation Rate計測、月次PDCA確立"),
                        ("継続改善", "Week 13〜", "PDCA運用、AI別最適化、引用率60%維持"),
                    ], widths=[3.0, 2.5, 11.5])

    add_paragraph(doc, "", size=8)
    add_paragraph(doc, "出典: 瀧内賢『これからはじめるAIO AI最適化の教科書』技術評論社, 2026年3月発行, 全464ページ", size=9, color="888888")

    return doc


# ============== Build & save ==============

doc_50 = build_50p_doc()
doc_50.save("/home/user/AIO/AIO_Textbook_Summary_50p.docx")
print("Saved: AIO_Textbook_Summary_50p.docx")

doc_10 = build_10p_doc()
doc_10.save("/home/user/AIO/AIO_Textbook_Summary_10p.docx")
print("Saved: AIO_Textbook_Summary_10p.docx")
