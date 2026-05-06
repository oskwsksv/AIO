"""Generate 4-page summary Word doc of AIOutmation Lean Canvas (no cover)."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.section import WD_ORIENT
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


def write_cell(cell, text, size=9, bold=False, bg=None, color=None, align=None, italic=False):
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
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(0)
        if align:
            p.alignment = align
        run = p.add_run(line)
        style_run(run, size=size, bold=bold, color=color, italic=italic)


def write_block(cell, header, body, header_bg, body_bg, header_color="FFFFFF", header_size=10, body_size=8):
    """Lean Canvas block with colored header strip and body."""
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    set_cell_bg(cell, body_bg)

    # Header
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(2)
    p1.paragraph_format.space_before = Pt(0)
    r = p1.add_run(header)
    style_run(r, size=header_size, bold=True, color=header_color)
    # Set header bg via separate visual approach is complex; use small font color contrast on body bg

    # Body
    lines = body.split("\n")
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        r = p.add_run(line)
        style_run(r, size=body_size, color="222222")


def add_heading(doc, text, level=1, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8 if level <= 1 else 6)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    sizes = {0: 22, 1: 16, 2: 13, 3: 11}
    default_colors = {0: "1F3A5F", 1: "1F3A5F", 2: "2E5C8A", 3: "444444"}
    style_run(run, size=sizes.get(level, 11), bold=True, color=color or default_colors.get(level, "000000"))


def add_paragraph(doc, text, size=10, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    style_run(run, size=size, bold=bold, color=color)


def setup_doc(doc):
    section = doc.sections[0]
    new_w, new_h = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_w
    section.page_height = new_h
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    style = doc.styles["Normal"]
    style.font.name = "Yu Gothic"
    style.font.size = Pt(10)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "Yu Gothic")


# ============== Build Document ==============
doc = Document()
setup_doc(doc)


# ============== PAGE 1: Lean Canvas 全景 (visual grid) ==============
add_heading(doc, "AIOutmation リーンキャンバス(全景)", level=0)
add_paragraph(doc, "ターゲット顧客: 上場会社のWebマーケティング部門 + ネット広告代理店  /  作成: 2026年5月", size=9, color="555555")

# UVP banner across the top
banner = doc.add_table(rows=1, cols=1)
write_cell(banner.cell(0, 0),
    "③UVP: 瀧内本「AIO教科書」をソフトウェアで完全実装。測定→改修→効果検証→経営報告までワンプラットフォームで自動化する、日本市場初の統合型AIO基盤。\n年間1,500時間の手作業を100時間に圧縮し、AI時代のマーケROIを再定義する。",
    size=10, bold=True, color="1F3A5F", bg="FFF8E0", align=WD_ALIGN_PARAGRAPH.CENTER)

# 3x3 Lean Canvas Grid
t = doc.add_table(rows=3, cols=3)
t.autofit = False
col_w = Cm(8.6)
row_h = [Cm(5.5), Cm(5.5), Cm(5.0)]
for ri, r in enumerate(t.rows):
    for c in r.cells:
        c.width = col_w

# Row 1
write_block(t.cell(0, 0), "①課題 (Problem) Top 5",
"""・AI検索可視性の測定が手作業地獄(月12分超×多モデル×数百本)
・計測→改修(コンテンツ・Schema・CMS)が完全分断、4部門横断で1施策2-6週間
・AIO業界SOPが日本市場に存在せず属人化・知識流出
・経営/クライアントへのROI証明が困難
・既存ベンチマークツールは計測特化で改修・PR・経営報告まで届かない""",
    header_bg="C04000", body_bg="FFE8E0")

write_block(t.cell(0, 1), "④ソリューション (Solution) - 主要機能8本",
"""①SEO×AIO監査エンジン  ②文章構成最適化AI(PREP/FAQ/HowTo)
③プロンプト自動生成エンジン(CRM連携)
④文書構造化エンジン(JSON-LD自動+CMS PR起票)
⑤引用率測定エンジン(PDCA完全自動化+Signal/Noise)
⑥知識資産化アシスタント(llms.txt自動)
⑦Executive Reporter(CMO 3行サマリ+ROI物語)
⑧マルチテナント運用(代理店向けホワイトラベル)""",
    header_bg="2E5C8A", body_bg="E8EEF7")

write_block(t.cell(0, 2), "②顧客セグメント (Customer Segments)",
"""【主要①】上場企業Webマーケ部門
B2B SaaS / 金融 / 大手リテール / 製薬 / 通信
決裁: マーケ部長→CMO→CFO/取締役会
個人ペルソナ: Sr.Manager(GEO/AEO Lead)35-40代

【主要②】ネット広告代理店
50-300名規模、上場企業5-12社/担当
月額40-450万円リテイナー、Account Director""",
    header_bg="C04000", body_bg="FFF0F0")

# Row 2
write_block(t.cell(1, 0), "⑧主要指標 (Key Metrics)",
"""・ARR(24ヶ月で5億円目標)
・上場企業ロゴ獲得数 / 代理店パートナー数
・NRR 120%以上、CAC回収12ヶ月以内
・顧客平均Citation Rate改善量(導入前→3ヶ月後)
・自動生成CMS PR数 / 採用率""",
    header_bg="2E5C8A", body_bg="E8EEF7")

write_block(t.cell(1, 1), "⑨圧倒的優位 (Unfair Advantage)",
"""【1】瀧内本「AIO教科書」のソフトウェア実装で先行
【2】日本語LLM完全対応(Felo/Genspark/PKSHA)
【3】Signal/Noise Engine(統計的処理を商用実装)
【4】計測→改修PR→効果検証クローズドループ
【5】業界ベンチマーク独自蓄積データ
【6】上場企業×代理店の同時最適化""",
    header_bg="2E7A4E", body_bg="E0F4E0")

write_block(t.cell(1, 2), "⑤チャネル (Channels)",
"""1. ダイレクトセールス(エンプラ向け、PoC/CSM)
2. ネット広告代理店との戦略提携(売上シェア)
3. AIO教科書著者・SEO業界KOLとの連携
4. コンテンツマーケ(ベンチマーク/Webinar)
5. PLG(セルフサーブ)→アップセル
6. プロダクト統合(HubSpot/Salesforce AppExchange)""",
    header_bg="C9622A", body_bg="FFF4E0")

# Row 3
write_block(t.cell(2, 0), "⑦コスト構造 (Cost Structure)",
"""【固定】開発5-10名、セールス/CSM 5-10名、管理2名
【変動】LLM API(顧客あたり月1-5万円)
   UIスクレイピング基盤(月100-500万円)
   クラウドインフラ(月50-300万円)
【一時】SOC 2 取得 1500-3000万円、PR/マーケ 1000-3000万円
損益分岐点: ARR 3億円(エンプラ30社相当)""",
    header_bg="A02020", body_bg="FFE8E8")

write_block(t.cell(2, 1), "③UVP (再掲)",
"""書く前から書いた後まで、AIOを自動で回す。
─ Webマーケ向け『AIO版GitHub Copilot+Datadog』""",
    header_bg="C04000", body_bg="FFF8E0")

write_block(t.cell(2, 2), "⑥収益の流れ (Revenue Streams)",
"""【SaaS主軸】Starter月3万/Growth月10万/Business月30万
   Enterprise 月50-200万/代理店 月20万〜(クレジット制)
【補助】オンボード/コンサル 50-300万、CMSコネクタ開発 100-500万
   業界ベンチマークレポート 単発10-50万
24ヶ月後: ARR 5億円目標(エンプラ50+代理店30+SMB200)""",
    header_bg="2E7A4E", body_bg="E0F4E0")

doc.add_page_break()


# ============== PAGE 2: Customer Segments + Problem 詳細 ==============
add_heading(doc, "②顧客セグメント (Customer Segments) + ①課題 (Problem) 詳細", level=0)

add_heading(doc, "②顧客セグメント", level=2)
t = doc.add_table(rows=3, cols=2)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(6.0)
    r.cells[1].width = Cm(20.5)

write_cell(t.cell(0, 0), "【主要顧客①】\n上場会社のWebマーケティング部門", size=11, bold=True, bg="FFE0E0")
write_cell(t.cell(0, 1),
"""・想定: 東証プライム/グロース市場上場の事業会社、ARR/売上100億円〜数千億円規模
・業種優先順位: B2B SaaS > 金融・保険 > 大手リテール・EC > 製薬・ヘルスケア > 通信・メディア
・購買決裁ライン: マーケ部長/Director → CMO/VP Marketing → CFO/取締役会
・関与チーム: SEO/コンテンツ/PR/Web開発/法務/情シス/購買
・個人ペルソナ: Sr. Manager(GEO/AEO Lead)、30代後半、KGIはパイプライン貢献+ブランドVisibility
・既存ツール: GA4 / Search Console / Semrush / HubSpot / Salesforce / Looker""",
                  size=10, bg="FFF8F8")

write_cell(t.cell(1, 0), "【主要顧客②】\nネット広告代理店\n(上場企業をクライアントとする)", size=11, bold=True, bg="FFE0E0")
write_cell(t.cell(1, 1),
"""・想定: 50-300名規模、SEO/コンテンツ/運用広告を展開、月額リテイナー型
・顧客層: 上場B2B SaaS 5-12社/担当、月額40〜450万円
・購買決裁ライン: Group Director / Partner
・個人ペルソナ: Account Director/GEO Lead、KGIはリテンション+新規受注+リテイナー単価
・既存ツール: Semrush / Ahrefs / GA4 / Looker Studio / AgencyAnalytics / Asana""",
                  size=10, bg="FFF8F8")

write_cell(t.cell(2, 0), "【副次セグメント(将来拡張)】", size=11, bold=True, bg="F5F5F5")
write_cell(t.cell(2, 1),
"国内Tier 2のコンサルティングファーム、PR代理店、SI企業のマーケDX部門",
                  size=10)

add_heading(doc, "①課題 (Top 7)", level=2)
t = doc.add_table(rows=7, cols=2)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(0.8)
    r.cells[1].width = Cm(25.7)

problems = [
    "AI検索可視性の測定が手作業地獄(月12.5分×多モデル×多プロンプト=実質数時間)",
    "計測→改修(コンテンツ・Schema・CMS入稿)が完全分断、4部門横断で1施策2-6週間",
    "AIO業界SOPが日本市場に存在せず、各社が独自にゼロから設計→属人化・知識流出",
    "経営/クライアントへのROI証明が困難(Visibility→売上の因果が示せない、業界ベンチマーク無し)",
    "ノイズ(LLM揺らぎ)とシグナルの分離が統計的にできず、誤検知or見逃しが頻発",
    "多言語・多リージョン・サブブランドへのスケール対応で人手が爆発(代理店は3社目で破綻)",
    "既存ベンチマークツール(Profound/Peec/Passionfruit)は計測特化で改修・PR・経営報告まで届かない",
]
for i, p in enumerate(problems):
    write_cell(t.cell(i, 0), str(i+1), size=11, bold=True, bg="FFE0E0", color="C04000", align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(t.cell(i, 1), p, size=10)

doc.add_page_break()


# ============== PAGE 3: Solution 詳細 + Unfair Advantage ==============
add_heading(doc, "④ソリューション (Solution) + ⑨圧倒的優位 (Unfair Advantage)", level=0)

add_heading(doc, "④ソリューション - AIOピラミッド全層をカバーする主要機能8本", level=2)
t = doc.add_table(rows=9, cols=3)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(0.8)
    r.cells[1].width = Cm(7.0)
    r.cells[2].width = Cm(18.7)

write_cell(t.cell(0, 0), "#", size=10, bold=True, bg="2E5C8A", color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
write_cell(t.cell(0, 1), "機能名", size=10, bold=True, bg="2E5C8A", color="FFFFFF")
write_cell(t.cell(0, 2), "内容", size=10, bold=True, bg="2E5C8A", color="FFFFFF")

solutions = [
    ("①", "SEO×AIO監査エンジン",
     "既存サイトのHTML階層・構造化データ・E-E-A-T充足度を自動診断。AIOピラミッド層別ギャップ可視化と90日ロードマップ自動生成。"),
    ("②", "文章構成最適化AI(教科書Ch.2準拠)",
     "PREP法→FAQ構文→HowTo構文への自動リライト。既存ページを取り込み、AEO対応版を提案。"),
    ("③", "プロンプト自動生成エンジン(教科書Ch.1+4準拠)",
     "自社URL+CRM(Salesforce/HubSpot)+営業会話(Gong/Zoom)から、ICP×購買ファネル×製品でMoney Promptを30-1,000本自動生成。「なぜこのプロンプトか」をAIが自動文書化(属人化解消)。"),
    ("④", "文書構造化エンジン(教科書Ch.3準拠)",
     "JSON-LD自動生成(FAQPage/HowTo/Article/Person/Organization)。Evidence+Reference自動挿入。GitHub/Webflow/WordPress API連携でCMSにPR自動起票。"),
    ("⑤", "引用率測定エンジン(教科書Ch.4 PDCA完全自動化) ★中核",
     "LLM API+UIスクレイピングで毎日プロンプト実行。Citation Rate自動算出、FAQ管理表/記録表/集計表を自動更新。70%閾値割れで自動アラート、原因分析AIが改修候補提示。Signal/Noise Engine: LLM揺らぎを統計的に除去。"),
    ("⑥", "知識資産化アシスタント(教科書Ch.5準拠)",
     "llms.txt自動生成・更新。Perplexity/Gemini/ChatGPT別の最適化レコメンド。更新頻度・一貫性スコアを継続監視。"),
    ("⑦", "Executive Reporter",
     "CMO向け3行サマリ自動生成(前月比/カテゴリ/競合差分)。QBR用デック自動作成、ROI物語AI(過去成功事例DBから類似ストーリー提示)。業界ベンチマーク同梱。"),
    ("⑧", "代理店向けマルチテナント運用",
     "クライアント別ホワイトラベル(ロゴ/カラー/独自ドメイン/送信元メール)。クレジット制(プロンプト枠を流動配分)、全プラン無制限シート。Pitch環境(無料即発行、契約条件で自動失効)。"),
]
for i, (num, name, desc) in enumerate(solutions):
    is_core = "中核" in name
    bg = "FFF4E0" if is_core else None
    write_cell(t.cell(i+1, 0), num, size=11, bold=True, bg="E8EEF7", align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(t.cell(i+1, 1), name, size=10, bold=True, bg=bg)
    write_cell(t.cell(i+1, 2), desc, size=9, bg=bg)

add_heading(doc, "⑨圧倒的優位 - 6つの参入障壁", level=2)
t = doc.add_table(rows=6, cols=2)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(5.0)
    r.cells[1].width = Cm(21.5)

advantages = [
    ("【1】AIO教科書のソフトウェア実装で先行",
     "瀧内賢氏「AIO教科書」(2026年3月技術評論社)をプロダクト化する世界初のソフトウェア。著者監修・推奨・共同マーケが成立すれば差別化決定打。"),
    ("【2】日本語LLM完全対応",
     "Felo/Genspark/PKSHA/国産LLMの完全カバー。Profound/Peec/Passionfruitの最弱点を突く。Perplexity日本語/ChatGPT日本語/Gemini日本語の最適化分離。"),
    ("【3】Signal/Noise Engine",
     "LLM評価の研究フロンティア(Ai2 framework等)を商用実装。競合3社は閾値ベース判定、当社は信頼区間付き判定で誤検知激減。"),
    ("【4】計測→改修PR→効果検証のクローズドループ",
     "競合3社は計測特化、当社のみCMS PR自動起票・改修ロジック蓄積まで実装。「測って終わり」を脱する唯一のプラットフォーム。"),
    ("【5】業界ベンチマーク独自蓄積データ",
     "顧客横断データ(匿名化)から業種別中央値・上位10%を提示。経営説得力の決定的な武器。"),
    ("【6】上場企業×代理店の同時最適化",
     "一般的なSaaSは「直販 vs 代理店」を二者択一にするが、当社は同一プラットフォームで両セグメント最適化(クレジット制+ホワイトラベル+Pitch)。"),
]
for i, (head, body) in enumerate(advantages):
    write_cell(t.cell(i, 0), head, size=10, bold=True, bg="E0F4E0", color="2E7A4E")
    write_cell(t.cell(i, 1), body, size=10)

doc.add_page_break()


# ============== PAGE 4: Channels / Revenue / Cost / Key Metrics ==============
add_heading(doc, "⑤チャネル + ⑥収益 + ⑦コスト + ⑧主要指標", level=0)

# Channels
add_heading(doc, "⑤チャネル (Channels)", level=2)
t = doc.add_table(rows=2, cols=2)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(13.0)
    r.cells[1].width = Cm(13.5)
write_cell(t.cell(0, 0), "【主要チャネル】", size=10, bold=True, bg="FFF4E0", color="C9622A")
write_cell(t.cell(0, 1), "【サブチャネル】", size=10, bold=True, bg="FFF4E0", color="C9622A")
write_cell(t.cell(1, 0),
"""1. ダイレクトセールス(エンタープライズ向け): 上場企業マーケ部門への営業/PoC/CSM
2. パートナーシップ販売: 大手ネット広告代理店との戦略提携(売上シェアモデル)
3. 著者・専門家ネットワーク: AIO教科書著者瀧内氏/SEO業界KOLとの連携
4. コンテンツマーケ: ベンチマークレポート・業界調査・Webinar公開
5. PLG(セルフサーブ): 下位プランで個人事業主・SMBを獲得→アップセル""", size=9)
write_cell(t.cell(1, 1),
"""6. 業界カンファレンス(MarkeZine Day, AdTech Tokyo, Adobe Summit Japan等)
7. プロダクト統合パートナー(HubSpot Marketplace, Salesforce AppExchange)
8. メディア掲載(日経クロストレンド, MarkeZine, Web担当者Forum)""", size=9)

# Revenue
add_heading(doc, "⑥収益の流れ (Revenue Streams)", level=2)
t = doc.add_table(rows=6, cols=3)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(5.5)
    r.cells[1].width = Cm(7.0)
    r.cells[2].width = Cm(14.0)
write_cell(t.cell(0, 0), "プラン", size=10, bold=True, bg="2E7A4E", color="FFFFFF")
write_cell(t.cell(0, 1), "価格", size=10, bold=True, bg="2E7A4E", color="FFFFFF")
write_cell(t.cell(0, 2), "対象/特徴", size=10, bold=True, bg="2E7A4E", color="FFFFFF")
plans = [
    ("Starter", "月額3万円〜", "50プロンプト、3LLMモデル、SMB向けセルフサーブ"),
    ("Growth", "月額10万円", "150プロンプト、5LLM、CMS連携1個"),
    ("Business", "月額30万円", "500プロンプト、全LLM、CMS連携無制限、CSM"),
    ("Enterprise", "月額50-200万円", "無制限プロンプト、SOC 2、SSO、専任CSM、API"),
    ("代理店プラン", "月額20万円〜", "クレジット制、ホワイトラベル、Pitch環境、無制限シート"),
]
for i, (name, price, desc) in enumerate(plans):
    write_cell(t.cell(i+1, 0), name, size=10, bold=True, bg="E0F4E0")
    write_cell(t.cell(i+1, 1), price, size=10)
    write_cell(t.cell(i+1, 2), desc, size=9)

add_paragraph(doc,
    "【補助収益】オンボード/コンサル スポット50-300万円、カスタムCMSコネクタ開発 100-500万円、業界ベンチマークレポート販売 10-50万円",
    size=9, color="555555")
add_paragraph(doc,
    "【KPI想定】24ヶ月で ARR 5億円(エンプラ50社 + 代理店30社 + SMB 200社)",
    size=10, bold=True, color="2E7A4E")

# Cost Structure
add_heading(doc, "⑦コスト構造 (Cost Structure)", level=2)
t = doc.add_table(rows=4, cols=2)
t.autofit = False
for r in t.rows:
    r.cells[0].width = Cm(4.0)
    r.cells[1].width = Cm(22.5)
write_cell(t.cell(0, 0), "【固定費】", size=10, bold=True, bg="FFE8E8", color="A02020")
write_cell(t.cell(0, 1),
    "開発人件費(エンジニア5-10名 年4,000-8,000万円) + セールス/CSM(5-10名 年4,000-8,000万円) + 経営/管理(年2,000万円) + オフィス/SaaS管理(年500-1,000万円)", size=9)
write_cell(t.cell(1, 0), "【変動費】", size=10, bold=True, bg="FFE8E8", color="A02020")
write_cell(t.cell(1, 1),
    "LLM API利用料(1顧客月1-5万円) + UIスクレイピング基盤運用(月100-500万円) + クラウドインフラ(月50-300万円)", size=9)
write_cell(t.cell(2, 0), "【一時費用】", size=10, bold=True, bg="FFE8E8", color="A02020")
write_cell(t.cell(2, 1),
    "SOC 2 Type II取得 1,500-3,000万円(初年度) + 業界ベンチマーク用パートナーシップ 500-1,000万円 + PR/マーケティング初期投資 1,000-3,000万円", size=9)
write_cell(t.cell(3, 0), "【損益分岐点】", size=10, bold=True, bg="FFE8E8", color="A02020")
write_cell(t.cell(3, 1), "ARR 3億円(エンプラ30社相当)", size=10, bold=True)

# Key Metrics
add_heading(doc, "⑧主要指標 (Key Metrics)", level=2)
t = doc.add_table(rows=1, cols=3)
t.autofit = False
for r in t.rows:
    for c in r.cells:
        c.width = Cm(8.83)
write_cell(t.cell(0, 0),
    "【グロース指標】\n・ARR(24ヶ月で5億円目標)\n・上場企業ロゴ獲得数\n・代理店パートナー数 / 管理クライアント数\n・NRR 120%以上",
    size=10, bg="E8EEF7")
write_cell(t.cell(0, 1),
    "【プロダクト指標】\n・顧客平均Citation Rate改善量(導入前→3ヶ月後)\n・月間プロンプト実行数(プラットフォーム全体)\n・自動生成CMS PR数 / 採用率\n・経営報告自動生成回数 / 顧客満足度",
    size=10, bg="E8EEF7")
write_cell(t.cell(0, 2),
    "【効率指標】\n・CAC回収期間 12ヶ月以内\n・LTV/CAC比率 3倍以上\n・セルフサーブ→有料転換率 5%以上\n・エンタープライズ商談勝率 30%以上",
    size=10, bg="E8EEF7")

doc.save("/home/user/AIO/AIOutmation_LeanCanvas_Summary4p.docx")
print("Saved: AIOutmation_LeanCanvas_Summary4p.docx")
