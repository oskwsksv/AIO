# Passionfruit Labs の GA連携：3つの観点で整理

## 1️⃣ どのように連携するか（連携方法）

**GA4 をネイティブ統合**（OAuth でGA4プロパティを接続）。加えて **Google Search Console (GSC)** とも連携します。

技術的な肝は、**AI流入を正しく識別する仕組み**です：

| 課題 | Passionfruitの解決 |
|---|---|
| GA4は標準だとAI流入を「Direct」や「Referral」に誤分類してしまう | リファラドメイン（chatgpt.com / perplexity.ai / claude.ai / gemini.google.com 等）を正規表現で識別し、**「AI検索」チャネル**として分離 |

つまり、放っておくと埋もれてしまうAI経由トラフィックを、専用チャネルとして可視化した上でGA4の売上・CVデータと突合します。

---

## 2️⃣ どのような情報を連携するか

| データソース | 連携する情報 |
|---|---|
| **Passionfruit本体** | AI上での引用・メンション状況（どのプロンプトで、どのLLMが、自社をどう引用したか） |
| **GA4** | エンゲージメント率、トラフィック、購入、**収益**（LLM別＝ChatGPT / Perplexity / Claude / Gemini / Grok / DeepSeek / Qwen 等の流入別） |
| **Search Console** | ブランド検索のインプレッション・クリック（**ブランド検索ハロー効果**の測定） |

---

## 3️⃣ どのような価値を提供するか

**「AI可視性 → 売上」の因果を、ファネル全体でつなぐ**ことが最大の価値です。

### Passionfruit が提唱する GEOファネル（5段階）

```
①ユーザーがプロンプト入力
   ↓
②クエリのfan-out（AIが質問を分解）
   ↓
③取得・評価（Retrieval）
   ↓
④統合・引用（Citation）← Passionfruitが測定
   ↓
⑤ユーザー行動（クリック / ブランド検索 / 成約）← GA4・GSCが測定
```

### 提供される具体的な価値指標

- **Revenue per Citation**（1引用あたり収益）
- **Cost per Lead**（AI経由のリード単価）
- **LLM別の収益貢献**（ChatGPTとPerplexity、どちらが売上を生むか）
- **ブランド検索ハロー**：AI可視性が上がると、GA4では直接追えなくても、GSCのブランド検索が増える相関を示すことで「AIメンションが売上を牽引している」と立証

---

## 💡 私たちのAIOutmationへの示唆

これは前回の分析で **Passionfruitの唯一の差別化（◎）** と評価した部分です。改めて整理すると：

| 観点 | Passionfruit | Profound / Peec | AIOutmation(提案) |
|---|---|---|---|
| GA4収益アトリビューション | ◎ ネイティブ統合 | △ / ✕ | **◎ + CRM(Salesforce)連携でBtoBパイプラインまで** |

**重要な気づき**：Passionfruitの収益アトリビューションは、主に **EC・DTC（GA4のpurchase/revenueイベントが取りやすい業態）** に最適化されています。

一方、私たちのターゲット（**上場B2B SaaS**）では、購買は「サイト即購入」ではなく「**商談→数ヶ月の検討→受注**」のため、GA4だけでは追えません。そこでAIOutmationは：

- GA4連携（Passionfruit同等）に加えて
- **Salesforce / HubSpot連携で MQL→SQL→受注までアトリビューション**
- GSCのブランド検索ハロー測定

を組み合わせることで、**B2B特有の長い購買サイクルに対応した収益証明**が差別化軸になります。これは MVP仕様書の F7「AIコメンタリー」＋V2「ROIアトリビューション」で実装予定の領域です。

---

## Sources

- [The GEO Funnel: How AI Search Queries Turn Into Revenue – Passionfruit](https://www.getpassionfruit.com/blog/how-ai-search-queries-turn-into-revenue-(and-how-to-track-every-step))
- [Track AI Search Traffic in GA4 Setup Guide – Passionfruit](https://www.getpassionfruit.com/blog/track-ai-search-traffic-in-ga4-setup-guide-and-best-practices)
- [Passionfruit Labs Review – ecommerceguide.com](https://ecommerceguide.com/apps/passionfruit-labs/)
- [Passionfruit Labs Reviews – CheckThat.ai](https://checkthat.ai/brands/passionfruit-labs/reviews)
- [How to Track ChatGPT, Perplexity, AI Overviews Traffic in GA4 – Discovered Labs](https://discoveredlabs.com/blog/how-to-track-chatgpt-perplexity-and-ai-overviews-traffic-in-ga4-without-guessing)
