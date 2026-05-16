// 宇宙戦略基金 採択機関協業マッピング Wordレポート
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
        AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink, HeadingLevel,
        BorderStyle, WidthType, ShadingType, PageNumber, PageBreak, TabStopType, TabStopPosition } = require('docx');

// ====== 共通ユーティリティ ======
const F = "Yu Gothic"; // 日本語フォント

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: [new TextRun({ text, font: F, size: 22, ...(opts.run || {}) })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "1F3864", space: 4 } },
    children: [new TextRun({ text, font: F, size: 32, bold: true, color: "1F3864" })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: F, size: 26, bold: true, color: "2E75B6" })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: F, size: 23, bold: true, color: "333333" })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: F, size: 22 })],
  });
}

function bulletRich(runs) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: runs,
  });
}

function link(text, url) {
  return new ExternalHyperlink({
    children: [new TextRun({ text, font: F, size: 20, color: "0563C1", underline: {} })],
    link: url,
  });
}

const tBorder = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const tBorders = { top: tBorder, bottom: tBorder, left: tBorder, right: tBorder,
                   insideHorizontal: tBorder, insideVertical: tBorder };

function tCell(text, opts = {}) {
  const { width, fill, bold = false, color = "000000", size = 20 } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    children: [new Paragraph({
      children: [new TextRun({ text, font: F, size, bold, color })],
    })],
  });
}

// ===== 既存協業データ（厳選40件） =====
const cases = [
  { date: "2026-05", parties: "ElevationSpace × 日本低軌道社中", field: "探査等",
    summary: "ELS-RS(回収)とHTV-XC/Japan Module(輸送・補給)を統合し、ポストISS時代の低軌道インフラを構築するMoU。打上げ～補給～実験～回収を一気通貫で提供。",
    url: "https://elevation-space.com/posts/news_20260511" },
  { date: "2025-09", parties: "ElevationSpace × ispace", field: "探査等",
    summary: "民間日本初の月面サンプルリターン実現に向けたMoU。ispaceの軌道間輸送機(OTV)とElevationSpaceの再突入回収カプセルを組み合わせ、共同実証。",
    url: "https://prtimes.jp/main/html/rd/p/000000067.000074085.html" },
  { date: "2026-01", parties: "三菱重工 × 三菱電機 × 三井物産 × 日本低軌道社中", field: "探査等",
    summary: "三井物産100%子会社の日本低軌道社中に三菱重工・三菱電機が第三者割当出資。HTV-XCインテグレーション、与圧モジュール、ランデブー・ドッキングを結集。",
    url: "https://www.mitsubishielectric.co.jp/ja/pr/2026/pdf/0126.pdf" },
  { date: "2025-04", parties: "日本郵船 × 三菱重工業", field: "宇宙輸送",
    summary: "海運初の宇宙戦略基金採択。再使用型ロケットの洋上回収システムを共同研究開発。回収船+司令船の2隻構成、2028年度実証予定。",
    url: "https://www.nyk.com/news/2025/20250724_02.html" },
  { date: "2025-06", parties: "将来宇宙輸送システム × SPACE COTAN", field: "宇宙輸送",
    summary: "HOSPO(北海道スペースポート)の射場開発・利用に関するMoU締結。ASCA1.2試験機の打上げ運用計画を共同検討。",
    url: "https://innovative-space-carrier.co.jp/news/20250616_1" },
  { date: "2025-12", parties: "SPACE COTAN × 三井物産", field: "宇宙輸送",
    summary: "HOSPOを核とした「宇宙版シリコンバレー構想」実現に向けたMoU。北海道の宇宙関連産業集積・観光促進を共同推進。",
    url: "https://prtimes.jp/main/html/rd/p/000000150.000078016.html" },
  { date: "2025-05", parties: "スペースワン × Space BD", field: "宇宙輸送",
    summary: "防衛省「多軌道観測実証衛星」打上げ輸送サービスを共同受注。Space BDが業務受注し、スペースワンのカイロスロケットで2026年度打上げ予定。",
    url: "https://www.space-one.co.jp/news/news_20250528.html" },
  { date: "2025-01", parties: "インターステラテクノロジズ × ウーブン・バイ・トヨタ", field: "宇宙輸送",
    summary: "約70億円の資本業務提携。トヨタの量産技術をロケット「ZERO」量産化に活用。",
    url: "https://toyotatimes.jp/newscast/139.html" },
  { date: "2025-02", parties: "アークエッジ・スペース × スカパーJSAT", field: "衛星等",
    summary: "超小型衛星コンステレーション商用化加速のための業務提携。衛星管制・地上局相互利用・事業開発で協業。",
    url: "https://prtimes.jp/main/html/rd/p/000000045.000073065.html" },
  { date: "2025-10", parties: "アークエッジ × ソフトバンク × NICT × 清原光学", field: "衛星等",
    summary: "宇宙-成層圏・宇宙-地上間の光無線通信実証4社連携。2026年実証衛星打上げ、2027年HAPS-LEO双方向10Gbps光通信を目指す。",
    url: "https://space-connect.jp/ntn-arkedge/" },
  { date: "2024-12", parties: "三菱電機 × Synspective", field: "衛星等",
    summary: "60億円出資・戦略的パートナーシップ覚書。三菱電機が筆頭株主となり、Synspectiveの小型SAR画像を安全保障用途に共同販売。",
    url: "https://www.mitsubishielectric.co.jp/ja/pr/2024/1219/" },
  { date: "2025-12", parties: "三菱電機+スカパーJSAT+アクセル+Synspective+QPS", field: "衛星等",
    summary: "防衛省「衛星コンステレーション整備・運営等事業」共同受注。5年2,831億円。SAR=Synspective+QPS、光学=アクセルスペース担当。",
    url: "https://www.mod.go.jp/j/press/news/2025/12/24a.html" },
  { date: "2025-02", parties: "Marble Visions × NTTデータ × PASCO × キヤノン電子", field: "衛星等",
    summary: "高分解能・高頻度光学衛星8機コンステレーション開発合意。資本業務提携。",
    url: "https://www.nttdata.com/global/ja/news/release/2025/022500/" },
  { date: "2025-03", parties: "Space Compass × ESA(欧州宇宙機関)", field: "衛星等",
    summary: "衛星間光通信ネットワークの軌道上共同実証における相互運用性検討覚書。",
    url: "https://space-compass.com/news/000075.html" },
  { date: "2025-05", parties: "アストロスケール × Honda", field: "衛星等",
    summary: "軌道上衛星給油口接続システム共同開発。HondaのロボティクスとアストロスケールのRPOD技術を融合、2029年に低軌道燃料補給技術実証。",
    url: "https://global.honda/jp/topics/2025/c_2025-05-30.html" },
  { date: "2025-12", parties: "アクセルスペース × Pale Blue", field: "衛星等",
    summary: "ホールスラスタ「PBH-100」の2027年軌道上実証契約・協業深化MoU。AxelLiner Laboratory利用。",
    url: "https://www.axelspace.com/ja/news/paleblue/" },
  { date: "2025-11", parties: "立命館 × コマツ × ispace", field: "探査等",
    summary: "宇宙戦略基金「月面拠点建設のための測量・地盤調査技術」共同採択。地形データ取得・レゴリス調査・道路建設を体系化。",
    url: "https://www.komatsu.jp/ja/newsroom/2025/20251121" },
  { date: "2026-02", parties: "東京大学 × 京都大学 × 立命館 × トプコンほか", field: "探査等",
    summary: "「水・金属元素探査装置のフライトモデル開発」(LUNAR-RABBIT)共同採択。月面資源量の実測を目指す。",
    url: "https://www.s.u-tokyo.ac.jp/ja/info/11065/" },
  { date: "2024-11", parties: "KDDI × 福井工業大学 × ispace", field: "探査等",
    summary: "「月-地球間通信システム開発・実証FS」採択。福井工大あわら13.5m地上局活用、ispaceは関連調査を受託。",
    url: "https://newsroom.kddi.com/news/detail/kddi_nr-415_3688.html" },
  { date: "2026-04", parties: "横浜国立大 × 慶應 × 東京大学", field: "衛星等",
    summary: "「空間自在移動／宇宙ロジスティクス」共同採択。軌道力学・宇宙機システム工学・物流需要予測を統合。",
    url: "https://www.t.u-tokyo.ac.jp/topics/tp2026-04-14-001" },
  { date: "2025", parties: "NeSTRA × ElevationSpace × 藤倉航装 × 金沢工業大", field: "探査等",
    summary: "「展開型エアロシェル技術の地球大気圏突入実証と火星着陸機適用」(Mars Touch Project)を共同実施。",
    url: "https://prtimes.jp/main/html/rd/p/000000063.000074085.html" },
  { date: "2025-02", parties: "Space Tech Accelerator × 衛星データサービス企画 × 三菱電機 × 大日本印刷", field: "衛星等",
    summary: "「衛星データ利用システム海外実証FS」採択。インドネシアのパームヤシ農家支援アプリ共同開発。",
    url: "https://www.jaxa.jp/press/2025/02/20250207-1_j.html" },
  { date: "2026-03", parties: "RESTEC × アクセル × Synspective × PASCO × NSI", field: "衛星等",
    summary: "「衛星データ利用システム実装加速化事業」コンソーシアム採択。SAR・光学センサの評価・校正・検証手法の環境整備。",
    url: "https://www.axelspace.com/news/spacestrategyfund_satellitedata/" },
  { date: "2026-03", parties: "アクセルスペース × 明星電気 × ANAHD × JIJ", field: "衛星等",
    summary: "「次世代地球観測衛星に向けた観測機能高度化技術」コンソーシアム採択。衛星編隊・旅客機観測でCO2モニタリング。",
    url: "https://prtimes.jp/main/html/rd/p/000000058.000066150.html" },
  { date: "2025-01", parties: "スペースデータ × IHI", field: "衛星等",
    summary: "協業MoU締結。IHIの小型衛星コンステレーションデータとスペースデータのデジタルツインを融合し新事業創出。",
    url: "https://www.nikkei.com/article/DGXZQOUC302R40Q5A131C2000000/" },
  { date: "2026-02", parties: "早稲田 × 慶應 × 東京理科大 × JAMSS × パナソニック × ジャムコほか", field: "分野共通",
    summary: "SX-CRANE「宇宙QOL向上」拠点共同採択。私大唯一の代表機関として宇宙居住・医療領域を牽引。",
    url: "https://www.keio.ac.jp/ja/press-releases/2026/3/4/28-172973/" },
  { date: "2025", parties: "立命館 × 東京大学 × 島津製作所 × 会津大ほか", field: "分野共通",
    summary: "SX研究開発拠点「月面探査・利用を産業化するための宇宙機器開発・人材育成」共同採択。",
    url: "https://www.ritsumei.ac.jp/news/detail/?id=4081" },
  { date: "2025-06", parties: "名古屋大 × 慶應 × 東大 × 横国 × ネッツほか", field: "分野共通",
    summary: "SX研究開発拠点「デトネーションエンジン・宇宙推進工学革新研究拠点」キックオフ。観測ロケットS-520-31で実証済。",
    url: "https://www.nagoya-u.ac.jp/info/press/sx.html" },
  { date: "2026-03", parties: "大熊ダイヤモンドデバイス × Space BD", field: "分野共通",
    summary: "SX-ARK「ダイヤモンド半導体による小型SARの熱制約打破」採択。Space BDが協力機関として参画。",
    url: "https://sorae.info/biz/20260410-spacebd-sx-ark.html" },
  { date: "2025-09", parties: "エネコート × トヨタ × 日揮 × KDDI × 京都大学", field: "分野共通",
    summary: "ペロブスカイト太陽電池の量産・実装に向けた産学連合(100億円規模)。SX-ARK(熱とデバイス)にも採択。",
    url: "https://enecoat.com/news/20250910/" },
  { date: "2025-09", parties: "ElevationSpace × Exobiosphere(ルクセンブルク)", field: "探査等",
    summary: "高頻度宇宙サンプルリターンMoU。創薬・バイオ実証で連携。",
    url: "https://www.exobiosphere.com/news/elevationspace-and-exobiosphere-sign-mou-to-enable-high-frequency-sample-return-from-space" },
  { date: "2025-09", parties: "ispace × 三井住友銀行", field: "探査等",
    summary: "SMBCがHAKUTO-R初のオフィシャルパートナーに参画。協調融資100億円実行。シスルナ経済圏ビジョン共有。",
    url: "https://prtimes.jp/main/html/rd/p/000000029.000140640.html" },
  { date: "2025-10", parties: "NICT × スカパーJSAT × 東京大学 × 日本無線", field: "衛星等",
    summary: "GEO/LEO衛星線と地上線を含む衛星×5Gネットワーク統合運用実証実験成功。",
    url: "https://www.nict.go.jp/press/2025/10/02-1.html" },
  { date: "2025-07", parties: "Synspective × Spectee", field: "衛星等",
    summary: "SARによる広域浸水データとSpecteeのSNS解析を統合し災害対応高度化。",
    url: "https://synspective.com/press-release/2025/spectee_partnership/" },
  { date: "2024-09", parties: "将来宇宙輸送システム × 荏原製作所", field: "宇宙輸送",
    summary: "電動ポンプ活用ロケットエンジン共同開発の包括連携協定。2026年3月着火試験成功。",
    url: "https://innovative-space-carrier.co.jp/news/20240927" },
  { date: "2025-03", parties: "三菱重工 × UACJ × 富山住友電工", field: "宇宙輸送",
    summary: "ロケット向け高強度アルミ合金(Sc入りアルミ合金ワイヤー)のWAAM研究開発。日本初のアルミ-スカンジウム合金のロケット実用化を目指す。",
    url: "https://www.uacj.co.jp/release/20250303.html" },
  { date: "2026-02", parties: "アストロスケール × 三井住友ファイナンス&リース", field: "衛星等",
    summary: "軌道上サービスを利用した衛星オペレーティング・リース事業など、衛星二次利用マーケット創出のMoU。",
    url: "https://prtimes.jp/main/html/rd/p/000000169.000084204.html" },
  { date: "2025-12", parties: "ElevationSpace × JAXA", field: "探査等",
    summary: "JAXAパートナースタートアップに認定。大気圏再突入・回収技術の知見支援を受け開発を加速。",
    url: "https://prtimes.jp/main/html/rd/p/000000071.000074085.html" },
  { date: "2025", parties: "アストロスケール × 三井物産", field: "探査等",
    summary: "三井物産パートナー企業として「日本モジュール」事業化検討に参画。軌道上点検・修理・燃料補給サービスを検討。",
    url: "https://prtimes.jp/main/html/rd/p/000000056.000067481.html" },
  { date: "2025-05", parties: "Star Signal Solutions × Inovor Technologies(豪)", field: "衛星等",
    summary: "SSA・宇宙交通協調・衛星通信分野での豪日協力強化MoU。",
    url: "https://www.inovor.com.au/inovor-technologies-and-star-signal-solutions-sign-strategic-mou-to-strengthen-australia-japan-space-industry-collaborations/" },
];

// ===== 既存協業テーブル =====
function buildCasesTable() {
  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        tCell("時期", { width: 900, fill: "1F3864", bold: true, color: "FFFFFF", size: 18 }),
        tCell("関与機関", { width: 3400, fill: "1F3864", bold: true, color: "FFFFFF", size: 18 }),
        tCell("分野", { width: 900, fill: "1F3864", bold: true, color: "FFFFFF", size: 18 }),
        tCell("概要 / 出典", { width: 4160, fill: "1F3864", bold: true, color: "FFFFFF", size: 18 }),
      ],
    }),
  ];
  for (const c of cases) {
    const fillMap = { "宇宙輸送": "FCE4D6", "衛星等": "DDEBF7", "探査等": "E2EFDA", "分野共通": "FFF2CC" };
    const fill = fillMap[c.field] || "FFFFFF";
    rows.push(new TableRow({
      cantSplit: true,
      children: [
        tCell(c.date, { width: 900, fill, size: 18 }),
        tCell(c.parties, { width: 3400, fill, bold: true, size: 18 }),
        tCell(c.field, { width: 900, fill, size: 18 }),
        new TableCell({
          width: { size: 4160, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          shading: { fill, type: ShadingType.CLEAR, color: "auto" },
          children: [
            new Paragraph({ children: [new TextRun({ text: c.summary, font: F, size: 18 })] }),
            new Paragraph({ spacing: { before: 60 }, children: [link("[出典]", c.url)] }),
          ],
        }),
      ],
    }));
  }
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [900, 3400, 900, 4160],
    borders: tBorders,
    rows,
  });
}

// ===== 将来協業候補（4観点で整理） =====
const futureCases = [
  {
    title: "①「日本版シスルナ・ロジスティクス」: Elevation Space × ispace × 低軌道社中 × アストロスケール",
    body: [
      "回収(Elevation Space)、月軌道間輸送(ispace)、LEOステーション(日本低軌道社中)、軌道上燃料補給(アストロスケール)が連携することで、月-LEO-地上を貫く軌道輸送パッケージが完成する。",
      "技術補完性: 4社の宇宙戦略基金採択テーマが上下流で完全に補完的(打上げ→補給→実験→回収)。",
      "バリューチェーン: 上流〜下流までの完全統合型サービス。Space BDがブッキング窓口に。",
      "海外展開: ESA/NASA/CSA向けに月面サンプルリターンの統合サービスとして売り込み可能。Elevation Space×ispaceのMoU、アストロスケール×インド3社の流れが先行。",
      "事業化: 2030年前後の商用月面サンプルリターン市場・LEO実験回収市場で年間複数ミッションのオペレーターを目指す。",
      "想定される未来像: 打上げ・補給・実験・回収を一気通貫提供する世界唯一のサービス。米SpaceX×Axiom、欧ArianeGroup×Thalesに対し、日本連合が独自ポジションを確立。"
    ]
  },
  {
    title: "② 「日本版Earth Observation as a Service」: Synspective × QPS × アクセルスペース × Marble Visions",
    body: [
      "SAR(Synspective/QPS)、光学(アクセル)、ハイパースペクトル(Marble Visions)で全天候・高分解能・分光の地球観測スペクトルが揃う。",
      "技術補完性: 各社が異なる観測モダリティを持ち、組合せで「日本版Sentinelミッション」相当の体制が可能。",
      "バリューチェーン: 衛星製造→運用→データ販売(RESTEC/PASCO)→AI/ソリューション(Tellus/PFN/スペースデータ)の縦串。すでにRESTEC主導のコンソーシアムが2026年3月に採択済。",
      "海外展開: 防衛省コンステ事業を基盤に、ASEAN・中東・アフリカへ「日本版センチネル」を提案。アクセルスペース×Geoimageの豪州展開、Synspective×Specteeの災害アプリ実績を活用。",
      "事業化: 安全保障+商用の二重需要で年商数百億円規模のセグメント形成。クロスユーのアフリカパートナーシップでODA案件と接続。",
      "想定される未来像: 米Maxar/Planetに対抗する第3極として、日本連合が「中立で信頼できる衛星データ提供国」のポジションを確保。"
    ]
  },
  {
    title: "③ 「日本打上げトライアングル」: 将来宇宙輸送 × インターステラ × スペースワン + 日本郵船 × SPACE COTAN",
    body: [
      "ベンチャー再使用機(ISC)、量産小型機(IST)、固体ロケット(スペースワン)、洋上回収(日本郵船)、射場(SPACE COTAN)で打上げポートフォリオを形成。",
      "技術補完性: 大中小・固体液体・陸上洋上のあらゆる打上げ需要をカバー。一機種の事故・遅延に対する冗長性も確保。",
      "バリューチェーン: 部品(IHI/UACJ/丸八)→機体(各社)→射場(SPACE COTAN/紀伊)→回収(日本郵船)→運用。日本初の純国産・全層型バリューチェーンが完成。",
      "海外展開: 東南アジア・中東衛星オペレータ向けに「日本打上げトリオ」(質量・予算別に最適機種を提案)としてセット販売。",
      "事業化: 年間20本以上の国内打上げ需要が成立。SPACE COTANの「宇宙版シリコンバレー構想」と接続して産業集積に。",
      "想定される未来像: アジア最大の打上げハブとして、価格・冗長性で米SpaceXとの併用先に。日本の打上げ自立性が国家インフラとして確保。"
    ]
  },
  {
    title: "④ 「宇宙コンピューティング・バックボーン」: Space Compass × NICT × アークエッジ × ワープスペース × 各観測衛星社",
    body: [
      "光通信地上系(NICT)、光端末(ワープスペース)、LEO小型衛星(アークエッジ等)、データソース(Synspective/QPS)、中継(Space Compass)が連結。",
      "技術補完性: 周波数・光通信・端末・衛星・運用がすべて国産で揃う。NICTの宇宙量子暗号通信、Space Compassの光リレー、アークエッジ×ソフトバンクのHAPS-LEO通信が相互補完。",
      "バリューチェーン: 研究(NICT)→端末(ワープ)→衛星(各社)→中継(Space Compass)→顧客(防衛・商用・通信MNO)。",
      "海外展開: Space Compass×ESA、×Hellas Sat、×JSAT International×Apolinkとのグローバル光通信メッシュにつなぎ込む国際ハブ機能。",
      "事業化: 2030年に光通信トラフィック契約で年商数百億円。データ送信レート10Gbps級が普及することで衛星画像のリアルタイム伝送が標準化。",
      "想定される未来像: 6G時代の宇宙コンピューティング基盤として、RF依存から光通信時代への移行を日本連合が主導。楽天モバイル等のMNOとの統合NTNサービスも誕生。"
    ]
  },
  {
    title: "⑤ 「日本版Refuel & Deorbit as a Service」: アストロスケール × Pale Blue × IHI × Star Signal Solutions",
    body: [
      "水推進(Pale Blue)、燃料補給(アストロスケール)、SSA(IHI/Star Signal/パワーレーザー)で軌道機動 & 監視のフルパッケージが揃う。",
      "技術補完性: 推進・接続・観測・除去の各レイヤーが基金採択テーマで縦串化。アストロスケール×Honda、×Exotrailの先行例を拡張。",
      "バリューチェーン: OTV(Pale Blue/三菱電機)→燃料補給(アストロスケール)→デブリ除去(アストロスケール)→監視(IHI/Star Signal/パワーレーザー)→廃棄処理。",
      "海外展開: 英アストロスケール本社経由でESA・UK Space Agency案件にアクセス、米FAA/Space ForceサービスバウンドにDigantara/Bellatrixと共同提案。",
      "事業化: 三井住友FLとのMoUで衛星オペレーティング・リース市場が立ち上がる。2030年代の「軌道上サービス市場」(数千億円)で先行者利益。",
      "想定される未来像: 世界初の「Refuel as a Service + Deorbit as a Service」のパッケージを提供する日本連合。サステナブル軌道経済の国際標準化を主導。"
    ]
  },
  {
    title: "⑥ 「月面水道事業」: 東京大学(資源)× 高砂熱学 × 栗田工業 × 横河電機 × ispace",
    body: [
      "月面の水資源探査(東大/ispace)→電解水素生成(高砂熱学)→水処理(栗田工業)→計装(横河電機)。すでに東大「宇宙資源連携研究機構」の連携機関として活動。",
      "技術補完性: 月の水を「発見→採取→精製→利用」へ垂直統合。地上の水処理大手の技術ノウハウが宇宙転用。",
      "バリューチェーン: 月面ISRU装置の縦串。トヨタ再生型燃料電池、JAEA原子力電池とも組合せて月面拠点の電力源と統合可能。",
      "海外展開: NASA Artemis計画/ESA Argonaut/CSA Lunar Gatewayに日本版ISRU装置を提案。HAKUTO-R Mission2で水電解実証済。",
      "事業化: 2030年代の月面ISRU市場(米Astrolab/iSpace USなどが参入する数千億円市場)で装置・サービス事業を確立。",
      "想定される未来像: 月の水を燃料・生命維持・電力に変換する「月の水道事業」を日本連合が確立。Artemisに不可欠な装置サプライヤーに。"
    ]
  },
  {
    title: "⑦ 「日本版LunaNet」: KDDI × 福井工業大 × NICT × ispace × Space Compass",
    body: [
      "月-地球間通信(KDDI/福井工大)と光リレー(NICT/Space Compass)、月面ノード(ispace/アークエッジ月測位)で完全な月通信網。",
      "技術補完性: 地上局(KDDI)→中継衛星(Space Compass)→月軌道(ispace)→月面(ispace)の通信スタック。月測位(アークエッジ)を加えて月版GPS的機能も。",
      "バリューチェーン: 通信MNO(KDDI)が月通信のキャリアサービスを提供し、ispaceミッションが顧客でもありノードでもあるシナジー。",
      "海外展開: NASA LunaNet/ESA Moonlight計画とのインターオペラビリティ確保。日米欧Artemis同盟内の役割分担を獲得。",
      "事業化: 2030年代の月探査支援通信市場(年数百億円)。月面産業の前提インフラとして必須。",
      "想定される未来像: Artemis計画における日本の独自性が「月通信網のキャリア」として確立。地上のKDDI事業との接続で連続的なシスルナ通信が実現。"
    ]
  },
  {
    title: "⑧ 「日本ホスピタリティ宇宙旅行」: 岩谷技研 × ISC × JAMSS × ジャムコ × 早稲田大学",
    body: [
      "気球(岩谷技研)、サブオービタル(将来宇宙輸送システム)、与圧モジュール(JAMSS)、内装(ジャムコ)、QOL設計(早稲田大学)で有人宇宙旅行・居住の体験設計が成立。",
      "技術補完性: 機体技術と人間中心設計が連結。岩谷技研×JAL、早稲田×東京女子医科大の医療連携の蓄積が活用される。",
      "バリューチェーン: 機体(岩谷/ISC)→キャビン(JAMSS/ジャムコ)→運用支援(早稲田研)→顧客提供。気球→サブオービタル→軌道とアップグレードする商品階段。",
      "海外展開: スイス/UAE/インド富裕層向け「日本式ホスピタリティ宇宙旅行」として販売。Axiom Space/Orbital ReefなどポストISS民間ステーションへの装備供給。",
      "事業化: 2030年代に年間数百回の有人準軌道飛行で観光・実験・PR収益を確立。",
      "想定される未来像: 米Virgin Galactic/Blue Originとの差別化要素として「日本のおもてなし」を宇宙ブランド化。SX-CRANE早稲田拠点の研究成果が世界の宇宙居住標準に。"
    ]
  },
  {
    title: "⑨ 「衛星データ×AI×ODA」: 衛星オペレータ × 業務SaaS × 国際金融",
    body: [
      "上流の衛星データ(Synspective/QPS/アクセル)と下流の業務アプリ(ウミトロン/オーシャンソリューションテクノロジー/パシフィックコンサルタンツ/Solafune)が直結。",
      "技術補完性: 同じ事業で複数衛星社のデータを横断利用するモデル(RESTECコンソーシアムのアプローチ)が、海外社会実装プロジェクトに展開可能。",
      "バリューチェーン: 衛星画像→AI解析→産業別アプリ(漁業/IUU/防災/林業)→ASEAN/アフリカの政府・企業顧客→JICA/世銀資金で社会実装。",
      "海外展開: インドネシア(パームヤシ)、フィリピン(IUU)、ベトナム、アフリカ各国で衛星×SaaSをパッケージ販売。デロイト、野村総研、三菱総研が現地調整。",
      "事業化: 2028年に海外売上比率50%超を目標。SDGs/TNFD関連の社会実装案件で安定的に収益化。",
      "想定される未来像: 日本発「衛星データ×AI×国際協力」の社会実装モデルが世界標準に。Solafuneのコンペでアフリカの若手データサイエンティスト育成も。"
    ]
  },
  {
    title: "⑩ 「Forest-as-an-Asset」: 住友林業 × Green Carbon × Archeda × Synspective × アクセル × 国際航業 × 東京海上レジリエンス",
    body: [
      "森林・農地(住友林業/Green Carbon)、衛星解析(Synspective/アクセル/Archeda)、地理空間(国際航業)、保険(東京海上)でカーボン金融×宇宙の縦串。",
      "技術補完性: 計測(SAR/光学)→検証(MRV)→クレジット組成→保険・金融商品化が一体化。住友林業×IHI(NeXT Forest)、Archeda×Green Carbonの先行例を本格事業化。",
      "バリューチェーン: 森林管理→衛星モニタリング→MRV→クレジット発行→保険付帯→流通(MUFG等)。",
      "海外展開: COP/世銀向けにアジア森林・農地クレジットの認証システムを提案。RESTEC×ITTOのMoUを発展形に拡張。",
      "事業化: 2028年に年間100万トン規模の高品質クレジット発行。TNFD開示需要と直結し、上場企業向け年商数十億円ビジネスに。",
      "想定される未来像: 日本発「Forest-as-an-Asset」プラットフォームがグリーン金融×宇宙データの先進事例として国際展開。"
    ]
  },
  {
    title: "⑪ 「衛星部品の国産化フルセット」: エネコート × 大熊ダイヤモンドデバイス × 三洋化成 × NECスペーステクノロジー × シャープエネルギーソリューション",
    body: [
      "ペロブスカイト太陽電池×ダイヤモンド半導体×電解質材料×衛星部品×太陽電池モジュール。SX-ARK「熱とデバイス」採択機関群が連結。",
      "技術補完性: 衛星の3大重要部品(電源・冷却・通信)に必要な日本独自の材料が揃う。米中依存からの脱却を実現。",
      "バリューチェーン: 材料(エネコート/大熊/三洋化成)→部品(NECスペース/シャープ)→衛星(アクセル/Synspective)→運用。",
      "海外展開: 欧米衛星メーカーへの戦略物資としての輸出。経済安全保障案件として米国IRA/CHIPS法と連動した取引も可能。",
      "事業化: 2030年に衛星部品市場で日本シェアを5%→15%へ拡大。エネコートはトヨタ・日揮・KDDIとの産学連合(地上量産)と相乗効果。",
      "想定される未来像: 日本の素材・材料力が「宇宙の半導体・電池」として復活し、衛星部品の輸出産業化。"
    ]
  },
  {
    title: "⑫ 「月面電気・水・道路」: トヨタ × JAEA × IHIエアロスペース × 東大(資源) × 立命館 × コマツ",
    body: [
      "燃料電池(トヨタ)、原子力電池(JAEA)、物資補給ドッキング(IHIエアロ)が月面拠点の動力源を提供。月面建設(コマツ/立命館)が消費先になる完全な動力-建設エコシステム。",
      "技術補完性: 動力源の冗長性(発電・蓄電・補給)と建設・利用シーンの統合。立命館×JAXA連携協定の有人与圧ローバ研究を中核に。",
      "バリューチェーン: 電源(トヨタ/JAEA)→補給(IHIエアロ)→建設(コマツ)→拠点(立命館)→利用(東大資源)。",
      "海外展開: NASA Artemis有人探査ローバ(LUNAR CRUISER)を経由した米国市場アクセス。Artemis VII以降の月面産業の必須プレイヤーに。",
      "事業化: 2030年代の月面拠点・有人ローバの動力・建設市場で日本連合が独占的地位。",
      "想定される未来像: 月面拠点の「電気・水・道路」を日本連合が供給する月版コンソーシアム。Artemis計画の中核装置サプライヤーに。"
    ]
  },
  {
    title: "⑬ 「6G NTNコンソーシアム」: 楽天モバイル × アークエッジ × Space Compass × ソフトバンク × NICT",
    body: [
      "楽天モバイル衛星×Astera(地上)、アークエッジLEO、Space Compass光リレー、NICT周波数技術で6G/NTN実装。",
      "技術補完性: 周波数(NICT)→中継(Space Compass)→LEO(アークエッジ)→地上MNO(楽天/SBM)→顧客。「衛星通信と地上ネットワークの統合運用」テーマ群が結合。",
      "バリューチェーン: 衛星オペレータ+MNOが連合し、IoTおよびスマートフォンSatellite Directのインフラを提供。",
      "海外展開: 東南アジア・アフリカの未電化地域向けにNTN-MNOパッケージを輸出。災害時の代替通信としても各国政府に提案。",
      "事業化: 6Gコア技術として2030年代に世界標準化に貢献。ARPU高い法人IoT/防災案件で安定収益。",
      "想定される未来像: 日本発「Non-Terrestrial Network」が6G世界標準に。通信衛星と地上MNOの統合運用で世界をリード。"
    ]
  },
  {
    title: "⑭ 「LEO Lab as a Service」: 低軌道社中 × Elevation Space × ispace × Space BD × 各大学拠点",
    body: [
      "LEOステーション(低軌道社中)・回収(Elevation)・月(ispace)に大学の実験(医薬/食料/材料)をブッキング(Space BD)。SX-CRANE/SX-ARK各拠点の研究成果が需要源に。",
      "技術補完性: 実験提案(大学)→ブッキング(Space BD)→輸送(各社)→回収(Elevation)→分析。早稲田QOL、東女医大医療、山形大宇宙食、東北大材料が需要源。",
      "バリューチェーン: 研究者→アクセラレータ(Space BD)→宇宙アセット→地上分析。RESTEC・産総研も認証で関与。",
      "海外展開: ESA/NASAだけでなくシンガポール/タイ/UAEの大学・製薬会社にも実験枠を提供。Exobiosphereのような海外バイオ企業とも連携。",
      "事業化: LEOでの微小重力実験市場(年数百億円)で「日本枠」を獲得。Sarmonyの「革新的衛星ミッション技術実証支援」も同枠で展開可能。",
      "想定される未来像: 日本版「LEO Lab as a Service」が創薬・素材・宇宙食研究の世界拠点に。ポストISS時代の日本のソフトパワー源泉。"
    ]
  },
  {
    title: "⑮ 「日の丸ロケット素材」: 東レ × UACJ × 東レ・カーボンマジック × コンポジットテーラーズ × 丸八 × 三菱重工 × IHI",
    body: [
      "炭素繊維(東レ/CTL/カーボンマジック)、アルミ合金(UACJ)、CFRPタンク(丸八)、機体(三菱重/IHI)で「日本素材機体」の純国産化。",
      "技術補完性: 上流素材から中間素材・部品・機体までを国内で完結。UACJ×三菱重工Sc合金、丸八×東大熱可塑CFRPがその先行例。",
      "バリューチェーン: 素材→中間素材→部品→機体→打上げ→運用までの完全純国産。",
      "海外展開: 国際素材市場で「日本製ロケット素材」が標準に。欧州ロケットメーカー(ArianeGroup/RFA等)にも輸出可能。",
      "事業化: 2030年代に素材輸出で年商数百億円。地上の自動車・航空機軽量化需要にも転用。",
      "想定される未来像: 「日の丸ロケット」が素材→機体→打上げまで全て国産化された強靭なサプライチェーンを完成。経済安全保障の象徴的事例。"
    ]
  },
];

// 将来候補テーブル/段落生成
function buildFutureSection() {
  const items = [];
  for (const f of futureCases) {
    items.push(h3(f.title));
    items.push(p(f.body[0]));
    // 4観点 + 想定未来像
    const labels = ["【技術補完性】", "【バリューチェーン】", "【海外展開】", "【事業化】", "【想定される未来像】"];
    for (let i = 1; i < f.body.length; i++) {
      const label = labels[i - 1];
      items.push(new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({ text: label, font: F, size: 22, bold: true, color: "C00000" }),
          new TextRun({ text: " " + f.body[i].replace(/^(技術補完性|バリューチェーン|海外展開|事業化|想定される未来像)[:：]\s*/, ""), font: F, size: 22 }),
        ],
      }));
    }
  }
  return items;
}

// ===== ハブ機関表 =====
const hubs = [
  { name: "ElevationSpace", count: "4-5件", partners: "ispace, 日本低軌道社中, Exobiosphere, NeSTRA, JAXA, 東北大学", role: "低軌道-月面の回収・帰還ハブ" },
  { name: "日本低軌道社中", count: "3-4件", partners: "三井物産, 三菱重工, 三菱電機, ElevationSpace, アストロスケール", role: "ポストISS LEOステーション統合" },
  { name: "アストロスケール", count: "5件以上", partners: "Honda, Exotrail, インド3社, 三井物産, 三井住友FL, エアバス", role: "軌道上サービスの国際展開ハブ" },
  { name: "Space Compass", count: "5件以上", partners: "ESA, Hellas Sat, アクセルスペース, QPS研究所, JSAT International, Apolink", role: "光通信・データリレーの国際ハブ" },
  { name: "アークエッジ・スペース", count: "4-5件", partners: "スカパーJSAT, MUFG, 清水建設, ソフトバンク, NICT, セーレン", role: "小型衛星製造・運用エコシステム" },
  { name: "三菱電機", count: "6件以上", partners: "Synspective, Pale Blue, 日本低軌道社中, スカパー, アクセル, QPS, NEC", role: "日本最大級の宇宙コングロマリット" },
  { name: "ispace", count: "5件以上", partners: "ElevationSpace, 立命館, コマツ, KDDI, SMBC, 高砂熱学, 中央大学", role: "月面探査・産業化のハブ" },
  { name: "立命館大学", count: "3-4件", partners: "コマツ, ispace, 島津製作所, 東京大学, 慶應, JAXA", role: "月面拠点形成のハブ" },
  { name: "SPACE COTAN", count: "3件", partners: "ISC, 三井物産, 岩谷技研, 清水建設", role: "北海道宇宙クラスター(HOSPO)" },
  { name: "東京大学", count: "5件以上", partners: "Pale Blue, 京都大学, 立命館, 慶應, ispace, 高砂熱学", role: "アカデミア中核ハブ" },
];

function buildHubTable() {
  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        tCell("ハブ機関", { width: 1800, fill: "2E75B6", bold: true, color: "FFFFFF" }),
        tCell("連携件数", { width: 1100, fill: "2E75B6", bold: true, color: "FFFFFF" }),
        tCell("主要連携先", { width: 4200, fill: "2E75B6", bold: true, color: "FFFFFF" }),
        tCell("役割・特徴", { width: 2260, fill: "2E75B6", bold: true, color: "FFFFFF" }),
      ],
    }),
  ];
  for (const h of hubs) {
    rows.push(new TableRow({
      cantSplit: true,
      children: [
        tCell(h.name, { width: 1800, bold: true }),
        tCell(h.count, { width: 1100 }),
        tCell(h.partners, { width: 4200, size: 18 }),
        tCell(h.role, { width: 2260, size: 18 }),
      ],
    }));
  }
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1800, 1100, 4200, 2260],
    borders: tBorders,
    rows,
  });
}

// ===== ドキュメント構築 =====
const doc = new Document({
  creator: "Claude",
  title: "宇宙戦略基金 採択機関協業マッピング",
  styles: {
    default: { document: { run: { font: F, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: F, color: "1F3864" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: F, color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 23, bold: true, font: F, color: "333333" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "宇宙戦略基金 採択機関協業マッピング", font: F, size: 18, color: "808080" })],
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", font: F, size: 18, color: "808080" }),
            new TextRun({ children: [PageNumber.CURRENT], font: F, size: 18, color: "808080" }),
          ]
        })]
      })
    },
    children: [
      // 表紙タイトル
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 1200, after: 200 },
        children: [new TextRun({ text: "宇宙戦略基金", font: F, size: 28, bold: true, color: "1F3864" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: "採択機関同士のMoU・協業マッピングと将来協業可能性", font: F, size: 40, bold: true, color: "1F3864" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: "── Elevation Space × 日本低軌道社中 MoU(2026年5月)を起点として ──", font: F, size: 24, italics: true, color: "595959" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 1600, after: 80 },
        children: [new TextRun({ text: "作成日: 2026年5月16日", font: F, size: 22, color: "595959" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "対象機関: 宇宙戦略基金 第一期・第二期 全採択機関(計126機関)", font: F, size: 22, color: "595959" })],
      }),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 1. エグゼクティブサマリー =====
      h1("1. エグゼクティブサマリー"),
      p("2026年5月11日、Elevation Spaceと日本低軌道社中は、宇宙戦略基金の採択テーマを相互連携させる基本合意書(MoU)を締結したと発表した。Elevation Spaceの「高頻度物資回収システム技術」と、日本低軌道社中の「低軌道自律飛行型モジュールシステム技術」「船外利用効率化技術」「HTV-XC」を統合し、ポストISS時代における日本の自立的・国際競争力ある低軌道活動インフラの確立を目指すものだ。"),
      p("本レポートでは、宇宙戦略基金第一期・第二期の全採択機関(126機関)を対象に、(1) 採択機関同士のMoU・協業の現状を網羅的に整理し、(2) 技術補完性・バリューチェーン・海外展開・事業化の4観点で今後の協業可能性を予想する。"),
      h2("主要な発見"),
      bullet("2024-2026年に確認された採択機関を含む主要MoU・協業は40件超。2025年後半から加速度的に増加し、宇宙戦略基金第二期採択(2025-2026年)を契機に「テーマ間連携」を明示的に宣言する事例が顕在化している。"),
      bullet("最も連携ハブとして機能しているのは、Elevation Space(月-LEO回収帰還の結節点)、アストロスケール(国際的軌道上サービス)、Space Compass(光通信)、三菱電機(多面的出資)、ispace(月探査の中核)、東京大学・立命館大学(アカデミア中核)の6機関。"),
      bullet("商用安全保障案件(防衛省衛星コンステレーション5年2,831億円)が、Synspective+QPS+アクセルスペース+三菱電機+スカパーJSATを束ねる強力な「需要側ハブ」として機能している。"),
      bullet("月探査では、ispace×立命館×コマツ×東大×京大の月面産業クラスタが形成され、月面拠点建設・水資源探査・有人ローバの3軸が同時並行で動き始めている。"),
      bullet("Elevation Space×日本低軌道社中型の「採択機関同士のテーマ間統合MoU」は、今後さらに5-10件規模で生まれる可能性が高い。本レポートでは特に有望な15シナリオを抽出した。"),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 2. 起点となるMoU =====
      h1("2. 起点となる事例 — Elevation Space × 日本低軌道社中 MoU"),
      p("2026年5月11日、株式会社ElevationSpaceと株式会社日本低軌道社中は、宇宙戦略基金で採択された両者の技術開発テーマの成果を有機的に連携させることに関する基本合意書(MoU)を締結した。"),
      h2("両社の宇宙戦略基金 採択テーマ"),
      bulletRich([
        new TextRun({ text: "ElevationSpace: ", font: F, size: 22, bold: true }),
        new TextRun({ text: "「高頻度物資回収システム技術」(第二期・探査等)。小型無人回収機ELS-RSを開発。", font: F, size: 22 }),
      ]),
      bulletRich([
        new TextRun({ text: "日本低軌道社中: ", font: F, size: 22, bold: true }),
        new TextRun({ text: "①「国際競争力と自立・自在性を有する物資補給システム(近傍通信)」、②「低軌道自律飛行型モジュールシステム技術」(第一期・探査等)、③「船外利用効率化技術」(第二期・探査等)。HTV-XCとJapan Moduleを軸にポストISS時代のLEOプラットフォームを構築。", font: F, size: 22 }),
      ]),
      h2("MoUの狙い"),
      p("両社は、ポストISS時代における日本の自立的かつ国際競争力ある「低軌道活動インフラ」の確立を目指し、打上げ→物資補給→軌道上実験・実証→物資回収という一連のサービスを有機的に接続することに合意した。これは宇宙戦略基金が想定する「複数テーマの成果を統合し、産業化に資する事業」の最も象徴的な事例となる。"),
      h2("起点としての意義"),
      p("従来、宇宙戦略基金の採択は各テーマ単位で完結する想定だったが、本MoUは「テーマ間連携」を公的に宣言した日本初の事例となる。これは今後、他の採択機関でも同様の連携MoUが連鎖的に発生することを示唆する。実際、Elevation Space自身もispaceとの月面サンプルリターンMoU(2025年9月)、Exobiosphereとのバイオ実証MoU(2025年9月)、NeSTRA・藤倉航装との火星着陸機MoU(2025年)を立て続けに締結しており、宇宙戦略基金採択を契機としたMoU連鎖の中心人物となっている。"),
      p("また、2026年1月には三菱重工・三菱電機・三井物産・日本低軌道社中の4社が連携体制を発表しており、HTV-XCを核とした「ポストISSコンソーシアム」が既に形成されつつある。Elevation Spaceとのテーマ間連携は、このコンソーシアムをLEO上での実験回収まで延長する役割を果たす。"),
      new Paragraph({
        spacing: { before: 200 },
        children: [
          new TextRun({ text: "[出典] ", font: F, size: 20, bold: true }),
          link("ElevationSpace公式発表 (2026年5月11日)", "https://elevation-space.com/posts/news_20260511"),
          new TextRun({ text: " / ", font: F, size: 20 }),
          link("日本低軌道社中 英文発表", "https://japan-leo-shachu.com/en/news/japan-leo-shachu-and-elevationspace-launch-collaboration-to-explore-leo-activity-infrastructure-for-the-post-iss-era/"),
        ],
      }),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 3. 既存MoU・協業の一覧 =====
      h1("3. 採択機関同士の既存MoU・協業(40件)"),
      p("以下は、宇宙戦略基金の採択機関を1社以上含むMoU・業務提携・資本提携・共同採択・出資・包括連携協定など、2024年以降の主要事例である。背景色は分野別(オレンジ:宇宙輸送、青:衛星等、緑:探査等、黄:分野共通)を示す。詳細リスト(58件)はExcel別添「既存MoU・協業一覧」シートを参照。"),
      buildCasesTable(),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 4. 連携ハブ機関 =====
      h1("4. 連携ハブ機関の整理"),
      p("既存MoU・協業の集積を見ると、以下の機関が連携の結節点として機能している。これらの機関を中心に、今後さらなる連携が生まれる蓋然性が高い。"),
      buildHubTable(),
      h2("ハブ機関の構造的特徴"),
      p("(1) スタートアップ系のハブ(Elevation Space、ispace、アストロスケール、Space Compass、アークエッジ・スペース)は、自社の中核技術を周辺機関と組み合わせて「サービス化」する戦略を採っている。"),
      p("(2) 重電・電機系の三菱電機は、Synspective・Pale Blue・日本低軌道社中など複数のスタートアップに出資し、衛星バス・通信・推進・LEOステーションを横断する事業ポートフォリオを形成。NECも空間自在移動・利用テーマで複数案件に関与し、両社が「日本版宇宙コングロマリット」の双璧となりつつある。"),
      p("(3) 大学系の中核ハブ(東京大学、立命館大学)は、SX研究開発拠点やSX-ARK/SX-CRANEを通じて産学連携プラットフォームを提供。特に立命館は月面産業ビジョンを軸に、ispace・コマツ・島津製作所・JAXA有人宇宙技術部門と多面的連携。"),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 5. 今後の協業可能性(4観点) =====
      h1("5. 今後の協業可能性 — 4観点での将来予想"),
      p("以下では、技術補完性・バリューチェーン上の連携・海外展開・国際協力・事業化への道筋という4観点を踏まえ、宇宙戦略基金採択機関同士に生じうる15の主要協業シナリオを予想する。各シナリオは「Elevation Space × 日本低軌道社中」型のテーマ間連携の発展形として、現実的に2026-2030年の間に具体化する可能性が高いものを抽出している。"),
      ...buildFutureSection(),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 6. 総括 =====
      h1("6. 総括 — 日本宇宙バリューチェーンの胎動"),
      p("Elevation Spaceと日本低軌道社中のMoU(2026年5月)を契機として浮き彫りになるのは、宇宙戦略基金が単なる個別技術支援ではなく、「日本独自の宇宙バリューチェーンを統合的に立ち上げるための連携加速装置」として機能し始めているという事実である。"),
      h2("バリューチェーンの形成"),
      p("素材(東レ・UACJ・丸八・東レカーボンマジック等)→部品(NECスペーステクノロジー・三菱電機・シャープエネルギーソリューション等)→機体(三菱重工・IHI・将来宇宙輸送システム・スペースワン・インターステラ等)→打上げサービス(SPACE COTAN・日本郵船等)→衛星オペレータ(Synspective・QPS・アクセル・Marble Visions等)→軌道上サービス(アストロスケール・Pale Blue等)→回収(Elevation Space)→月探査(ispace・KDDI・立命館・東大)→データ解析(Tellus・PFN・スペースデータ等)→産業応用(住友林業・ウミトロン・東京海上等)という、上流から下流までを日本企業・大学で網羅できる体制が見え始めている。"),
      h2("国際連携の地理的拡大"),
      p("一方で、日本連合は閉鎖系として完結するのではなく、Space Compass×ESA/Hellas Sat、アストロスケール×インド/フランス、アークエッジ×豪州、Star Signal×豪州、アクセルスペース×Geoimage、Synspective×Spectee(防災)など、欧州・豪州・インド・東南アジアとの国際MoU網を急速に展開している点が特徴的だ。米国一辺倒ではなく、Quad諸国(日米豪印)とESA諸国を主要パートナーとする多角的国際ネットワークが形成されつつある。"),
      h2("月-LEO-地上の三層構造"),
      p("加えて注目すべきは、月面産業(ispace・立命館・東大・コマツ・KDDI・JAEA・トヨタ)、LEOプラットフォーム(日本低軌道社中・Elevation Space・Space BD・各大学拠点)、地上産業応用(住友林業・Green Carbon・Archeda・東京海上等)という三層構造が同時に立ち上がっていることだ。Elevation Space×日本低軌道社中、Elevation Space×ispaceの2件のMoUは、まさにこの3層を縦に貫く軌道輸送網の中核として機能する。"),
      h2("今後の展望"),
      p("2026-2030年の間に、本レポートで示した15シナリオの多くが現実のMoU・コンソーシアム・共同事業として具体化すると見込まれる。特に有望なのは、(a)「日本版シスルナ・ロジスティクス」(Elevation Space+ispace+低軌道社中+アストロスケール)、(b)「日本版Earth Observation as a Service」(Synspective+QPS+アクセル+Marble Visions)、(c)「日本打上げトライアングル」(ISC+IST+スペースワン+日本郵船+SPACE COTAN)、(d)「LEO Lab as a Service」(低軌道社中+Elevation+ispace+Space BD+各大学)の4シナリオである。これらが成立すれば、日本は宇宙産業において欧米・中国・インドに次ぐ独自の地位を確立できるだろう。"),
      p("宇宙戦略基金の本来の意義は、個別テーマの技術成果ではなく、これら技術成果を組み合わせた「日本ならではの事業」を生み出すことにある。Elevation Space×日本低軌道社中のMoUは、その構造的可能性を示した先駆的事例であり、本レポートで描いた将来協業シナリオの多くが現実のものとなることを期待したい。"),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 出典 =====
      h1("付録: 主要出典URL一覧"),
      ...cases.map(c => new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: `${c.date}  `, font: F, size: 18 }),
          new TextRun({ text: `${c.parties}: `, font: F, size: 18, bold: true }),
          link(c.url, c.url),
        ],
      })),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/sessions/quirky-great-turing/mnt/outputs/space_fund/宇宙戦略基金_採択機関_協業マッピング.docx", buf);
  console.log("Saved Word document");
});
