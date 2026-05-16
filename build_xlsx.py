"""宇宙戦略基金 採択機関 既存協業＆将来協業可能性 Excel生成スクリプト"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

# ===== 共通スタイル =====
HEADER_FILL = PatternFill('solid', start_color='1F3864')
HEADER_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=11)
TITLE_FONT = Font(name='Arial', bold=True, size=14, color='1F3864')
NORMAL_FONT = Font(name='Arial', size=10)
WRAP_ALIGN = Alignment(wrap_text=True, vertical='top', horizontal='left')
CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
THIN_BORDER = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF'),
)

# テーマ別の色
CATEGORY_COLORS = {
    '宇宙輸送': 'FCE4D6',
    '衛星等': 'DDEBF7',
    '探査等': 'E2EFDA',
    '分野共通': 'FFF2CC',
    '複合': 'EDEDED',
}

# =========================
# Sheet 1: サマリー（表紙）
# =========================
ws1 = wb.active
ws1.title = 'サマリー'
ws1['A1'] = '宇宙戦略基金 採択機関同士のMoU・協業マッピングと将来協業予想'
ws1['A1'].font = Font(name='Arial', bold=True, size=16, color='1F3864')
ws1.merge_cells('A1:E1')
ws1.row_dimensions[1].height = 28

ws1['A3'] = '作成日'
ws1['B3'] = '2026年5月16日'
ws1['A4'] = '起点'
ws1['B4'] = 'ElevationSpace × 日本低軌道社中 MoU (2026年5月11日発表)'
ws1['A5'] = '対象'
ws1['B5'] = '宇宙戦略基金 第一期・第二期 採択機関 全126機関'
ws1['A6'] = 'シート構成'

ws1['A8'] = 'シート名'
ws1['B8'] = '内容'
ws1['A8'].font = HEADER_FONT
ws1['B8'].font = HEADER_FONT
ws1['A8'].fill = HEADER_FILL
ws1['B8'].fill = HEADER_FILL
ws1['A8'].alignment = CENTER_ALIGN
ws1['B8'].alignment = CENTER_ALIGN

sheets_info = [
    ('1.サマリー', '本シート。全体構成と参照方法。'),
    ('2.既存MoU・協業一覧', '採択機関同士の確認済みMoU・協業・連携・出資事例の網羅リスト（出典URL付）'),
    ('3.採択機関一覧', '宇宙戦略基金 第一期・第二期 採択機関の整理（分野別）'),
    ('4.将来協業候補', '4観点（技術補完性／バリューチェーン／海外展開／事業化）で予想される協業候補と未来像'),
    ('5.連携ネットワーク', '頻出ハブ機関（Elevation Space、アストロスケール、三菱電機等）と接続機関の整理'),
]

for i, (n, d) in enumerate(sheets_info, start=9):
    ws1.cell(row=i, column=1, value=n).font = NORMAL_FONT
    ws1.cell(row=i, column=2, value=d).font = NORMAL_FONT
    ws1.cell(row=i, column=2).alignment = WRAP_ALIGN

ws1.column_dimensions['A'].width = 24
ws1.column_dimensions['B'].width = 80
for col_a, col_b in [('A','B')]:
    pass

# 注記
note_row = 9 + len(sheets_info) + 2
ws1.cell(row=note_row, column=1, value='注記').font = Font(name='Arial', bold=True, size=11)
notes = [
    '・本表は公開情報（プレスリリース、ニュース記事、企業公式サイト）に基づき作成しています。',
    '・「採択機関同士」の協業を中心に整理していますが、採択機関 × 重要パートナー（非採択企業）の事例も一部含めています。',
    '・「将来協業候補」は技術テーマ・事業領域の親和性から推定したものであり、各機関の公式見解ではありません。',
    '・出典URLは作成時点（2026年5月）で確認できるものを記載しています。',
]
for i, n in enumerate(notes):
    ws1.cell(row=note_row+1+i, column=1, value=n).font = NORMAL_FONT
    ws1.merge_cells(start_row=note_row+1+i, start_column=1, end_row=note_row+1+i, end_column=5)

# =========================
# Sheet 2: 既存MoU・協業一覧
# =========================
ws2 = wb.create_sheet('既存MoU・協業一覧')

# 確認できた事例（80件超のうち、採択機関2者以上が関与する/重要なもの厳選40件）
collaborations = [
    # No, 発表年月, 関与機関, テーマ分野, 連携形態, 内容要約, 出典URL
    (1, '2026-05', 'ElevationSpace × 日本低軌道社中', '探査等', 'MoU(テーマ間連携)',
     'ELS-RS(回収)とHTV-XC/Japan Module(輸送・補給)を統合し、ポストISS時代の低軌道インフラを構築。打上げ～補給～実験～回収を一気通貫で提供。',
     'https://elevation-space.com/posts/news_20260511'),
    (2, '2025-09', 'ElevationSpace × ispace', '探査等', 'MoU(月面サンプルリターン)',
     '民間日本初の月面サンプルリターン実現に向けたMoU。ispaceの軌道間輸送機(OTV)とElevationSpaceの再突入回収カプセルを組み合わせ、共同で技術実証・顧客開拓。',
     'https://prtimes.jp/main/html/rd/p/000000067.000074085.html'),
    (3, '2026-01', '三菱重工 × 三菱電機 × 三井物産 × 日本低軌道社中', '探査等', '出資・連携体制',
     '三井物産100%子会社の日本低軌道社中に三菱重工・三菱電機が第三者割当出資。HTV-XCインテグレーション、与圧モジュール、ランデブー・ドッキングを結集しポストISS開発を加速。',
     'https://www.mitsubishielectric.co.jp/ja/pr/2026/pdf/0126.pdf'),
    (4, '2025-04/07', '日本郵船 × 三菱重工業', '宇宙輸送', '共同研究開発',
     '海運初の宇宙戦略基金採択。再使用型ロケットの洋上回収システムを共同研究開発。回収船+司令船の2隻構成、2025年7月にAiP取得、2028年度実証予定。',
     'https://www.nyk.com/news/2025/20250724_02.html'),
    (5, '2025-03', '三菱重工 × UACJ × 富山住友電工', '宇宙輸送', 'JAXA共同採択',
     'ロケット向け高強度アルミ合金(Sc入りアルミ合金ワイヤー)のWAAM研究開発。日本初のアルミ-スカンジウム合金のロケット実用化を目指す。',
     'https://www.uacj.co.jp/release/20250303.html'),
    (6, '2024-10', '三菱重工 × 清水建設 × 大陽日酸', '宇宙輸送', '宇宙戦略基金内連携',
     '大型極低温推進薬タンクのWAAM/金属3D積層技術開発で連携。三菱重工・清水建設は同一テーマ内で並走しつつ、製造プロセス開発を進める。',
     'https://www.shimz.co.jp/company/about/news-release/2025/2024060.html'),
    (7, '2025-04', '丸八 × 東京大学 × 金沢工業大学', '宇宙輸送', '産学共同開発',
     '熱可塑性CFRP製極低温ロケット燃料タンクを世界初共同開発。直径5.2m級ライナーレスタンクの実証を目指す。',
     'https://www.kanazawa-it.ac.jp/kitnews/2025/0726_icc.html'),
    (8, '2025-06', '将来宇宙輸送システム(ISC) × SPACE COTAN', '宇宙輸送', 'MoU(射場利用)',
     'HOSPO(北海道スペースポート)の射場開発・利用に関するMoU締結。ASCA1.2試験機の打上げ運用計画を共同検討。',
     'https://innovative-space-carrier.co.jp/news/20250616_1'),
    (9, '2025-12', 'SPACE COTAN × 三井物産', '宇宙輸送', 'MoU',
     'HOSPOを核とした「宇宙版シリコンバレー構想」実現に向けたMoU。北海道の宇宙関連産業集積・観光促進を共同推進。',
     'https://prtimes.jp/main/html/rd/p/000000150.000078016.html'),
    (10, '2025-05', 'スペースワン × Space BD', '宇宙輸送', '共同受注(防衛省)',
     '防衛省「多軌道観測実証衛星」打上げ輸送サービスを共同受注。Space BDが業務受注し、スペースワンのカイロスロケットで2026年度打上げ予定。',
     'https://www.space-one.co.jp/news/news_20250528.html'),
    (11, '2025-01', 'インターステラテクノロジズ × ウーブン・バイ・トヨタ', '宇宙輸送', '資本業務提携',
     '約70億円の資本業務提携。トヨタの量産技術をロケット「ZERO」量産化に活用。低コスト・量産可能なロケット製造体制構築を加速。',
     'https://toyotatimes.jp/newscast/139.html'),
    (12, '2025-12', 'インターステラテクノロジズ × 東京科学大・岩手大ほか', '衛星等', '宇宙戦略基金内連携',
     '「高精度衛星編隊飛行技術」採択。アレーアンテナ地上原理実験成功。連携機関：東京科学大、奈良先端大、大阪大、湘南工科大、会津大。',
     'https://prtimes.jp/main/html/rd/p/000000073.000043667.html'),
    (13, '2025-02', 'アークエッジ・スペース × スカパーJSAT', '衛星等', '業務提携',
     '超小型衛星コンステレーション商用化加速のための業務提携。衛星管制・地上局相互利用・事業開発で協業。',
     'https://prtimes.jp/main/html/rd/p/000000045.000073065.html'),
    (14, '2025-03', 'アークエッジ・スペース × 三菱UFJ × 清水建設', '衛星等', 'MoU',
     'ハイパースペクトル技術によるリモートセンシング事業協業MoU。GHGモニタリング実証等を推進。',
     'https://prtimes.jp/main/html/rd/p/000000048.000073065.html'),
    (15, '2025-10', 'アークエッジ × ソフトバンク × NICT × 清原光学', '衛星等', '4社連携協定',
     '宇宙-成層圏・宇宙-地上間の光無線通信実証。2026年に実証衛星打上げ、2027年HAPS-LEO双方向10Gbps光通信を目指す。',
     'https://space-connect.jp/ntn-arkedge/'),
    (16, '2024-12', '三菱電機 × Synspective', '衛星等', '戦略的パートナーシップ',
     '60億円出資・覚書締結。三菱電機が筆頭株主となり、Synspectiveの小型SAR画像を安全保障用途に共同販売。',
     'https://www.mitsubishielectric.co.jp/ja/pr/2024/1219/'),
    (17, '2025-08', '三菱電機 × Pale Blue', '衛星等', '出資',
     'ME Innovation FundがPale BlueシリーズCに出資。水推進剤小型衛星推進系の技術協業を推進。',
     'https://www.mitsubishielectric.co.jp/ja/pr/2025/0807/'),
    (18, '2025-12', 'アクセルスペース × Pale Blue', '衛星等', 'MoU・軌道上実証契約',
     'ホールスラスタ「PBH-100」の2027年軌道上実証契約。AxelLiner Laboratory利用、両社で協業深化のMoUも併せて締結。',
     'https://www.axelspace.com/ja/news/paleblue/'),
    (19, '2025-12', '三菱電機 × スカパーJSAT × アクセルスペース × Synspective × QPS研究所', '衛星等', '防衛省コンステ事業共同受注',
     '防衛省「衛星コンステレーション整備・運営等事業」共同受注。5年2,831億円。SAR=Synspective+QPS、光学=アクセルスペース担当。',
     'https://www.mod.go.jp/j/press/news/2025/12/24a.html'),
    (20, '2025-02', 'Marble Visions × NTTデータ × PASCO × キヤノン電子', '衛星等', '資本業務提携',
     '高分解能・高頻度光学衛星8機コンステレーション開発合意。PASCO運用、キヤノン電子製造。2027年初号機/2028年8機体制。',
     'https://www.nttdata.com/global/ja/news/release/2025/022500/'),
    (21, '2025-03', 'Space Compass × ESA(欧州宇宙機関)', '衛星等', 'MoU',
     '衛星間光通信ネットワークの軌道上共同実証における相互運用性検討覚書。',
     'https://space-compass.com/news/000075.html'),
    (22, '2025-12', 'Space Compass × Hellas Sat', '衛星等', 'MoU',
     '衛星間光通信ネットワーク相互接続のMoU。国境を超える「宇宙統合コンピューティング・ネットワーク」構築を加速。',
     'https://space-compass.com/news/000081.html'),
    (23, '2025', 'Space Compass × QPS研究所', '衛星等', '活用検討',
     'QPSの小型SAR衛星に対しSpace Compassの光データリレーサービス活用を本格検討。防衛省案件にも両社関与。',
     'https://space-compass.com/news/000042.html'),
    (24, '2025-05', 'アストロスケール × Honda', '衛星等', '共同開発',
     '軌道上衛星給油口接続システム共同開発。HondaのロボティクスとアストロスケールのRPOD技術を融合、2029年に低軌道燃料補給技術実証。',
     'https://global.honda/jp/topics/2025/c_2025-05-30.html'),
    (25, '2026-02', 'アストロスケール × 三井住友ファイナンス&リース', '衛星等', 'MoU',
     '軌道上サービスを利用した衛星オペレーティング・リース事業など、衛星二次利用マーケット創出のMoU。',
     'https://prtimes.jp/main/html/rd/p/000000169.000084204.html'),
    (26, '2025-03', 'アストロスケール × インド3社(Digantara/Bellatrix/MEMCO)', '衛星等', 'MoU',
     'SSA・軌道上サービス分野でMoUおよび共同事業契約。インド市場・第三国市場への協力体制を構築。',
     'https://www.excite.co.jp/news/article/Prtimes_2025-03-21-67481-85/'),
    (27, '2025', 'アストロスケール × 三井物産', '探査等', '事業化検討参画',
     '三井物産パートナー企業として「日本モジュール」事業化検討に参画。軌道上点検・修理・燃料補給サービスを検討。',
     'https://prtimes.jp/main/html/rd/p/000000056.000067481.html'),
    (28, '2025-11', '立命館 × コマツ × ispace', '探査等', '宇宙戦略基金共同採択',
     '「月面拠点建設のための測量・地盤調査技術」採択。立命館代表、コマツ連携、ispace協力。地形データ取得・レゴリス調査・道路建設を体系化。',
     'https://www.komatsu.jp/ja/newsroom/2025/20251121'),
    (29, '2026-02', '東京大学 × 京都大学 × 立命館 × トプコンほか', '探査等', '宇宙戦略基金共同採択',
     '「水・金属元素探査装置のフライトモデル開発」(LUNAR-RABBIT)採択。月面資源量の実測を目指す。',
     'https://www.s.u-tokyo.ac.jp/ja/info/11065/'),
    (30, '2024-11/2025-02', 'KDDI × 福井工業大学 × ispace', '探査等', '宇宙戦略基金内連携',
     '「月-地球間通信システム開発・実証FS」採択。福井工大あわら13.5m地上局活用、ispaceは関連調査を受託。',
     'https://newsroom.kddi.com/news/detail/kddi_nr-415_3688.html'),
    (31, '2026-04', '横浜国立大 × 慶應 × 東京大学', '衛星等', '宇宙戦略基金共同採択',
     '「空間自在移動の実現に向けた技術(C)宇宙ロジスティクス」採択。軌道力学・宇宙機システム工学・物流需要予測を統合。',
     'https://www.t.u-tokyo.ac.jp/topics/tp2026-04-14-001'),
    (32, '2025', 'NeSTRA × ElevationSpace × 藤倉航装 × 金沢工業大', '探査等', '共同実施',
     '「展開型エアロシェル技術の地球大気圏突入実証と火星着陸機への適用」受託。Mars Touch Project。',
     'https://prtimes.jp/main/html/rd/p/000000063.000074085.html'),
    (33, '2025-02', 'Space Tech Accelerator × 衛星データサービス企画 × 三菱電機 × 大日本印刷', '衛星等', '宇宙戦略基金共同採択',
     '「衛星データ利用システム海外実証FS」採択。インドネシアのパームヤシ農家支援アプリ共同開発。',
     'https://www.jaxa.jp/press/2025/02/20250207-1_j.html'),
    (34, '2026-03', 'RESTEC × アクセルスペース × Synspective × PASCO × NSI', '衛星等', '宇宙戦略基金共同採択',
     '「衛星データ利用システム実装加速化事業」コンソーシアム採択。SAR・光学センサの評価・校正・検証・補正手法の環境整備。',
     'https://www.axelspace.com/news/spacestrategyfund_satellitedata/'),
    (35, '2026-03', 'アクセルスペース × 明星電気 × ANAHD × JIJ', '衛星等', '宇宙戦略基金共同採択',
     '「次世代地球観測衛星に向けた観測機能高度化技術」コンソーシアム採択。衛星編隊・旅客機観測でCO2モニタリング。',
     'https://prtimes.jp/main/html/rd/p/000000058.000066150.html'),
    (36, '2025-01', 'スペースデータ × IHI', '衛星等', 'MoU',
     '協業MoU締結。IHIの小型衛星コンステレーションデータとスペースデータのデジタルツインを融合し新事業創出。',
     'https://www.nikkei.com/article/DGXZQOUC302R40Q5A131C2000000/'),
    (37, '2026-02', '早稲田 × 慶應 × 東京理科大 × JAMSS × パナソニック × ジャムコほか', '分野共通', 'SX-CRANE共同採択',
     '「一般民間人の健康・快適宇宙空間を実現する宇宙QOL向上」拠点採択。私大唯一の代表機関。',
     'https://www.keio.ac.jp/ja/press-releases/2026/3/4/28-172973/'),
    (38, '2025', '立命館 × 東京大学 × 島津製作所 × 会津大ほか', '分野共通', 'SX研究開発拠点',
     '「月面探査・利用を産業化するための宇宙機器開発・人材育成拠点」採択。島津は水分子識別装置を担当。',
     'https://www.ritsumei.ac.jp/news/detail/?id=4081'),
    (39, '2025-06', '名古屋大 × 慶應 × 東大 × 横国 × ネッツほか', '分野共通', 'SX研究開発拠点',
     '「デトネーションエンジン・宇宙推進工学革新研究拠点」キックオフ。観測ロケットS-520-31で実証済。',
     'https://www.nagoya-u.ac.jp/info/press/sx.html'),
    (40, '2026-03/04', '大熊ダイヤモンドデバイス × Space BD', '分野共通', 'SX-ARK採択+協力',
     '「ダイヤモンド半導体による小型SARの熱制約打破」採択(SX-ARK)。Space BDが協力機関として参画。',
     'https://sorae.info/biz/20260410-spacebd-sx-ark.html'),
    (41, '2025-09', 'エネコートテクノロジーズ × トヨタ × 日揮 × KDDI × 京都大学', '分野共通', '産学連合・SX-ARK',
     'ペロブスカイト太陽電池の量産・実装に向けた産学連合(100億円規模)。SX-ARK(熱とデバイス)にも採択され宇宙用も開発。',
     'https://enecoat.com/news/20250910/'),
    (42, '2024-03/2025', '高砂熱学工業 × ispace', '探査等', 'MoU+月面実証',
     '月面用水電解装置をHAKUTO-R Mission2に搭載。サーマルマイニング技術の月面実証MoU。両機関とも東大宇宙資源拠点メンバー。',
     'https://prtimes.jp/main/html/rd/p/000000077.000140640.html'),
    (43, '2025-08', '山形大学 × 不二製油', '分野共通', '実証(SX-CRANE関連)',
     '3Dフードプリンタ麺×MIRACORE未来スープ「Beyond Ramen FURUKAWA」共同実証。両機関ともSX-CRANE山形拠点メンバー。',
     'https://swel.jp/2025/08/13/'),
    (44, '2025-09', 'ElevationSpace × Exobiosphere', '探査等', 'MoU',
     'ElevationSpaceとルクセンブルクのExobiosphereによる高頻度宇宙サンプルリターンMoU。創薬・バイオ実証で連携。',
     'https://www.exobiosphere.com/news/elevationspace-and-exobiosphere-sign-mou-to-enable-high-frequency-sample-return-from-space'),
    (45, '2025-12', 'ElevationSpace × JAXA', '探査等', 'パートナースタートアップ認定',
     'JAXAパートナースタートアップに認定。大気圏再突入・回収技術の知見支援を受け開発を加速。',
     'https://prtimes.jp/main/html/rd/p/000000071.000074085.html'),
    (46, '2025-09', 'ispace × 三井住友銀行', '探査等', 'スポンサー＋協調融資',
     'SMBCがHAKUTO-R初のオフィシャルパートナーに新規参画。2025年5月に協調融資100億円実行。シスルナ経済圏ビジョン共有。',
     'https://prtimes.jp/main/html/rd/p/000000029.000140640.html'),
    (47, '2025-07', 'NICT × 三菱ケミカル × TECHLAB × シャープ', '衛星等', '共同開発合意',
     'モビリティ向け超小型軽量衛星通信ユーザー端末の共同開発合意。',
     'https://corporate.jp.sharp/news/250730-a.html'),
    (48, '2025-10', 'NICT × スカパーJSAT × 東京大学 × 日本無線', '衛星等', '実証実験',
     'GEO/LEO衛星線と地上線を含む衛星×5Gネットワーク統合運用実証実験成功。',
     'https://www.nict.go.jp/press/2025/10/02-1.html'),
    (49, '2025-05', 'Star Signal Solutions × Inovor Technologies(豪)', '衛星等', 'MoU',
     'SSA・宇宙交通協調・衛星通信分野での豪日協力強化MoU。',
     'https://www.inovor.com.au/inovor-technologies-and-star-signal-solutions-sign-strategic-mou-to-strengthen-australia-japan-space-industry-collaborations/'),
    (50, '2025-07', 'Synspective × Spectee', '衛星等', '協業',
     'SARによる広域浸水データとSpecteeのSNS解析を統合し災害対応高度化。',
     'https://synspective.com/press-release/2025/spectee_partnership/'),
    (51, '2025-10', 'アクセルスペース × Geoimage(豪)', '衛星等', 'パートナーシップ',
     'オーストラリア市場拡大に向けた販売パートナーシップ契約。GRUSデータの海外販売チャネル拡張。',
     'https://prtimes.jp/main/html/rd/p/000000042.000066150.html'),
    (52, '2025-04', 'インターステラテクノロジズ × NICT', '衛星等', '共同研究契約',
     '超々小型衛星を用いた通信システムに関する共同研究契約。',
     'https://www.istellartech.com/news/press/8321'),
    (53, '2025-01', 'スペースワン × JR西日本イノベーションズ', '宇宙輸送', '資本業務提携',
     'シリーズDラウンドで資本業務提携。JR西日本グループの事業基盤を活用し打上げの地域価値創出を推進。',
     'https://www.westjr.co.jp/press/article/items/250116_00_press_spaceone_1.pdf'),
    (54, '2024-09', '将来宇宙輸送システム × 荏原製作所', '宇宙輸送', '包括連携',
     '電動ポンプ活用ロケットエンジン共同開発の包括連携協定。2026年3月着火試験成功。',
     'https://innovative-space-carrier.co.jp/news/20240927'),
    (55, '2025-03', '将来宇宙輸送システム × 旭化成', '宇宙輸送', '包括連携',
     '宇宙輸送分野での包括連携協定。旭化成の推進システム製造技術・評価施設(滋賀)でロケットエンジン試験。',
     'https://www.asahi-kasei.com/jp/news/2024/ze250311.html'),
    (56, '2026-01', 'Pale Blue × 東京大学', '衛星等', '宇宙戦略基金内連携',
     '「空間自在移動／軌道間輸送機」採択。Pale Blue代表、東京大学連携。Micro-OTVを共同開発。',
     'https://www.t.u-tokyo.ac.jp/topics/tp2026-01-30-001'),
    (57, '2026-04', 'アストロスケール × Exotrail(仏)', '衛星等', '戦略連携',
     '仏Exotrailとの軌道離脱ミッション契約。2030年までにLEOデオービット実証を目指す日仏宇宙協力。',
     'https://prtimes.jp/main/html/rd/p/000000104.000067481.html'),
    (58, '2025', 'JAEA × JAXA × 産業技術総合研究所', '探査等', '宇宙戦略基金共同実施',
     'アメリシウム241活用の原子力電池(半永久電源)を受託。2029年初頭プロトタイプ完成目標。月の夜・深宇宙探査用。',
     'https://www.jaea.go.jp/02/press2024/p25031801/'),
]

# ヘッダー
headers = ['No.', '発表年月', '関与機関', '主分野', '連携形態', '内容要約', '出典URL']
for col, h in enumerate(headers, start=1):
    c = ws2.cell(row=1, column=col, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = CENTER_ALIGN
    c.border = THIN_BORDER

# データ
for row_idx, row in enumerate(collaborations, start=2):
    for col_idx, val in enumerate(row, start=1):
        c = ws2.cell(row=row_idx, column=col_idx, value=val)
        c.font = NORMAL_FONT
        c.alignment = WRAP_ALIGN
        c.border = THIN_BORDER
    # カテゴリで色付け
    field = row[3]
    color = CATEGORY_COLORS.get(field, 'FFFFFF')
    for col_idx in range(1, 8):
        ws2.cell(row=row_idx, column=col_idx).fill = PatternFill('solid', start_color=color)

# 列幅
ws2.column_dimensions['A'].width = 5
ws2.column_dimensions['B'].width = 10
ws2.column_dimensions['C'].width = 40
ws2.column_dimensions['D'].width = 9
ws2.column_dimensions['E'].width = 18
ws2.column_dimensions['F'].width = 60
ws2.column_dimensions['G'].width = 50

# 行高
for row_idx in range(2, 2 + len(collaborations)):
    ws2.row_dimensions[row_idx].height = 70

ws2.freeze_panes = 'A2'

# =========================
# Sheet 3: 採択機関一覧
# =========================
ws3 = wb.create_sheet('採択機関一覧')

import openpyxl as opx
src = opx.load_workbook('/sessions/quirky-great-turing/mnt/uploads/宇宙戦略基金_採択機関一覧.xlsx', data_only=True)
ssh = src['Sheet1']

# ヘッダー
src_headers = ['No.', '実施機関名(代表機関)', '機関種別', '宇宙参画歴', '上場区分/宇宙SU', '期', '分野', '担当省', '技術開発テーマ名', 'サブテーマ/区分', '研究代表者', '連携機関']
for col, h in enumerate(src_headers, start=1):
    c = ws3.cell(row=1, column=col, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = CENTER_ALIGN
    c.border = THIN_BORDER

# データ取得
out_row = 2
for row in ssh.iter_rows(min_row=2, values_only=True):
    no, org, kind, history, listed, period, field, ministry, theme, subtheme, task_name, task_summary, lead, partners, pdf, date = row
    if org is None:
        continue
    vals = [no, org, kind, history, listed, period, field, ministry, theme, subtheme, lead, partners]
    for col, v in enumerate(vals, start=1):
        c = ws3.cell(row=out_row, column=col, value=v)
        c.font = NORMAL_FONT
        c.alignment = WRAP_ALIGN
        c.border = THIN_BORDER
    color = CATEGORY_COLORS.get(field, 'FFFFFF')
    for col in range(1, 13):
        ws3.cell(row=out_row, column=col).fill = PatternFill('solid', start_color=color)
    out_row += 1

# 列幅
widths = [5, 30, 12, 10, 12, 8, 10, 14, 40, 25, 14, 30]
for i, w in enumerate(widths, start=1):
    ws3.column_dimensions[get_column_letter(i)].width = w

ws3.freeze_panes = 'A2'

# =========================
# Sheet 4: 将来協業候補
# =========================
ws4 = wb.create_sheet('将来協業候補')

future_collabs = [
    # No, 候補機関A, 候補機関B(等), 技術補完性, バリューチェーン, 海外展開, 事業化, 想定される未来像
    (1, 'ElevationSpace', 'ispace + 日本低軌道社中 + アストロスケール',
     '回収(ElevationSpace)・OTV/月着陸(ispace)・LEOステーション(低軌道社中)・燃料補給(アストロスケール)で月-LEO-地上を貫く軌道輸送パッケージが完成',
     '上流(ロケット)〜中流(軌道機動)〜下流(回収・データ)まで完全な4社統合バリューチェーン',
     '月面サンプルリターンの欧米需要を取り込み、ESA/NASA/CSA向けに統合サービスを提案可能',
     '2030年前後の商用月面サンプルリターン市場・LEO実験回収市場で年間複数ミッションのオペレーターに',
     '日本版「シスルナ・ロジスティクス」プラットフォーム。打上げ・補給・実験・回収を一気通貫提供する世界唯一のサービス。'),
    (2, 'Synspective + QPS研究所', 'アクセルスペース + Marble Visions',
     'SAR(全天候)×光学(高分解能)×ハイパースペクトル(分光)の地球観測スペクトルが一気に揃う',
     '衛星製造(各社)→運用→データ販売(RESTEC/PASCO)→ソリューション(Tellus/Spacedata/Preferred Networks)',
     '防衛省コンステ事業を国内基盤に、ASEAN・中東・アフリカへ「日本版センチネル」を提案',
     '安全保障+商用の二重需要で年商数百億円規模のセグメント形成',
     '日本版「Earth Observation as a Service」が確立。米Maxar/Planetに対抗する選択肢を世界へ提供。'),
    (3, '将来宇宙輸送システム(ISC) + インターステラ', 'スペースワン + 日本郵船 + SPACE COTAN',
     'ベンチャー再使用機(ISC)・量産小型機(IST)・カイロス(スペースワン)・洋上回収(NYK)・射場(SPACE COTAN)が国内打上げポートフォリオを形成',
     '射場(SPACE COTAN/紀伊)→機体(各社)→部品(IHI/UACJ/丸八)→回収(NYK)→運用(各社)',
     '東南アジア・中東衛星オペレータ向け「日本トリオ打上げ」を提案',
     '年間20本以上の国内打上げ市場が成立。日本の打上げ自立性が確保される',
     '日本独自の打上げ価格競争力で「アジアの打上げハブ」に。1機体に依存しない冗長性を持つ国家インフラ。'),
    (4, 'Space Compass', 'NICT + アークエッジ・スペース + ワープスペース + Synspective + QPS研究所',
     '光通信地上系(NICT)・光端末(ワープスペース)・低軌道機(アークエッジ等)・データソース(SAR)を統合',
     '研究(NICT)→端末(ワープ)→衛星(各社)→中継(Space Compass)→顧客(防衛・商用)',
     'ESA・Hellas Sat・Apolinkとのグローバル光通信メッシュにつなぎ込む国際ハブ',
     '2030年に光通信トラフィック契約で年商数百億円。データ送信レート10Gbps級が普及',
     '日本版「宇宙コンピューティング・バックボーン」が完成。RF依存から光通信時代へ移行する世界の中核。'),
    (5, 'Pale Blue + アストロスケール + Astroscale', 'IHI + Star Signal Solutions',
     '水推進(Pale Blue)・燃料補給(アストロスケール)・SSA(IHI/Star Signal)で軌道機動 ＆ 監視のフルパッケージ',
     'OTV/燃料補給/デブリ除去/監視を統合した循環型宇宙経済の基盤',
     '英Astroscale本社経由でESA・UK Space Agency案件にアクセス、米FAAサービス需要も対応',
     '2030年代の「軌道上サービス市場」(数千億円)で先行者利益',
     '世界初の「Refuel as a Service + Deorbit as a Service」を提供する日本連合。サステナブル軌道経済の標準化。'),
    (6, '東京大学(資源)', '高砂熱学 + 栗田工業 + 横河電機 + ispace',
     '月水資源探査(東大/ispace)→電解水素生成(高砂)→水処理(栗田)→計装(横河)で月面ISRU装置の縦串',
     '月での水資源「発見→採取→精製→利用」の月面産業バリューチェーン',
     'NASA Artemis計画/ESA Argonautに日本版ISRU装置を提案',
     '2030年代の月面ISRU市場で装置・サービス事業の収益化',
     '月の水を燃料・生命維持・電力に変換する「月の水道事業」を日本連合が確立。Artemisに不可欠な装置サプライヤーに。'),
    (7, 'KDDI + 福井工業大学', 'NICT + ispace + Space Compass',
     '月-地球間通信(KDDI/福井工大)と光リレー(NICT/Space Compass)、月面ノード(ispace)で完全な月通信網',
     '地上局(KDDI)→中継衛星(Space Compass)→月軌道(ispace)→月面(ispace)の通信スタック',
     'NASA LunaNet/ESA Moonlight計画とのインターオペラビリティ確保',
     '2030年代の月探査支援通信市場(年数百億円)',
     '日本版「LunaNet」相当の月通信網を国内連合で構築。Artemis計画における日本の独自性を確保。'),
    (8, '岩谷技研 + 将来宇宙輸送システム', 'JAMSS(有人宇宙) + ジャムコ + 早稲田大学',
     '気球(岩谷)・サブオービタル(ISC)・与圧モジュール(JAMSS/ジャムコ)・QOL設計(早稲田)で有人宇宙旅行の体験設計が成立',
     '機体(岩谷/ISC)→キャビン(JAMSS/ジャムコ)→運用支援(早稲田研)→顧客提供',
     'スイス/UAE/インド富裕層向け「日本式ホスピタリティ宇宙旅行」を販売',
     '2030年代に年間数百回の有人準軌道飛行で観光・実験収益を確立',
     '日本版「Virgin Galactic + Blue Origin」とも言える有人準軌道産業。日本の宇宙ホスピタリティが世界ブランドに。'),
    (9, 'アクセルスペース + Synspective + QPS', 'ウミトロン + オーシャンソリューションテクノロジー + パシフィックコンサルタンツ + Solafune',
     '上流の衛星データ(SAR/光学)と下流の業務アプリ(漁業/IUU/防災/AI)が直結。同じ事業で複数衛星社のデータを横断利用',
     '衛星画像→AI解析→産業別アプリ→海外政府/企業顧客(東南アジア)',
     'インドネシア・フィリピン・ベトナム・アフリカで衛星×SaaSをパッケージ販売',
     '2028年に海外売上比率50%超、JICA/世銀資金活用',
     '日本発「衛星データ×AI×ODA」の社会実装モデル。SDGs海外実証の標準パートナーに。'),
    (10, '住友林業 + Green Carbon + Archeda', 'Synspective + アクセルスペース + 国際航業 + 東京海上レジリエンス',
     '森林・農地(住友林業/Green Carbon)・衛星解析(Synspective/アクセル/Archeda)・地理空間(国際航業)・保険(東京海上)でカーボン金融×宇宙の縦串',
     '計測(衛星)→検証(MRV)→クレジット組成→保険・金融商品化',
     'COP/世銀向けにアジア森林・農地クレジットの認証システムを提案',
     '2028年に年間100万トン規模の高品質クレジット発行',
     '日本発「Forest-as-an-Asset」プラットフォーム。グリーン金融×宇宙データの先進事例に。'),
    (11, 'NEC + 三菱電機 + NECスペーステクノロジー', 'IHI + IHIエアロスペース + 川崎重工',
     '衛星バス(NEC/三菱電機)・推進(IHI)・物資補給(IHIエアロ/川重)で「日本版Lockheed Martin」級の総合宇宙企業連合',
     '衛星設計→製造→打上げ→運用までの完全垂直統合',
     '中東・南米の国家プロジェクトを共同受注。三菱電機+NECで「日本宇宙総合商社」化',
     '2030年代に統合グループで年商5000億円規模',
     '個社では世界に勝てない日本の重電・電機が連合で「日本版宇宙コングロマリット」を形成。'),
    (12, 'スペースデータ + Tellus + Preferred Networks', 'アクセルスペース + Synspective + QPS + IHI',
     'PFNのAIモデル・スペースデータのデジタルツイン・Tellusのプラットフォーム・各衛星のデータが融合',
     'データ(衛星)→AI/Twin化(PFN/Spacedata)→配信(Tellus)→産業ユーザー',
     '米Palantirに対抗する「Spatial Intelligence Cloud」を北米・欧州に展開',
     '2030年に地理空間AIプラットフォームでARR数百億円',
     '日本発「地球のデジタルツイン」が世界の業務基盤に。リアルワールドAIの中核。'),
    (13, '日本低軌道社中 + ElevationSpace + ispace', 'Space BD + 慶應 + 早稲田 + 東京理科大 + 山形大学',
     'LEOステーション(低軌道社中)・回収(Elevation)・月(ispace)に大学の実験(医薬/食料/材料)をブッキング',
     '実験提案(大学)→ブッキング(Space BD)→輸送(各社)→回収(Elevation)→分析',
     'ESA/NASAだけでなくシンガポール/タイ/UAEの大学・製薬会社にも実験枠を提供',
     'LEOでの微小重力実験市場(年数百億円)で「日本枠」を獲得',
     '日本版「LEO Lab as a Service」。創薬・素材・宇宙食の研究で世界に必要不可欠なプラットフォームに。'),
    (14, '楽天モバイル + アークエッジ・スペース', 'Space Compass + ソフトバンク + NICT',
     '楽天Moblie衛星×Astera(地上)、アークエッジLEO、Space Compass光リレー、NICT周波数技術で6G/NTN実装',
     '周波数(NICT)→中継(Space Compass)→LEO(アークエッジ)→地上MNO(楽天/SBM)→顧客',
     '東南アジア・アフリカの未電化地域向けにNTN-MNOパッケージを輸出',
     '6Gコア技術として2030年代に世界標準化に貢献',
     '日本発「Non-Terrestrial Network」が6G世界標準に。通信衛星と地上MNOの統合運用で世界をリード。'),
    (15, 'エネコート + 大熊ダイヤモンドデバイス + 三洋化成', 'NECスペーステクノロジー + シャープエネルギーソリューション + アクセルスペース',
     'ペロブスカイト太陽電池×ダイヤモンド半導体×小型衛星×推進系で「衛星部品の国産化フルセット」',
     '材料(エネコート/大熊/三洋化成)→部品(NECスペース)→衛星(アクセル)→運用',
     '欧米衛星メーカーへの戦略物資としての輸出。経済安全保障案件',
     '2030年に衛星部品市場で日本シェアを5%→15%へ',
     '日本の素材・材料力が「宇宙の半導体・電池」として復活。衛星部品の輸出産業化。'),
    (16, 'トヨタ自動車 + JAEA + IHIエアロスペース', '東京大学(資源) + 立命館 + コマツ',
     '燃料電池(トヨタ)・原子力電池(JAEA)・物資補給(IHIエアロ)が月面拠点の動力源を提供。月面建設(コマツ/立命館)が消費先',
     '電源(トヨタ/JAEA)→補給(IHIエアロ)→建設(コマツ)→拠点(立命館)→利用(東大)',
     'NASA Artemis有人探査ローバ(LUNAR CRUISER)を経由した米国市場アクセス',
     '2030年代の月面拠点・有人ローバの動力・建設市場で日本連合が独占的地位',
     '月面拠点の「電気・水・道路」を日本連合が供給。Artemis VII以降の月面産業の必須プレイヤー。'),
    (17, '三菱電機 × NEC + JAXA', '岩谷技研 + ジャムコ + 早稲田 + 東京女子医科大',
     '与圧モジュール(JAMSS系)・QOL設計(早稲田)・医療(東女医大)・気球技術(岩谷)で有人宇宙居住・医療の縦串',
     '居住モジュール(JAMSS)→生命維持(JAXA/JAMSS)→医療(東女医大)→QOL(早稲田)',
     'Axiom Space/Orbital ReefなどポストISS民間ステーションへの装備供給',
     '2030年代の民間宇宙ステーション市場で日本シェアを確保',
     '宇宙居住・医療領域で「人を最も大切にする日本の宇宙建築」が世界標準に。'),
    (18, '岩谷技研 + アストロスケール + パワーレーザー', 'Star Signal Solutions + IHI',
     '気球で運ぶ精密センサ・パワーレーザー光学・地上望遠鏡(Star Signal)・SSAデータ(IHI)で「日本版SSAネットワーク」を構築',
     '観測(岩谷気球+地上局)→処理(IHI/Star Signal)→応用(アストロスケール:衝突回避)',
     '日米・日豪SSA協定への独自データ提供。NATO・QUAD SSA協力枠組み参加',
     '2030年に防衛省・米宇宙軍・米FCC向けにSSAデータを供給',
     '日本版「Space Domain Awareness」が国際SSAの一角に。商用衛星オペレータも顧客に。'),
    (19, '東レ + UACJ + 東レ・カーボンマジック + コンポジットテーラーズ', '丸八 + 三菱重工 + IHI',
     '炭素繊維(東レ/CTL/カーボンマジック)・アルミ合金(UACJ)・タンク(丸八)・機体(三菱重/IHI)で「日本素材機体」の純国産化',
     '素材→中間素材→部品→機体→打上げ→運用までの完全純国産バリューチェーン',
     '国際素材市場で「日本製ロケット素材」が標準に。欧州ロケットメーカーにも輸出',
     '2030年代に素材輸出で年商数百億円',
     '日本の素材力で「日の丸ロケット」が完成。素材→機体→打上げまで全て国産化された強靭なサプライチェーン。'),
    (20, 'クロスユー(三井不動産)', '住友林業 + アクセルスペース + Synspective + Location Mind + Solafune',
     '都市/森林データ(住友林業)・衛星(各社)・人流(Location Mind)・AI(Solafune)で「都市OS」のグローバル展開',
     '都市データ→AI解析→政策・不動産・観光ソリューション→海外都市',
     'アフリカ・中央アジアのスマートシティ案件で日本連合パッケージを提案',
     '2030年に海外スマートシティで都市OSライセンス収益',
     '三井不動産発「Cross U」が世界の都市の頭脳に。日本のスマートシティ輸出の旗艦事例。'),
]

future_headers = ['No.', '中核機関', '連携候補機関', '技術補完性', 'バリューチェーン上の連携', '海外展開・国際協力', '事業化・収益化までの道筋', '想定される未来像']
for col, h in enumerate(future_headers, start=1):
    c = ws4.cell(row=1, column=col, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = CENTER_ALIGN
    c.border = THIN_BORDER

for row_idx, row in enumerate(future_collabs, start=2):
    for col_idx, val in enumerate(row, start=1):
        c = ws4.cell(row=row_idx, column=col_idx, value=val)
        c.font = NORMAL_FONT
        c.alignment = WRAP_ALIGN
        c.border = THIN_BORDER

widths4 = [5, 22, 30, 35, 35, 30, 30, 35]
for i, w in enumerate(widths4, start=1):
    ws4.column_dimensions[get_column_letter(i)].width = w

for row_idx in range(2, 2 + len(future_collabs)):
    ws4.row_dimensions[row_idx].height = 130

ws4.freeze_panes = 'A2'

# =========================
# Sheet 5: 連携ネットワーク(ハブ機関)
# =========================
ws5 = wb.create_sheet('連携ネットワーク')

hub_data = [
    ('ElevationSpace', '4-5件', 'ispace, 日本低軌道社中, Exobiosphere, NeSTRA, JAXA, 東北大学', '低軌道-月面の回収・帰還ハブ'),
    ('日本低軌道社中', '3-4件', '三井物産, 三菱重工, 三菱電機, ElevationSpace, アストロスケール', 'ポストISS時代のLEOステーション統合'),
    ('アストロスケール', '5件以上', 'Honda, Exotrail, インド3社, 三井物産, 三井住友FL, エアバス', '軌道上サービスの国際展開ハブ'),
    ('Space Compass', '5件以上', 'ESA, Hellas Sat, アクセルスペース, QPS研究所, JSAT International, Apolink', '光通信・データリレーの国際ハブ'),
    ('アークエッジ・スペース', '4-5件', 'スカパーJSAT, MUFG, 清水建設, ソフトバンク, NICT, セーレン', '小型衛星製造・運用エコシステム'),
    ('三菱電機', '6件以上', 'Synspective, Pale Blue, 日本低軌道社中, スカパー, アクセル, QPS, NEC', '日本最大級の宇宙コングロマリット'),
    ('ispace', '5件以上', 'ElevationSpace, 立命館, コマツ, KDDI, SMBC, 高砂熱学, 中央大学', '月面探査・産業化のハブ'),
    ('立命館大学', '3-4件', 'コマツ, ispace, 島津製作所, 東京大学, 慶應, JAXA', '月面拠点形成のハブ'),
    ('SPACE COTAN', '3件', 'ISC, 三井物産, 岩谷技研, 清水建設', '北海道宇宙クラスター(HOSPO)'),
    ('NICT', '3-4件', 'インターステラ, スカパーJSAT, シャープ, ソフトバンク, アークエッジ', '通信技術の連携ハブ'),
    ('東京大学', '5件以上', 'Pale Blue, 京都大学, 立命館, 慶應, ispace, 高砂熱学', 'アカデミア中核ハブ'),
    ('JAXA(基金事務局)', '全体', '全採択機関', '宇宙戦略基金事業全体の取りまとめ'),
]

hub_headers = ['ハブ機関', '連携先数', '主要連携先', '役割・特徴']
for col, h in enumerate(hub_headers, start=1):
    c = ws5.cell(row=1, column=col, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = CENTER_ALIGN
    c.border = THIN_BORDER

for row_idx, row in enumerate(hub_data, start=2):
    for col_idx, val in enumerate(row, start=1):
        c = ws5.cell(row=row_idx, column=col_idx, value=val)
        c.font = NORMAL_FONT
        c.alignment = WRAP_ALIGN
        c.border = THIN_BORDER

ws5.column_dimensions['A'].width = 22
ws5.column_dimensions['B'].width = 12
ws5.column_dimensions['C'].width = 60
ws5.column_dimensions['D'].width = 40

for row_idx in range(2, 2 + len(hub_data)):
    ws5.row_dimensions[row_idx].height = 40

ws5.freeze_panes = 'A2'

# 保存
out = '/sessions/quirky-great-turing/mnt/outputs/space_fund/宇宙戦略基金_採択機関_協業マッピング.xlsx'
wb.save(out)
print(f"Saved: {out}")
print(f"既存協業: {len(collaborations)}件")
print(f"将来候補: {len(future_collabs)}件")
print(f"ハブ機関: {len(hub_data)}件")
