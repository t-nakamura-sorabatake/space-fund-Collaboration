"""HTMLファイルの各ノードに採択機関(orgs)フィールドを追加するスクリプト"""
import re

# ノードID → 採択機関リスト
orgs_map = {
    # ===== 宇宙輸送 第一期 =====
    "T1": ["丸八(熱可塑CFRP)", "ニコン(精密3D積層)", "清水建設(大型3D積層)", "三菱重工(大型3D積層)"],
    "T2": ["IHIエアロスペース"],
    "T3": ["日本郵船"],
    "T4": ["SPACE COTAN"],
    "T5": ["三菱プレシジョン"],
    # ===== 宇宙輸送 第二期 =====
    "T6": ["岩谷技研", "宇宙システム開発"],
    "T7": ["将来宇宙輸送システム", "三菱重工"],
    "T8": ["IHIエアロスペース"],
    "T9": ["イーグル工業", "エア・ウォーター", "NECスペーステクノロジー", "シンフォニアテクノロジー", "SUIHO SPACE INNOVATIONS", "Space BD(連携:川崎重工)", "MJOLNIR SPACEWORKS"],
    "T10": ["IHI", "赤星工業", "スペースワン", "東レ・カーボンマジック", "徳田工業", "光製作所", "富士精工", "北斗", "UACJ"],
    "T11": ["(採択者なし)"],
    # ===== 衛星等 第一期 =====
    "S1": ["情報通信研究機構(NICT)"],
    "S2": ["日本電気(NEC)"],
    "S3": ["アークエッジ・スペース", "QPS研究所", "Synspective", "日本電気(NEC)"],
    "S4": ["Marble Visions"],
    "S5": ["京都大学"],
    "S6": ["インターステラテクノロジズ", "名古屋大学", "東京大学"],
    "S7": ["ウェルリサーチ", "NECスペーステクノロジー", "NU-Rei", "コンポジットテーラーズ", "シャープエネルギーソリューション", "GSユアサテクノロジー", "三菱電機(補助+委託)", "INDUSTRIAL-X", "衛星システム技術推進機構(JSPI)", "出光興産(追加公募)", "太陽金網(追加公募)", "ナノブリッジ・セミコンダクター(追加公募)"],
    "S8": ["ウミトロン", "オーシャンアイズ", "オーシャンソリューションテクノロジー", "Space Tech Accelerator", "Solafune", "パシフィックコンサルタンツ", "日本宇宙フォーラム"],
    # ===== 衛星等 第二期 =====
    "S9": ["Space Compass"],
    "S10": ["ワープスペース"],
    "S11": ["情報通信研究機構(NICT)", "日本電気(NEC)", "三菱電機"],
    "S12": ["三菱電機"],
    "S13": ["楽天モバイル"],
    "S14": ["アクセルスペース(新市場開拓)", "Synspective(既存市場拡大)"],
    "S15": ["スペースデータ", "Tellus", "Preferred Networks"],
    "S16": ["日本電気(NEC)", "Pale Blue", "三菱電機"],
    "S17": ["アストロスケール"],
    "S18": ["横浜国立大学"],
    "S19": ["三菱電機", "東レ"],
    "S20": ["パワーレーザー"],
    "S21": ["Star Signal Solutions(連携:東大木曽観測所)", "IHI", "パワーレーザー"],
    "S22": ["Sarmony"],
    "S23": ["Archeda", "ウエスコ", "沖電気工業", "Green Carbon", "国際航業", "Space Tech Accelerator", "住友林業", "東京海上レジリエンス", "RESTEC(リモートセンシング技術センター)", "LocationMind", "衛星システム技術推進機構(JSPI)", "クロスユー", "デロイトトーマツスペースアンドセキュリティ", "日本宇宙フォーラム", "野村総合研究所", "三菱総合研究所"],
    # ===== 探査等 第一期 =====
    "E1": ["東京科学大学"],
    "E2": ["KDDI", "福井工業大学"],
    "E3": ["日本低軌道社中(連携:三菱電機)"],
    "E4": ["IHIエアロスペース"],
    "E5": ["日本低軌道社中"],
    "E6": ["Space BD"],
    "E7": ["トヨタ自動車"],
    "E8": ["日本原子力研究開発機構(JAEA)"],
    "E9": ["次世代宇宙システム技術研究組合(NeSTRA)"],
    "E10": ["アークエッジ・スペース"],
    # ===== 探査等 第二期 =====
    "E11": ["SpaceBlast"],
    "E12": ["ElevationSpace"],
    "E13": ["ispace"],
    "E14": ["東京大学", "東北大学", "立命館大学"],
    "E15": ["日本低軌道社中"],
    # ===== 分野共通 =====
    "C1": ["名古屋大学(デトネーション牽引型)", "東京大学(宇宙資源牽引型)", "東京大学(次世代太陽電池牽引型)", "立命館大学(月面産業化共用型)", "国立天文台(共用型)"],
    "C2": ["東京海洋大学(PNT)", "東京科学大学(2件:有人居住, 健康)", "山形大学(宇宙食)", "早稲田大学(宇宙QOL)"],
    "C3": ["岩手大学", "エネコートテクノロジーズ", "愛媛大学", "大熊ダイヤモンドデバイス", "大阪大学", "キオクシア", "ケミトックス", "産業技術総合研究所", "ダイキン工業", "名古屋工業大学", "名古屋大学(2件)", "北海道大学"],
    "C4": ["NECスペーステクノロジー", "九州大学", "京都大学", "神戸大学", "産業技術総合研究所", "三洋化成工業", "国立天文台", "島根大学", "SteraVision", "東京大学", "東京理科大学", "東北大学", "名古屋大学", "日本化薬", "八田・山本宇宙推進機製作所", "パッチドコニックス", "フコク", "横浜国立大学"],
    "C5": ["IMV(各種環境試験)", "高エネルギー加速器研究機構", "日本原子力研究開発機構(2件)", "理化学研究所", "量子科学技術研究開発機構(QST)", "SEESE"],
}

# HTMLファイル読み込み
path = '/sessions/quirky-great-turing/mnt/outputs/space_fund/宇宙戦略基金_技術テーマ連携ネットワーク.html'
with open(path, 'r', encoding='utf-8') as f:
    html = f.read()

# 各ノードに orgs フィールドを挿入
# パターン: { id: "T1", label: "...", cat: "...", period: N,
#     summary: "..." },
# 挿入位置: summary: の直前

def insert_orgs(match):
    node_id = match.group(1)
    rest = match.group(2)  # summary: "..." の部分
    if node_id not in orgs_map:
        return match.group(0)  # 変更なし
    orgs_js = '[' + ', '.join(f'"{o}"' for o in orgs_map[node_id]) + ']'
    return f'{{ id: "{node_id}", {match.group(3)}\n    orgs: {orgs_js},\n    {rest}'

# 各ノード行を抽出して書き換え
# パターン: { id: "X1", label: "...", cat: "...", period: N,
#     summary: "..." }
node_pattern = re.compile(
    r'\{ id: "([TSEC]\d+)", ((?:label|cat|period)[^,\n]*(?:, (?:label|cat|period)[^,\n]*)*),\n    (summary: "[^"]*")\s*\}',
    re.MULTILINE
)

count = [0]
def replace_node(m):
    node_id = m.group(1)
    head = m.group(2)
    summary = m.group(3)
    if node_id not in orgs_map:
        return m.group(0)
    orgs_js = '[' + ', '.join('"' + o.replace('"', '\\"') + '"' for o in orgs_map[node_id]) + ']'
    count[0] += 1
    return f'{{ id: "{node_id}", {head},\n    orgs: {orgs_js},\n    {summary} }}'

html_new = node_pattern.sub(replace_node, html)
print(f"Updated {count[0]} nodes")

# 詳細パネルの表示部分も更新（orgs を表示するように）
# 元の表示部分を探す
old_detail = '''      <div class="theme-meta">
        <span style="background:${colorMap[d.cat]}; color:#fff">${catName[d.cat]}</span>
        <span>第${d.period}期</span>
      </div>
      <div class="theme-summary">${d.summary}</div>'''

new_detail = '''      <div class="theme-meta">
        <span style="background:${colorMap[d.cat]}; color:#fff">${catName[d.cat]}</span>
        <span>第${d.period}期</span>
      </div>
      <div class="theme-summary">${d.summary}</div>
      ${d.orgs && d.orgs.length ? `<div class="orgs-section">
        <h2 style="margin-bottom:8px">採択機関（${d.orgs.length}件）</h2>
        ${d.orgs.map(o => `<span class="org-chip">${o}</span>`).join("")}
      </div>` : ""}'''

# 上記は実際のJS内のテンプレートリテラル。前のHTMLでは少し違うインデント
old_detail_actual = '''    <div class="theme-meta">
      <span style="background:${colorMap[d.cat]}; color:#fff">${catName[d.cat]}</span>
      <span>第${d.period}期</span>
    </div>
    <div class="theme-summary">${d.summary}</div>'''

new_detail_actual = '''    <div class="theme-meta">
      <span style="background:${colorMap[d.cat]}; color:#fff">${catName[d.cat]}</span>
      <span>第${d.period}期</span>
    </div>
    <div class="theme-summary">${d.summary}</div>
    ${d.orgs && d.orgs.length ? `<div class="orgs-section">
      <div class="orgs-title">採択機関（${d.orgs.length}件）</div>
      ${d.orgs.map(o => `<span class="org-chip">${o}</span>`).join("")}
    </div>` : ""}'''

if old_detail_actual in html_new:
    html_new = html_new.replace(old_detail_actual, new_detail_actual)
    print("Updated detail panel template")
else:
    print("WARN: detail template not found")

# CSSにorg-chipスタイルを追加
old_css_marker = '  .stats { font-size: 11px;'
new_css = '''  .orgs-section { margin-top: 12px; padding-top: 10px; border-top: 1px dashed #2a3a4f; }
  .orgs-title { font-size: 12px; color: #8fb4e0; margin-bottom: 8px; font-weight: bold; }
  .org-chip { display: inline-block; background: #2a3a4f; color: #e0e6ed; padding: 3px 8px; margin: 2px 4px 2px 0; border-radius: 3px; font-size: 11px; line-height: 1.4; }
  .stats { font-size: 11px;'''
html_new = html_new.replace(old_css_marker, new_css)
print("Added CSS for org chips")

# ツールチップにも採択機関を表示
old_tooltip = '''node.on("mouseover", (event, d) => {
  tooltip.style("display", "block")
    .html(`<b>${d.label.replace(/\\n/g, " ")}</b><br>${catName[d.cat]}・第${d.period}期`);
})'''

new_tooltip = '''node.on("mouseover", (event, d) => {
  const orgsText = d.orgs && d.orgs.length
    ? `<br><span style="color:#8fb4e0">採択機関(${d.orgs.length}):</span> ${d.orgs.slice(0, 3).join(", ")}${d.orgs.length > 3 ? `... 他${d.orgs.length - 3}件` : ""}`
    : "";
  tooltip.style("display", "block")
    .html(`<b>${d.label.replace(/\\n/g, " ")}</b><br>${catName[d.cat]}・第${d.period}期${orgsText}`);
})'''

if old_tooltip in html_new:
    html_new = html_new.replace(old_tooltip, new_tooltip)
    print("Updated tooltip template")
else:
    print("WARN: tooltip template not found in expected form")

# 保存
with open(path, 'w', encoding='utf-8') as f:
    f.write(html_new)

print(f"\nSaved updated HTML: {path}")
print(f"File size: {len(html_new)} chars")
