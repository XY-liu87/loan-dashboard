"""
看板数据生成：从 大表.xlsx 生成 data_v2.js 和中间 CSV。
用法：python 生成看板数据.py
输入：大表.xlsx（自动查找看板数据目录或桌面）
输出：data_v2.js（看板同目录）, CSV 中间文件（看板结果/）
"""
import pandas as pd
import numpy as np
import json, os, sys, csv, io

# 强制设置输出编码为UTF-8，避免中文乱码（仅控制台有效，notebook自动跳过）
if sys.platform == 'win32':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        pass

# 确定脚本所在目录（兼容 Jupyter / 命令行）
try:
    BASE = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE = os.getcwd()

# 看板数据目录：优先找脚本同级的"看板数据"，其次桌面的"看板数据"
DATA_DIR = os.path.join(BASE, '看板数据')
if not os.path.isdir(DATA_DIR):
    alt_data = os.path.join(os.path.dirname(BASE), '看板数据')
    if os.path.isdir(alt_data):
        DATA_DIR = alt_data
if not os.path.isdir(DATA_DIR):
    desktop_data = os.path.join(os.path.expanduser('~'), 'Desktop', '看板数据')
    if os.path.isdir(desktop_data):
        DATA_DIR = desktop_data

# 看板结果目录
RESULT_DIR = os.path.join(BASE, '看板结果')
if not os.path.isdir(RESULT_DIR):
    alt_result = os.path.join(os.path.dirname(BASE), '看板结果')
    if os.path.isdir(alt_result):
        RESULT_DIR = alt_result

# 大表路径：优先数据目录，其次桌面
DATA_PATH = os.path.join(DATA_DIR, '大表.xlsx')
if not os.path.exists(DATA_PATH):
    desktop_dabiao = os.path.join(os.path.expanduser('~'), 'Desktop', '大表.xlsx')
    if os.path.exists(desktop_dabiao):
        DATA_PATH = desktop_dabiao

# 自动查找 M1 数据源：优先使用 累计.xlsx，其次 icmp_*.xlsx
M1_PATH = None
# 1) 看板数据目录下的 累计.xlsx
cumulative_path = os.path.join(DATA_DIR, '累计.xlsx')
if os.path.exists(cumulative_path):
    M1_PATH = cumulative_path
# 2) 桌面上的 累计.xlsx
if not M1_PATH:
    desktop_cumulative = os.path.join(os.path.expanduser('~'), 'Desktop', '累计.xlsx')
    if os.path.exists(desktop_cumulative):
        M1_PATH = desktop_cumulative
# 3) 兼容旧的 icmp_*.xlsx
if not M1_PATH:
    for f in os.listdir(DATA_DIR):
        if f.startswith('icmp_') and f.endswith('.xlsx'):
            M1_PATH = os.path.join(DATA_DIR, f)
            break

os.makedirs(RESULT_DIR, exist_ok=True)

# 检查大表是否存在
if not os.path.exists(DATA_PATH):
    print(f'[ERROR] 未找到大表.xlsx')
    print(f'  已搜索: {DATA_PATH}')
    desktop_check = os.path.join(os.path.expanduser('~'), 'Desktop', '大表.xlsx')
    print(f'          {desktop_check}')
    print(f'  请将大表.xlsx放到桌面或看板数据目录下')
    exit(1)

# ============================================================
# 1. 读取大表
# ============================================================
df = pd.read_excel(DATA_PATH)
print(f'读取: {DATA_PATH}')
print(f'共 {len(df)} 行, 列: {list(df.columns)}')

# 提取自然回收人员数据（刘桂欣→郑州，张静→成都），在过滤前保存
natural_recovery_rows = {}
for _, r in df.iterrows():
    g = str(r.iloc[2]).strip()
    region = str(r.iloc[3]).strip() if not pd.isna(r.iloc[3]) else ''
    name = str(r.iloc[0]).strip()
    if g in ('郑州', '成都', '自然回收'):
        principal = float(r['分案剩余本金']) if not pd.isna(r['分案剩余本金']) else 0
        total = float(r['累计回退']) if not pd.isna(r['累计回退']) else 0
        today = float(r['当日累计回退']) if not pd.isna(r['当日累计回退']) else 0
        # 确定归属地区：优先用地区列，其次用组别名
        if g == '自然回收':
            target_region = region if region and region != 'nan' and region != '' else '全部'
        else:
            target_region = g
        if name not in natural_recovery_rows:
            natural_recovery_rows[name] = {'region': target_region, 'principal': 0, 'total': 0, 'today': 0}
        natural_recovery_rows[name]['principal'] += principal
        natural_recovery_rows[name]['total'] += total
        natural_recovery_rows[name]['today'] += today
print(f'  自然回收人员: {list(natural_recovery_rows.keys())}')

def get_nr_group(nr_region):
    """自然回收属于所有地区，统一归入'自然流'"""
    return '自然流'

# 计算各小组人力占比（基于人力占比.xlsx）
hr_ratio_path = os.path.join(DATA_DIR, '人力占比.xlsx')
manpower_ratio = {}  # 各地区内部占比
overall_manpower_ratio = {}  # 全量占比（用于"全部"地区自然回收分配）
region_manpower = {}  # 两地人力权重
if os.path.exists(hr_ratio_path):
    # Sheet1: 小组人力占比
    hr_df = pd.read_excel(hr_ratio_path, sheet_name='小组')
    for _, r in hr_df.iterrows():
        g = str(r['组别']).strip()
        ratio_val = float(r['人力占比']) if not pd.isna(r['人力占比']) else 0
        if g.startswith('CD') or g.startswith('ZZ'):
            manpower_ratio[g] = ratio_val
    # Sheet2: 三地权重 (ZZ在列名, CD1/CD2在数据行)
    hr_s2 = pd.read_excel(hr_ratio_path, sheet_name='地区')
    # 郑州权重 = 第二列列名, 成都一权重 = 第一行第二列的值, 成都二权重 = 第二行第二列的值
    region_manpower['ZZ'] = float(hr_s2.columns[1])
    region_manpower['CD1'] = float(hr_s2.iloc[0, 1])
    region_manpower['CD2'] = float(hr_s2.iloc[1, 1])
    # 计算全量占比 = 三地权重 x 组内占比
    for g in manpower_ratio:
        if g == 'CD-F3':
            overall_manpower_ratio[g] = region_manpower.get('CD2', 0) * manpower_ratio[g]
        elif g.startswith('CD'):
            overall_manpower_ratio[g] = region_manpower.get('CD1', 0) * manpower_ratio[g]
        elif g.startswith('ZZ'):
            overall_manpower_ratio[g] = region_manpower.get('ZZ', 0) * manpower_ratio[g]
    print(f'  三地权重(Sheet2): ZZ={region_manpower.get("ZZ",0)*100:.2f}%, CD1={region_manpower.get("CD1",0)*100:.2f}%, CD2={region_manpower.get("CD2",0)*100:.2f}%')
else:
    print(f'  [WARN] 未找到人力占比.xlsx，自然回收数据将无法分配')
# 只保留标准组别
VALID_GROUPS = ['ZZ-F1', 'ZZ-F2', 'ZZ-F3', 'ZZ-F4', 'ZZ-F5', 'CD-F1', 'CD-F2', 'CD-F3']
df = df[df.iloc[:, 2].astype(str).str.strip().isin(VALID_GROUPS)]
print(f'过滤组别后: {len(df)} 行')

# 确保必需列存在
required = ['坐席', '组别', '案件数', '分案剩余本金', '累计回退', '回退率',
            '首催分案', '首催回退', '首催回退率', '排名', '当日累计回退',
            '当日老案回退', '10点前', '10-12点', '12-14点', '14-16点',
            '16-18点', '18-20点', '20-23点']
for c in required:
    if c not in df.columns:
        print(f'[WARN] 缺少列: {c}')
        df[c] = 0

# 备注/类别
if '备注' in df.columns:
    df['类别'] = df['备注'].fillna('0')
else:
    df['类别'] = '常规'

# 地区：优先使用Excel原生地区列，nan时从组别前缀推导
def derive_region(g):
    g = str(g).strip().upper()
    if g == '自然流': return '全部'
    if g.startswith('ZZ'): return '郑州'
    if g == 'CD-F3' or g.startswith('CD-F3-'): return '成都二'
    if g.startswith('CD'): return '成都一'
    return '郑州'

def get_region(row):
    try:
        orig = str(row.iloc[3]).strip()  # Col 3 = 地区 (when called with row)
        if orig and orig != 'nan':
            # Map old "成都" → "成都一"/"成都二" based on group
            if orig == '成都':
                g = str(row.iloc[2]).strip().upper() if hasattr(row, 'iloc') else ''
                return derive_region(g)
            return orig
    except AttributeError:
        pass  # Called with string
    g = row.iloc[2] if hasattr(row, 'iloc') else row
    return derive_region(g)

df['地区'] = df.apply(get_region, axis=1)

# ============================================================
# 2. 构建 TODAY_DETAIL
# ============================================================
today_rows = []
for _, r in df.iterrows():
    total = r['当日累计回退']
    if pd.isna(total) or total == 0:
        continue
    today_rows.append({
        'name': str(r['坐席']).strip(),
        'total': float(total),
        's1': float(r.get('10点前', 0) or 0),
        's2': float(r.get('10-12点', 0) or 0),
        's3': float(r.get('12-14点', 0) or 0),
        's4': float(r.get('14-16点', 0) or 0),
        's5': float(r.get('16-18点', 0) or 0),
        's6': float(r.get('18-20点', 0) or 0),
        's7': float(r.get('20-23点', 0) or 0),
        'rank': 0,
        'region': r['地区'],
        'group': str(r['组别']).strip(),
        'category': str(r['类别']).strip() if str(r['类别']).strip() in ['大额', '常规', '小额'] else '常规',
    })

# 添加自然回收人员到 TODAY_DETAIL（自然回收属于全部地区）
for name, nr in natural_recovery_rows.items():
    if nr['today'] > 0:
        today_rows.append({
            'name': name,
            'total': nr['today'],
            's1': 0, 's2': 0, 's3': 0, 's4': 0, 's5': 0, 's6': 0, 's7': 0,
            'rank': 0,
            'region': '全部',
            'group': '自然流',
            'category': '常规',
        })

today_rows.sort(key=lambda x: -x['total'])
for i, row in enumerate(today_rows):
    row['rank'] = i + 1

print(f'  TODAY_DETAIL: {len(today_rows)} entries')

# 写 today_detail.csv
with open(os.path.join(RESULT_DIR, 'today_detail.csv'), 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['催员', '当日累计回退', '10点前回款', '10-12点', '12-14点', '14-16点',
                '16-18点', '18-20点', '20-23点', '当日回退排名', '地区', '组别', '备注'])
    for r in today_rows:
        w.writerow([r['name'], r['total'], r['s1'], r['s2'], r['s3'], r['s4'],
                    r['s5'], r['s6'], r['s7'], r['rank'], r['region'], r['group'], r['category']])

# ============================================================
# 3. 构建 CUMULATIVE_RANK
# ============================================================
rank_rows = []
for _, r in df.iterrows():
    total = r['累计回退']
    if pd.isna(total) or total == 0:
        continue
    rate_raw = str(r['回退率']).replace('%', '')
    try:
        rate = float(rate_raw) / 100
    except:
        rate = 0

    first_rate_raw = str(r['首催回退率']).replace('%', '')
    try:
        first_rate = float(first_rate_raw) / 100
    except:
        first_rate = 0

    rank_rows.append({
        'name': str(r['坐席']).strip(),
        'group': str(r['组别']).strip(),
        'cases': int(r['案件数']) if not pd.isna(r['案件数']) else 0,
        'principal': float(r['分案剩余本金']) if not pd.isna(r['分案剩余本金']) else 0,
        'total': float(total),
        'rate': rate,
        'diff': 0,
        'firstAmount': float(r['首催分案']) if not pd.isna(r['首催分案']) else 0,
        'firstBack': float(r['首催回退']) if not pd.isna(r['首催回退']) else 0,
        'firstRate': first_rate,
        'rank': int(r['排名']) if not pd.isna(r['排名']) else 0,
        'todayOld': float(r['当日老案回退']) if not pd.isna(r['当日老案回退']) else 0,
        'category': str(r['类别']).strip() if str(r['类别']).strip() in ['大额', '常规', '小额'] else '常规',
    })

rank_rows.sort(key=lambda x: -x['total'])
# 添加自然回收人员到 CUMULATIVE_RANK（自然回收属于全部地区）
for name, nr in natural_recovery_rows.items():
    if nr['total'] > 0 and nr['principal'] > 0:
        nr_rate = nr['total'] / nr['principal']
        nr_region = nr['region']
        rank_rows.append({
            'name': name,
            'group': get_nr_group(nr_region),
            'cases': 0,
            'principal': nr['principal'],
            'total': nr['total'],
            'rate': nr_rate,
            'diff': 0,
            'firstAmount': 0,
            'firstBack': 0,
            'firstRate': 0,
            'rank': 0,
            'todayOld': 0,
            'category': '常规',
        })

rank_rows.sort(key=lambda x: -x['total'])
if rank_rows:
    top_total = rank_rows[0]['total']
    for r in rank_rows:
        r['diff'] = top_total - r['total']

print(f'  CUMULATIVE_RANK: {len(rank_rows)} entries')

# 写 cumulative_rank.csv
with open(os.path.join(RESULT_DIR, 'cumulative_rank.csv'), 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['催员', '组别', '案件数', '分案剩余本金', '累计回退', '累计回退率',
                '各段位与第一名差异', '首催分案', '首催回退', '首催回退率', '回退额排名',
                '当日老案回退', '类别', '地区'])
    for r in rank_rows:
        w.writerow([r['name'], r['group'], r['cases'], r['principal'], r['total'],
                    r['rate'], r['diff'], r['firstAmount'], r['firstBack'],
                    r['firstRate'], r['rank'], r['todayOld'], r['category'], ''])

# ============================================================
# 4. 构建 GROUP_SUMMARY
# ============================================================
group_data = {}
for _, r in df.iterrows():
    g = str(r['组别']).strip()
    if not g or g == 'nan':
        continue

    if g not in group_data:
        group_data[g] = {
            'principal': 0, 'total': 0, 'staff': 0,
            'firstAssign': 0, 'firstBack': 0, 'todayOld': 0,
        }
    gd = group_data[g]
    gd['staff'] += 1
    gd['principal'] += float(r['分案剩余本金']) if not pd.isna(r['分案剩余本金']) else 0
    gd['total'] += float(r['累计回退']) if not pd.isna(r['累计回退']) else 0
    gd['firstAssign'] += float(r['首催分案']) if not pd.isna(r['首催分案']) else 0
    gd['firstBack'] += float(r['首催回退']) if not pd.isna(r['首催回退']) else 0
    gd['todayOld'] += float(r['当日老案回退']) if not pd.isna(r['当日老案回退']) else 0

group_summary = []
for gname, gd in sorted(group_data.items()):
    rate = gd['total'] / gd['principal'] if gd['principal'] > 0 else 0
    first_rate = gd['firstBack'] / gd['firstAssign'] if gd['firstAssign'] > 0 else 0
    group_summary.append({
        'section': '全量数据',
        'group': gname,
        'staff': gd['staff'],
        'principal': gd['principal'],
        'total': gd['total'],
        'rate': rate,
        'branchRate': 0, 'diffPoint': 0, 'diffAmount': 0,
        'firstAssign': gd['firstAssign'],
        'firstBack': gd['firstBack'],
        'firstRate': first_rate,
        'branchFirstRate': 0,
        'oldRemain': 0, 'todayOld': gd['todayOld'],
        'oldRate': 0, 'branchOldRate': 0,
    })

# 将自然回收数据按人力占比分配到各小组
if (manpower_ratio or overall_manpower_ratio) and natural_recovery_rows:
    print(f'  分配自然回收数据到小组...')
    for name, nr in natural_recovery_rows.items():
        region = nr['region']
        principal = nr['principal']
        total = nr['total']
        if region == '郑州':
            target_groups = ['ZZ-F1', 'ZZ-F2', 'ZZ-F3', 'ZZ-F4', 'ZZ-F5']
            ratio_map = manpower_ratio
        elif region == '成都二':
            target_groups = ['CD-F3']
            ratio_map = manpower_ratio
        elif region == '成都一' or region == '成都':
            target_groups = ['CD-F1', 'CD-F2']
            ratio_map = manpower_ratio
        else:
            # 自然回收（无明确地区）→ 按全量人力占比分配到所有小组
            target_groups = ['ZZ-F1', 'ZZ-F2', 'ZZ-F3', 'ZZ-F4', 'ZZ-F5', 'CD-F1', 'CD-F2', 'CD-F3']
            ratio_map = overall_manpower_ratio
        for g in target_groups:
            ratio = ratio_map.get(g, 0)
            if ratio > 0:
                add_principal = round(principal * ratio, 2)
                add_total = round(total * ratio, 2)
                for gs in group_summary:
                    if gs['group'] == g:
                        gs['principal'] += add_principal
                        gs['total'] += add_total
                        break
        print(f'    {name}({region}): 分案{principal:,.2f} / 回退{total:,.2f} 已按人力占比分配')
    # 重新计算各小组的回退率
    for gs in group_summary:
        if gs['principal'] > 0:
            gs['rate'] = gs['total'] / gs['principal']

    # 计算分公司（三地）回退率
    cd1_groups = [gs for gs in group_summary if gs['group'] in ('CD-F1', 'CD-F2')]
    cd2_groups = [gs for gs in group_summary if gs['group'] == 'CD-F3']
    zz_groups = [gs for gs in group_summary if gs['group'].startswith('ZZ')]

    def calc_branch(groups_list):
        p = sum(gs['principal'] for gs in groups_list)
        t = sum(gs['total'] for gs in groups_list)
        fa = sum(gs['firstAssign'] for gs in groups_list)
        fb = sum(gs['firstBack'] for gs in groups_list)
        to = sum(gs['todayOld'] for gs in groups_list)
        return {
            'rate': t / p if p > 0 else 0,
            'firstRate': fb / fa if fa > 0 else 0,
            'oldRate': to / p if p > 0 else 0,
        }

    cd1 = calc_branch(cd1_groups)
    cd2 = calc_branch(cd2_groups)
    zz = calc_branch(zz_groups)

    print(f'  成都一分公司回退率: {cd1["rate"]:.4f} ({cd1["rate"]*100:.2f}%)')
    print(f'  成都二分公司回退率: {cd2["rate"]:.4f} ({cd2["rate"]*100:.2f}%)')
    print(f'  郑州分公司回退率: {zz["rate"]:.4f} ({zz["rate"]*100:.2f}%)')

    # 更新各小组的分公司回退率、差异点位、差异金额
    for gs in group_summary:
        if gs['group'] == 'CD-F3':
            gs['branchRate'] = cd2['rate']
            gs['branchFirstRate'] = cd2['firstRate']
            gs['branchOldRate'] = cd2['oldRate']
        elif gs['group'].startswith('CD'):
            gs['branchRate'] = cd1['rate']
            gs['branchFirstRate'] = cd1['firstRate']
            gs['branchOldRate'] = cd1['oldRate']
        else:
            gs['branchRate'] = zz['rate']
            gs['branchFirstRate'] = zz['firstRate']
            gs['branchOldRate'] = zz['oldRate']
        gs['diffPoint'] = gs['rate'] - gs['branchRate']
        gs['diffAmount'] = gs['diffPoint'] * gs['principal']
        gs['oldRate'] = gs['todayOld'] / gs['principal'] if gs['principal'] > 0 else 0

print(f'  GROUP_SUMMARY: {len(group_summary)} rows')

# 写 group_summary.csv
with open(os.path.join(RESULT_DIR, 'group_summary.csv'), 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['section', '组别', '分案人力', '分案剩余本金', '累计回退', '累计回退率',
                '分公司回退率', '差异点位', '差异金额', '首催分案', '首催回退', '首催回退率',
                '分公司首催回退率', '剩余老案', '当日老案回退', '当日老案回退率', '分公司当日老案回退率'])
    for r in group_summary:
        w.writerow([r['section'], r['group'], r['staff'], r['principal'], r['total'],
                    r['rate'], r['branchRate'], r['diffPoint'], r['diffAmount'],
                    r['firstAssign'], r['firstBack'], r['firstRate'],
                    r['branchFirstRate'], r['oldRemain'], r['todayOld'],
                    r['oldRate'], r['branchOldRate']])

# ============================================================
# 4b. 构建 CATEGORY_GROUP_SUMMARY（各金额段小组汇总）
# ============================================================
category_group_data = {}
for cat in ['大额', '常规', '小额']:
    cat_df = df[df['类别'] == cat]
    group_data = {}
    for _, r in cat_df.iterrows():
        g = str(r['组别']).strip()
        if not g or g == 'nan':
            continue
        if g not in group_data:
            group_data[g] = {
                'principal': 0, 'total': 0, 'staff': 0,
                'firstAssign': 0, 'firstBack': 0, 'todayOld': 0,
            }
        gd = group_data[g]
        gd['staff'] += 1
        gd['principal'] += float(r['分案剩余本金']) if not pd.isna(r['分案剩余本金']) else 0
        gd['total'] += float(r['累计回退']) if not pd.isna(r['累计回退']) else 0
        gd['firstAssign'] += float(r['首催分案']) if not pd.isna(r['首催分案']) else 0
        gd['firstBack'] += float(r['首催回退']) if not pd.isna(r['首催回退']) else 0
        gd['todayOld'] += float(r['当日老案回退']) if not pd.isna(r['当日老案回退']) else 0

    # 找到首催回退率最高的小组，以其首催分案为基准计算差异金额
    best_first_rate = -1
    best_first_assign = 0
    for gname, gd in group_data.items():
        fr = gd['firstBack'] / gd['firstAssign'] if gd['firstAssign'] > 0 else 0
        if fr > best_first_rate:
            best_first_rate = fr
            best_first_assign = gd['firstAssign']

    summary = []
    for gname, gd in sorted(group_data.items()):
        rate = gd['total'] / gd['principal'] if gd['principal'] > 0 else 0
        first_rate = gd['firstBack'] / gd['firstAssign'] if gd['firstAssign'] > 0 else 0
        first_diff = best_first_assign - gd['firstAssign']
        summary.append({
            'group': gname,
            'staff': gd['staff'],
            'principal': gd['principal'],
            'total': gd['total'],
            'rate': rate,
            'firstDiff': first_diff,
            'firstAssign': gd['firstAssign'],
            'firstBack': gd['firstBack'],
            'firstRate': first_rate,
            'todayOld': gd['todayOld'],
        })
    category_group_data[cat] = summary
    print(f'  CATEGORY_GROUP_SUMMARY[{cat}]: {len(summary)} rows')

# 写各金额段 CSVs
for cat in ['大额', '常规', '小额']:
    cat_summary = category_group_data[cat]
    filename = f'category_group_{cat}.csv'
    with open(os.path.join(RESULT_DIR, filename), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['组别', '分案人力', '分案剩余本金', '累计回退', '累计回退率',
                    '与首催第一差异金额', '首催分案', '首催回退', '首催回退率', '当日老案回退'])
        for r in cat_summary:
            w.writerow([r['group'], r['staff'], r['principal'], r['total'],
                        r['rate'], r['firstDiff'], r['firstAssign'], r['firstBack'],
                        r['firstRate'], r['todayOld']])

# 向后兼容
large_group_summary = category_group_data.get('大额', [])

# ============================================================
# 5. 计算 TODAY_BREAKDOWN
# ============================================================
# 5. 计算 TODAY_BREAKDOWN
# ============================================================
rank_idx = {r['name']: r for r in rank_rows}
today_idx = {r['name']: r for r in today_rows}

breakdown = []
for r in rank_rows:
    name = r['name']
    td = today_idx.get(name, {})
    region = get_region(r['group'])
    principal_val = r.get('principal', 0)
    cumul_val = r.get('total', 0)
    cumul_rate = round(cumul_val / principal_val, 4) if principal_val > 0 else 0
    breakdown.append({
        'name': name,
        'firstBack': r.get('firstBack', 0),
        'oldBack': r.get('todayOld', 0),
        'todayTotal': td.get('total', 0),
        'group': r['group'],
        'region': region,
        'category': r.get('category', '常规'),
        'rank': r.get('rank', 0),
        'assignPrincipal': principal_val,
        'cumulBack': cumul_val,
        'cumulBackRate': cumul_rate,
    })

# 补充 TODAY_DETAIL 中有但 CUMULATIVE_RANK 中没有的人员
for td_name, td in today_idx.items():
    if td_name not in {r['name'] for r in rank_rows}:
        region = get_region(td['group'])
        breakdown.append({
            'name': td_name,
            'firstBack': 0,
            'oldBack': 0,
            'todayTotal': td.get('total', 0),
            'group': td['group'],
            'region': region,
            'category': td.get('category', '常规'),
            'rank': 0,
            'assignPrincipal': 0,
            'cumulBack': 0,
            'cumulBackRate': 0,
        })

breakdown.sort(key=lambda x: -x['cumulBack'])
today_breakdown_all = breakdown
today_breakdown_zz = [r for r in breakdown if r['region'] == '郑州' and r['cumulBack'] > 0]
today_breakdown_cd1 = [r for r in breakdown if r['region'] == '成都一' and r['cumulBack'] > 0]
today_breakdown_cd2 = [r for r in breakdown if r['region'] == '成都二' and r['cumulBack'] > 0]

# ============================================================
# 6. 时段汇总
# ============================================================
seg_keys = ['s1', 's2', 's3', 's4', 's5', 's6', 's7']
seg_labels = ['10点前', '10-12点', '12-14点', '14-16点', '16-18点', '18-20点', '20-23点']

hourly = {k: 0 for k in seg_keys}
cat_hourly = {}
reg_hourly = {}

for r in today_rows:
    cat = r['category']
    reg = r['region']
    if cat not in cat_hourly:
        cat_hourly[cat] = {k: 0 for k in seg_keys}
    if reg not in reg_hourly:
        reg_hourly[reg] = {k: 0 for k in seg_keys}
    for k in seg_keys:
        v = r[k]
        hourly[k] += v
        cat_hourly[cat][k] += v
        reg_hourly[reg][k] += v

# ============================================================
# 7. KPI 汇总
# ============================================================
today_total = sum(r['total'] for r in today_rows) + sum(nr.get('today', 0) for nr in natural_recovery_rows.values())
today_active = sum(1 for r in today_rows if r['total'] > 1)
cumul_total = sum(r['total'] for r in rank_rows)
all_cases = sum(r['cases'] for r in rank_rows)

cat_totals = {}
for r in today_rows:
    cat = r['category']
    cat_totals[cat] = cat_totals.get(cat, 0) + r['total']

zz_total = sum(r['total'] for r in today_rows if r['region'] == '郑州')
cd1_total = sum(r['total'] for r in today_rows if r['region'] == '成都一')
cd2_total = sum(r['total'] for r in today_rows if r['region'] == '成都二')

top_today = max(today_rows, key=lambda x: x['total']) if today_rows else {'name': '', 'total': 0}
rank_sorted = sorted(rank_rows, key=lambda x: -x['total'])
top10 = rank_sorted[:10]
today_sorted = sorted(today_rows, key=lambda x: -x['total'])
top20 = [r for r in today_sorted if r['total'] > 0][:20]

group_today = {}
for r in today_rows:
    g = r['group']
    group_today[g] = group_today.get(g, 0) + r['total']

# ============================================================
# 8. 解析 M1 厂商绩效数据
# ============================================================
m1_data = []
if M1_PATH and os.path.exists(M1_PATH):
    try:
        m1_df = pd.read_excel(M1_PATH)
        print(f'M1数据: {M1_PATH} ({len(m1_df)} 行)')
        for _, row in m1_df.iterrows():
            item = {
                'company': str(row.get('公司', '')).strip(),
                'product': str(row.get('产品', '')).strip(),
                'queue': str(row.get('队列', '')).strip(),
                'recoveryRate': 0, 'amountRate': 0, 'accountRate': 0,
                'todayProgress': 0, 'isSubtotal': False,
            }
            # 解析百分比
            for col, key in [('剩余本金回退率', 'recoveryRate'), ('金额回收率', 'amountRate'),
                             ('户数回收率', 'accountRate'), ('今日进度', 'todayProgress')]:
                v = row.get(col, '0')
                try:
                    val = float(str(v).replace('%', ''))
                    if val < 1 and val > 0: val = val * 100  # 自动检测小数格式
                    item[key] = val
                except:
                    item[key] = 0
            is_xy = item['company'].lower() == 'xy'
            item['isSubtotal'] = is_xy
            if is_xy:
                for c in ['分案逾期本息', '分案剩余本金', '回退剩余本金']:
                    item[c] = str(row.get(c, '--'))
                item['assignPrincipal'] = str(row.get('分案剩余本金', '--'))
                try:
                    item['assignAccounts'] = int(row.get('分案户数', 0))
                except:
                    item['assignAccounts'] = 0
            m1_data.append(item)
        print(f'  M1_DATA: {len(m1_data)} rows, {sum(1 for r in m1_data if r["isSubtotal"])} subtotal')
    except Exception as e:
        print(f'  [WARN] M1解析失败: {e}')

# ============================================================
# 8a2. 解析上月数据（同期对比）
# ============================================================
last_month_data = []
last_month_region_data = []
last_month_group_data = []
LM_PATH = os.path.join(DATA_DIR, '上月数据.xlsx')
if not os.path.exists(LM_PATH):
    LM_PATH = os.path.join(os.path.expanduser('~'), 'Desktop', '上月数据.xlsx')

if os.path.exists(LM_PATH):
    try:
        lm_df = pd.read_excel(LM_PATH)
        print(f'上月数据: {LM_PATH} ({len(lm_df)} 行)')
        # 取上月同日数据（如今天6月2日 → 匹配5月2日）
        from datetime import datetime, timedelta
        today = datetime.now()
        # 计算上月同日：当前月份减1，处理跨年
        if today.month == 1:
            target_month, target_year = 12, today.year - 1
        else:
            target_month, target_year = today.month - 1, today.year
        target_day = today.day
        target_date = datetime(target_year, target_month, 1)
        # 处理上月天数不足的情况（如3月31日 → 上月2月只有28/29天）
        import calendar
        last_day_of_target = calendar.monthrange(target_year, target_month)[1]
        if target_day > last_day_of_target:
            target_day = last_day_of_target
        # 筛选匹配日期的数据
        lm_latest = lm_df[lm_df.iloc[:, 0].apply(lambda d: d.day == target_day and d.month == target_month and d.year == target_year)]
        # 如果同日数据不存在（如数据缺失），回退到最新可用日期
        if len(lm_latest) == 0:
            latest_date = lm_df.iloc[:, 0].max()
            lm_latest = lm_df[lm_df.iloc[:, 0] == latest_date]
            print(f'  上月同日({target_year}/{target_month:02d}/{target_day:02d})无数据，回退到最新: {str(latest_date)[:10]}')
        else:
            print(f'  匹配上月同日: {target_year}/{target_month:02d}/{target_day:02d}，共{len(lm_latest)}条')
        for _, row in lm_latest.iterrows():
            company = str(row.iloc[1]).strip()
            def parse_rate(raw_val):
                try:
                    v = float(str(raw_val).replace('%', ''))
                    # 归一化：如果值 < 1 且不为0，说明是小数格式（如0.6604=66.04%），需×100
                    if 0 < v < 1:
                        v = v * 100
                    return v
                except:
                    return 0
            recovery_rate = parse_rate(row.iloc[3])
            daily_progress = parse_rate(row.iloc[5])
            last_month_data.append({
                'company': company,
                'recoveryRate': recovery_rate,
                'dailyProgress': daily_progress,
            })
        print(f'  LAST_MONTH_M1_DATA: {len(last_month_data)} rows')

        # 解析地区 sheet（两地）
        last_month_region_data = []
        col_name = f'{target_month}.{target_day}日'
        try:
            lm_region_df = pd.read_excel(LM_PATH, sheet_name=1)  # Sheet 1 = 地区
            if col_name in lm_region_df.columns:
                for _, row in lm_region_df.iterrows():
                    region_name = str(row.iloc[0]).strip()
                    rate = parse_rate(row[col_name])
                    last_month_region_data.append({
                        'region': region_name,
                        'rate': rate,
                    })
                print(f'  LAST_MONTH_REGION_DATA: {len(last_month_region_data)} rows (col: {col_name})')
            else:
                print(f'  [WARN] 地区 sheet 中未找到列 {col_name}，可用列: {list(lm_region_df.columns)[:5]}...')
        except Exception as e:
            print(f'  [WARN] 地区 sheet 解析失败: {e}')

        # 解析小组 sheet
        last_month_group_data = []
        try:
            lm_group_df = pd.read_excel(LM_PATH, sheet_name=2)  # Sheet 2 = 小组
            if col_name in lm_group_df.columns:
                for _, row in lm_group_df.iterrows():
                    group_name = str(row.iloc[0]).strip()
                    rate = parse_rate(row[col_name])
                    last_month_group_data.append({
                        'group': group_name,
                        'rate': rate,
                    })
                print(f'  LAST_MONTH_GROUP_DATA: {len(last_month_group_data)} rows (col: {col_name})')
            else:
                print(f'  [WARN] 小组 sheet 中未找到列 {col_name}，可用列: {list(lm_group_df.columns)[:5]}...')
        except Exception as e:
            print(f'  [WARN] 小组 sheet 解析失败: {e}')

    except Exception as e:
        print(f'  [WARN] 上月数据解析失败: {e}')

# ============================================================
# 8b. 解析首催排名数据
# ============================================================
first_collection_data = []
FC_PATH = None
# 1) 看板数据目录下的 首催.xlsx
fc_candidate = os.path.join(DATA_DIR, '首催.xlsx')
if os.path.exists(fc_candidate):
    FC_PATH = fc_candidate
# 2) 桌面上的 首催.xlsx
if not FC_PATH:
    fc_desktop = os.path.join(os.path.expanduser('~'), 'Desktop', '首催.xlsx')
    if os.path.exists(fc_desktop):
        FC_PATH = fc_desktop

if FC_PATH and os.path.exists(FC_PATH):
    try:
        fc_df = pd.read_excel(FC_PATH)
        print(f'首催数据: {FC_PATH} ({len(fc_df)} 行)')
        for _, row in fc_df.iterrows():
            item = {
                'company': str(row.get('公司', '')).strip(),
                'product': str(row.get('产品', '')).strip(),
                'queue': str(row.get('队列', '')).strip(),
                'recoveryRate': 0, 'amountRate': 0, 'accountRate': 0,
                'todayProgress': 0, 'isSubtotal': False,
            }
            for col, key in [('剩余本金回退率', 'recoveryRate'), ('金额回收率', 'amountRate'),
                             ('户数回收率', 'accountRate'), ('今日进度', 'todayProgress')]:
                v = row.get(col, '0')
                try:
                    val = float(str(v).replace('%', ''))
                    if val < 1 and val > 0: val = val * 100  # 自动检测小数格式
                    item[key] = val
                except:
                    item[key] = 0
            is_xy = item['company'].lower() == 'xy'
            item['isSubtotal'] = is_xy
            if is_xy:
                for c in ['分案逾期本息', '分案剩余本金', '回退剩余本金']:
                    item[c] = str(row.get(c, '--'))
                item['assignPrincipal'] = str(row.get('分案剩余本金', '--'))
                try:
                    item['assignAccounts'] = int(row.get('分案户数', 0))
                except:
                    item['assignAccounts'] = 0
            first_collection_data.append(item)
        print(f'  FIRST_COLLECTION_DATA: {len(first_collection_data)} rows')
    except Exception as e:
        print(f'  [WARN] 首催解析失败: {e}')

# ============================================================
# 8c. 解析催员产能数据
# ============================================================
staff_productivity_data = []
SP_PATH = os.path.join(DATA_DIR, '催员产能.xlsx')
if not os.path.exists(SP_PATH):
    SP_PATH = os.path.join(os.path.expanduser('~'), 'Desktop', '催员产能.xlsx')

# Build agent→group and agent→name maps from 人员.xlsx (staff roster)
PERSONNEL_PATH = os.path.join(DATA_DIR, '人员.xlsx')
agent_group_map = {}
agent_name_map = {}
if os.path.exists(PERSONNEL_PATH):
    try:
        personnel_df = pd.read_excel(PERSONNEL_PATH)
        for _, r in personnel_df.iterrows():
            agent = str(r.iloc[0]).strip()
            nm = str(r.iloc[1]).strip()  # Col 1 = 姓名
            g = str(r.iloc[2]).strip()   # Col 2 = 组别
            if agent and agent != 'nan':
                if g and g != 'nan': agent_group_map[agent] = g
                if nm and nm != 'nan': agent_name_map[agent] = nm
        print(f'  人员表: {len(agent_name_map)} staff mapped')
    except Exception as e:
        print(f'  [WARN] 人员表读取失败: {e}')
# Fallback: also merge names from 大表
for _, r in df.iterrows():
    agent = str(r.iloc[0]).strip()
    nm = str(r.iloc[1]).strip()
    g = str(r.iloc[2]).strip()
    if agent and agent != 'nan':
        if agent not in agent_group_map and g and g != 'nan':
            agent_group_map[agent] = g
        if agent not in agent_name_map and nm and nm != 'nan':
            agent_name_map[agent] = nm

if os.path.exists(SP_PATH):
    try:
        sp_xl = pd.ExcelFile(SP_PATH)
        sp_df = pd.read_excel(SP_PATH, sheet_name=sp_xl.sheet_names[0])
        # Columns (by index): 0=统计时间, 1=日期, 2=产品, 3=队列, 4=公司, 5=催员,
        #   6=在催数, 7=当月新分案, 8=当月PTP, 9=出催数(拨打量), 10=预期回款(短信量),
        #   11=一线回收, 12=人工协同, 13=预计回款, 14=手自动回款,
        #   15=累计数, 16=回款数, 17=总计数(扣款量),
        #   18=新分案覆盖率, 19=案件覆盖率(拨打覆盖率), 20=累计覆盖率,
        #   21=人工接通率, 22=人工外呼接通率, 23-25=通话相关
        latest_date = str(sp_df.iloc[:, 1].max())
        latest = sp_df[sp_df.iloc[:, 1].astype(str) == latest_date]
        # Group by staff+queue, sum numeric columns
        latest['_staff'] = latest.iloc[:, 5].astype(str).str.strip()
        latest['_queue'] = latest.iloc[:, 3].astype(str).str.strip()
        # Exclude F5 queue
        latest = latest[latest['_queue'] != 'F5']
        staff_groups = latest.groupby(['_staff', '_queue'])
        for (name, queue), grp in staff_groups:
            if not name or name == 'nan':
                continue
            # Parse percentage strings
            def parse_pct(val_series):
                v = str(val_series.iloc[0]) if len(val_series) > 0 else '0'
                try: return float(v.replace('%', ''))
                except: return 0

            staff_productivity_data.append({
                'name': name,
                'group': agent_group_map.get(name, queue),
                'activeCases': int(grp.iloc[:, 6].sum()) if not pd.isna(grp.iloc[:, 6].sum()) else 0,
                'callVolume': float(grp.iloc[:, 9].sum()) if not pd.isna(grp.iloc[:, 9].sum()) else 0,
                'smsVolume': float(grp.iloc[:, 10].sum()) if not pd.isna(grp.iloc[:, 10].sum()) else 0,
                'deductionCount': float(grp.iloc[:, 17].sum()) if not pd.isna(grp.iloc[:, 17].sum()) else 0,
                'newCaseCoverage': parse_pct(grp.iloc[:, 18]),
                'callCoverage': parse_pct(grp.iloc[:, 19]),
            })
        # Sort by group order, then by activeCases desc
        GROUP_ORDER = ['ZZ-F1', 'ZZ-F2', 'ZZ-F3', 'ZZ-F4', 'ZZ-F5', 'CD-F1', 'CD-F2', 'CD-F3']
        def group_sort_key(item):
            g = item.get('group', '')
            try: gi = GROUP_ORDER.index(g)
            except ValueError: gi = 99
            return (gi, -item['activeCases'])
        staff_productivity_data.sort(key=group_sort_key)
        print(f'  催员产能: {len(staff_productivity_data)} staff on {latest_date}')
    except Exception as e:
        print(f'  [WARN] 催员产能解析失败: {e}')
        import traceback
        traceback.print_exc()

# ============================================================
# 8d. 解析一阶段小组数据
# ============================================================
stage1_group_summary = []
STAGE1_PATH = os.path.join(DATA_DIR, '一阶段小组信息.xlsx')
if not os.path.exists(STAGE1_PATH):
    STAGE1_PATH = os.path.join(os.path.expanduser('~'), 'Desktop', '一阶段小组信息.xlsx')

if os.path.exists(STAGE1_PATH):
    try:
        s1_df = pd.read_excel(STAGE1_PATH)
        print(f'一阶段小组数据: {STAGE1_PATH} ({len(s1_df)} 行)')
        for _, row in s1_df.iterrows():
            principal = float(row.iloc[1]) if not pd.isna(row.iloc[1]) else 0
            total = float(row.iloc[2]) if not pd.isna(row.iloc[2]) else 0
            rate = total / principal if principal > 0 else 0
            stage1_group_summary.append({
                'group': str(row.iloc[0]).strip(),
                'principal': principal,
                'total': total,
                'rate': rate,
            })
        # 计算一阶段分公司回退率
        s1_cd_principal = sum(r['principal'] for r in stage1_group_summary if r['group'].startswith('CD'))
        s1_cd_total = sum(r['total'] for r in stage1_group_summary if r['group'].startswith('CD'))
        s1_zz_principal = sum(r['principal'] for r in stage1_group_summary if r['group'].startswith('ZZ'))
        s1_zz_total = sum(r['total'] for r in stage1_group_summary if r['group'].startswith('ZZ'))
        s1_cd_rate = s1_cd_total / s1_cd_principal if s1_cd_principal > 0 else 0
        s1_zz_rate = s1_zz_total / s1_zz_principal if s1_zz_principal > 0 else 0
        print(f'  一阶段成都回退率: {s1_cd_rate:.4f} ({s1_cd_rate*100:.2f}%)')
        print(f'  一阶段郑州回退率: {s1_zz_rate:.4f} ({s1_zz_rate*100:.2f}%)')
        print(f'  STAGE1_GROUP_SUMMARY: {len(stage1_group_summary)} rows')
    except Exception as e:
        print(f'  [WARN] 一阶段小组数据解析失败: {e}')
# ============================================================
# 8e. 读取对公数据
corporate_data = {'ZZ': 0, 'CD': 0}
CORP_PATH = os.path.join(DATA_DIR, '对公.xlsx')
if os.path.exists(CORP_PATH):
    try:
        corp_df = pd.read_excel(CORP_PATH, sheet_name='Sheet1')
        if corp_df.shape[1] >= 2:
            corporate_data['ZZ'] = float(corp_df.columns[1])
            corporate_data['CD'] = float(corp_df.iloc[0, 1])
        elif corp_df.shape[1] == 1 and corp_df.shape[0] >= 2:
            corporate_data['ZZ'] = float(corp_df.iloc[0, 0])
            corporate_data['CD'] = float(corp_df.iloc[1, 0])
        print(f'  对公数据: ZZ={corporate_data["ZZ"]:,.0f}, CD={corporate_data["CD"]:,.0f}')
    except Exception as e:
        print(f'  [WARN] 对公数据解析失败: {e}')

# 8e. 读取对公数据
corporate_data = {'ZZ': 0, 'CD': 0}
CORP_PATH = os.path.join(DATA_DIR, '对公.xlsx')
if os.path.exists(CORP_PATH):
    try:
        corp_df = pd.read_excel(CORP_PATH, sheet_name='Sheet1')
        corporate_data['ZZ'] = float(corp_df.columns[1])
        corporate_data['CD'] = float(corp_df.iloc[0, 1])
        print(f'  对公数据: ZZ={corporate_data["ZZ"]:,.0f}, CD={corporate_data["CD"]:,.0f}')
    except Exception as e:
        print(f'  [WARN] 对公数据解析失败: {e}')

# ============================================================
# 8f. 解析厂商产能数据
# ============================================================
company_productivity_data = []
CP_PATH = os.path.join(DATA_DIR, '厂商产能.xlsx')
if not os.path.exists(CP_PATH):
    CP_PATH = os.path.join(os.path.expanduser('~'), 'Desktop', '厂商产能.xlsx')

if os.path.exists(CP_PATH):
    try:
        cp_df = pd.read_excel(CP_PATH)
        print(f'厂商产能: {CP_PATH} ({len(cp_df)} 行)')
        for _, row in cp_df.iterrows():
            def cp_val(col_idx):
                v = row.iloc[col_idx]
                if pd.isna(v) or str(v).strip() in ('--', '', 'nan'):
                    return None
                try:
                    return float(str(v).replace('%', '').replace(',', ''))
                except:
                    return str(v).strip()
            company_productivity_data.append({
                'company': str(row.iloc[4]).strip(),        # 公司
                'activeCases': cp_val(5),                    # 在催数
                'assignCases': cp_val(6),                    # 分案数
                'assignStaff': cp_val(7),                    # 分案人力
                'assignStaff2': cp_val(8),                   # 分案人力.1
                'estimatedRecovery': cp_val(9),              # 预计回款
                'newCases': cp_val(10),                      # 当月新案
                'monthPTP': cp_val(11),                      # 当月PTP
                'avgCallVolume': cp_val(12),                 # 人均拨打量
            })
        print(f'  COMPANY_PRODUCTIVITY_DATA: {len(company_productivity_data)} rows')
    except Exception as e:
        print(f'  [WARN] 厂商产能解析失败: {e}')

# 8e. 读取对公数据
corporate_data = {'ZZ': 0, 'CD': 0}
CORP_PATH = os.path.join(DATA_DIR, '对公.xlsx')
if os.path.exists(CORP_PATH):
    try:
        corp_df = pd.read_excel(CORP_PATH, sheet_name='Sheet1')
        corporate_data['ZZ'] = float(corp_df.columns[1])
        corporate_data['CD'] = float(corp_df.iloc[0, 1])
        print(f'  对公数据: ZZ={corporate_data["ZZ"]:,.0f}, CD={corporate_data["CD"]:,.0f}')
    except Exception as e:
        print(f'  ⚠ 对公数据解析失败: {e}')

# 9. 生成 data_v2.js
# ============================================================
js = []
js.append('// 贷后数据看板 v2 - 数据配置文件 (自动生成)\n')
js.append('var DATA_SOURCE_FILE = "大表.xlsx";\n')

# 提取各数据源的统计时间
data_source_times = {}
# 大表无统计时间列，跳过
# 累计.xlsx
if M1_PATH and os.path.exists(M1_PATH):
    try:
        m1_ts = pd.read_excel(M1_PATH).iloc[0, 0]
        data_source_times['累计'] = str(m1_ts)[:16]
    except: pass
# 首催.xlsx
if FC_PATH and os.path.exists(FC_PATH):
    try:
        fc_ts = pd.read_excel(FC_PATH).iloc[0, 0]
        data_source_times['首催'] = str(fc_ts)[:16]
    except: pass
# 催员产能
if SP_PATH and os.path.exists(SP_PATH):
    try:
        sp_ts = pd.read_excel(SP_PATH).iloc[0, 0]
        data_source_times['催员产能'] = str(sp_ts)[:16]
    except: pass

time_info = ', '.join(f'{k}: {v}' for k, v in data_source_times.items())
js.append(f'var DATA_SOURCE_TIME = "{time_info}";\n')
print(f'  数据源时间: {time_info}')

js.append('var TODAY_DETAIL = [')
for r in today_rows:
    js.append(f'  {{ name:"{r["name"]}", total:{r["total"]:.2f}, '
              f's1:{r["s1"]:.2f}, s2:{r["s2"]:.2f}, s3:{r["s3"]:.2f}, '
              f's4:{r["s4"]:.2f}, s5:{r["s5"]:.2f}, s6:{r["s6"]:.2f}, '
              f's7:{r["s7"]:.2f}, rank:{r["rank"]}, '
              f'region:"{r["region"]}", group:"{r["group"]}", '
              f'category:"{r["category"]}" }},')
js.append('];\n')

js.append('var CUMULATIVE_RANK = [')
for r in rank_rows:
    js.append(f'  {{ name:"{r["name"]}", group:"{r["group"]}", '
              f'cases:{r["cases"]}, principal:{r["principal"]:.2f}, '
              f'total:{r["total"]:.2f}, rate:{r["rate"]:.4f}, '
              f'diff:{r["diff"]:.2f}, firstAmount:{r["firstAmount"]:.2f}, '
              f'firstBack:{r["firstBack"]:.2f}, firstRate:{r["firstRate"]:.4f}, '
              f'rank:{r["rank"]}, todayOld:{r["todayOld"]:.2f}, '
              f'category:"{r["category"]}" }},')
js.append('];\n')

js.append('// 今日回款按催员明细')
js.append(f'var TODAY_BREAKDOWN_ALL = {json.dumps(today_breakdown_all, ensure_ascii=False)};')
js.append(f'var TODAY_BREAKDOWN_ZZ = {json.dumps(today_breakdown_zz, ensure_ascii=False)};')
js.append(f'var TODAY_BREAKDOWN_CD1 = {json.dumps(today_breakdown_cd1, ensure_ascii=False)};')
js.append(f'var TODAY_BREAKDOWN_CD2 = {json.dumps(today_breakdown_cd2, ensure_ascii=False)};\n')

js.append(f'var HOURLY_SEGMENTS = {json.dumps(hourly)};')
js.append(f'var CATEGORY_HOURLY = {json.dumps(cat_hourly)};')
js.append(f'var REGION_HOURLY = {json.dumps(reg_hourly)};\n')

js.append('var GROUP_SUMMARY = [')
for r in group_summary:
    js.append(f'  {{ section:"{r["section"]}", group:"{r["group"]}", '
              f'staff:{r["staff"]}, principal:{r["principal"]:.2f}, '
              f'total:{r["total"]:.2f}, rate:{r["rate"]:.4f}, '
              f'branchRate:{r["branchRate"]:.4f}, diffPoint:{r["diffPoint"]:.4f}, '
              f'diffAmount:{r["diffAmount"]:.2f}, firstAssign:{r["firstAssign"]:.2f}, '
              f'firstBack:{r["firstBack"]:.2f}, firstRate:{r["firstRate"]:.4f}, '
              f'branchFirstRate:{r["branchFirstRate"]:.4f}, oldRemain:{r["oldRemain"]:.2f}, '
              f'todayOld:{r["todayOld"]:.2f}, oldRate:{r["oldRate"]:.4f}, '
              f'branchOldRate:{r["branchOldRate"]:.4f} }},')
js.append('];\n')

js.append('// 各金额段小组数据汇总')
js.append('var CATEGORY_GROUP_SUMMARY = {')
for cat in ['大额', '常规', '小额']:
    cat_summary = category_group_data[cat]
    js.append(f'  "{cat}": [')
    for r in cat_summary:
        js.append(f'    {{ group:"{r["group"]}", staff:{r["staff"]}, principal:{r["principal"]:.2f}, '
                  f'total:{r["total"]:.2f}, rate:{r["rate"]:.4f}, '
                  f'firstDiff:{r["firstDiff"]:.2f}, firstAssign:{r["firstAssign"]:.2f}, '
                  f'firstBack:{r["firstBack"]:.2f}, firstRate:{r["firstRate"]:.4f}, '
                  f'todayOld:{r["todayOld"]:.2f} }},')
    js.append('  ],')
js.append('};')
js.append('var LARGE_GROUP_SUMMARY = CATEGORY_GROUP_SUMMARY["大额"] || [];\n')

js.append('// 一阶段小组数据（来源于 一阶段小组信息.xlsx）')
js.append('var STAGE1_GROUP_SUMMARY = [')
for r in stage1_group_summary:
    js.append(f'  {{ group:"{r["group"]}", principal:{r["principal"]:.2f}, total:{r["total"]:.2f}, rate:{r.get("rate", 0):.4f}, staff:0, firstRate:0, todayOld:0 }},')
js.append('];\n')

js.append(f'var CORPORATE_DATA = {{zz: {corporate_data["ZZ"]:.0f}, cd: {corporate_data["CD"]:.0f}}};')
js.append(f'var CORPORATE_DATA = {{zz: {corporate_data["ZZ"]:.0f}, cd: {corporate_data["CD"]:.0f}}};')
js.append(f'var CORPORATE_DATA = {{zz: {corporate_data["ZZ"]:.0f}, cd: {corporate_data["CD"]:.0f}}};')
js.append(f'var CORPORATE_DATA = {{zz: {corporate_data["ZZ"]:.0f}, cd: {corporate_data["CD"]:.0f}}};')
js.append(f'var CORPORATE_DATA = {{zz: {corporate_data["ZZ"]:.0f}, cd: {corporate_data["CD"]:.0f}}};')
js.append(f'var CORPORATE_DATA = {{zz: {corporate_data["ZZ"]:.0f}, cd: {corporate_data["CD"]:.0f}}};')
js.append(f'var KPI_TODAY_TOTAL = {today_total:.2f};')
js.append(f'var KPI_TODAY_ACTIVE = {today_active};')
js.append(f'var KPI_CUMUL_TOTAL = {cumul_total:.2f};')
js.append(f'var KPI_ALL_CASES = {all_cases};')
js.append(f'var KPI_ZZ_TOTAL = {zz_total:.2f};')
js.append(f'var KPI_CD1_TOTAL = {cd1_total:.2f};')
js.append(f'var KPI_CD2_TOTAL = {cd2_total:.2f};')
js.append(f'var KPI_CAT_TOTALS = {json.dumps(cat_totals)};')
js.append(f'var KPI_STAFF_COUNT = {len(today_rows)};')
js.append(f'var KPI_TOP_TODAY = {{ name:"{top_today["name"]}", amount:{top_today["total"]:.2f} }};')
js.append(f'var TOP10_NAMES = {json.dumps([r["name"] for r in top10])};')
js.append(f'var TOP10_AMOUNTS = {json.dumps([r["total"] for r in top10])};')
js.append(f'var TOP10_RATES = {json.dumps([r["rate"] for r in top10])};')
js.append(f'var TOP10_GROUPS = {json.dumps([r["group"] for r in top10])};')
# Derive region from group for top10 (CUMULATIVE_RANK entries don't have a region field)
def _top10_region(g):
    g = str(g).strip().upper()
    if g.startswith('ZZ'): return '郑州'
    if g == 'CD-F3': return '成都二'
    if g.startswith('CD'): return '成都一'
    return '郑州'
js.append(f'var TOP10_REGIONS = {json.dumps([_top10_region(r["group"]) for r in top10])};')
js.append(f'var TODAY_TOP20_NAMES = {json.dumps([r["name"] for r in top20])};')
js.append(f'var TODAY_TOP20_AMOUNTS = {json.dumps([r["total"] for r in top20])};')
js.append(f'var TODAY_TOP20_GROUPS = {json.dumps([r["group"] for r in top20])};')
js.append(f'var TODAY_TOP20_REGIONS = {json.dumps([r["region"] for r in top20])};')
js.append(f'var GROUP_TODAY_LABELS = {json.dumps(list(group_today.keys()))};')
js.append(f'var GROUP_TODAY_AMOUNTS = {json.dumps(list(group_today.values()))};')
js.append(f'var SEGMENT_LABELS = {json.dumps(seg_labels)};')
js.append(f'var SEGMENT_KEYS = {json.dumps(seg_keys)};')
js.append(f'var M1_DATA = {json.dumps(m1_data, ensure_ascii=False)};')
js.append(f'var LAST_MONTH_M1_DATA = {json.dumps(last_month_data, ensure_ascii=False)};')
js.append(f'var LAST_MONTH_REGION_DATA = {json.dumps(last_month_region_data, ensure_ascii=False)};')
js.append(f'var LAST_MONTH_GROUP_DATA = {json.dumps(last_month_group_data, ensure_ascii=False)};')
js.append(f'var FIRST_COLLECTION_DATA = {json.dumps(first_collection_data, ensure_ascii=False)};')
js.append(f'var STAFF_PRODUCTIVITY_DATA = {json.dumps(staff_productivity_data, ensure_ascii=False)};')
js.append(f'var STAFF_GROUP = {json.dumps(agent_group_map, ensure_ascii=False)};')
js.append(f'var COMPANY_PRODUCTIVITY_DATA = {json.dumps(company_productivity_data, ensure_ascii=False)};')

# 生成 STAFF_NAME 映射（坐席 -> 催员姓名）
staff_name_map = dict(agent_name_map)  # Start with names from 大表
# Also add staff from 催员产能 who may not be in the main table (use agent ID as fallback)
for sp in staff_productivity_data:
    if sp['name'] not in staff_name_map:
        staff_name_map[sp['name']] = sp['name']
js.append(f'var STAFF_NAME = {json.dumps(staff_name_map, ensure_ascii=False)};')

outpath = os.path.join(BASE, 'data_v2.js')
with open(outpath, 'w', encoding='utf-8') as f:
    f.write('\n'.join(js))

# 生成 base64 加密版本，避免查看源代码直接暴露数据
import base64
js_raw = '\n'.join(js)
js_b64 = base64.b64encode(js_raw.encode('utf-8')).decode('ascii')
enc_path = os.path.join(BASE, 'data_v2_enc.js')
with open(enc_path, 'w', encoding='utf-8') as f:
    f.write('var _D="' + js_b64 + '";')
print(f'[OK] data_v2_enc.js: {os.path.getsize(enc_path)} bytes (base64 encoded)')

# 同步 dashboard_v2.html 中的 data_v2_enc.js（如果 HTML 中使用了外部引用模式）
# 注意：HTML 现在使用 <script src="data_v2_enc.js"> + 运行时解码，
# 不再内嵌数据。只需要更新 data_v2_enc.js 文件即可。
# 以下保留对旧版 HTML 的兼容（如果 DATA_END 标记还存在则更新内嵌数据）
import re
dashboard_html = os.path.join(BASE, 'dashboard_v2.html')
if os.path.exists(dashboard_html):
    with open(dashboard_html, 'r', encoding='utf-8') as f:
        html = f.read()

    end_marker = '<!-- DATA_END -->'
    if end_marker in html:
        # 旧版 HTML（数据内嵌），自动替换为外部引用模式
        begin_marker = '<!-- DATA_BEGIN -->'
        new_block = begin_marker + '\n<script src="data_v2_enc.js"></script>\n<script>\ntry {\n  if (typeof _D !== \'string\') throw new Error(\'加密数据加载失败\');\n  var b = _D;\n  var bin = atob(b);\n  var arr = new Uint8Array(bin.length);\n  for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);\n  var js = new TextDecoder(\'utf-8\').decode(arr);\n  (0, eval)(js);\n} catch(e) {\n  document.getElementById(\'updateTime\').textContent = \'数据加载失败: \' + e.message;\n  console.error(\'Data decode error:\', e);\n}\n</script>\n' + end_marker
        start = html.index(begin_marker)
        end = html.index(end_marker) + len(end_marker)
        html = html[:start] + new_block + html[end:]
        with open(dashboard_html, 'w', encoding='utf-8') as f:
            f.write(html)
        print('[OK] dashboard_v2.html 已切换为加密外部文件模式')
    else:
        print('[OK] dashboard_v2.html 已是加密外部文件模式，无需修改')

    # 离线版同步（同时复制加密数据文件到离线目录）
    offline_path = os.path.join(RESULT_DIR, 'dashboard_offline.html')
    with open(offline_path, 'w', encoding='utf-8') as f:
        f.write(html)
    import shutil
    shutil.copy2(enc_path, os.path.join(RESULT_DIR, 'data_v2_enc.js'))

print(f'\n[OK] data_v2.js: {os.path.getsize(outpath)} bytes')
print(f'[OK] TODAY_DETAIL: {len(today_rows)} entries')
print(f'[OK] CUMULATIVE_RANK: {len(rank_rows)} entries')
print(f'[OK] GROUP_SUMMARY: {len(group_summary)} rows')
print(f'[OK] CATEGORY_GROUP_SUMMARY: 大额{len(category_group_data.get("大额",[]))}, 常规{len(category_group_data.get("常规",[]))}, 小额{len(category_group_data.get("小额",[]))} rows')
print(f'[OK] M1_DATA: {len(m1_data)} rows')
print(f'[OK] LAST_MONTH_M1_DATA: {len(last_month_data)} rows')
print(f'[OK] LAST_MONTH_REGION_DATA: {len(last_month_region_data)} rows')
print(f'[OK] LAST_MONTH_GROUP_DATA: {len(last_month_group_data)} rows')
print(f'[OK] 中间 CSV -> {RESULT_DIR}/')
if os.path.exists(os.path.join(RESULT_DIR, 'dashboard_offline.html')):
    print(f'[OK] 离线版 -> {RESULT_DIR}/dashboard_offline.html')
print(f'\n打开 dashboard_v2.html 即可查看看板')
