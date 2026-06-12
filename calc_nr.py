import pandas as pd
import numpy as np
import os

DATA_DIR = r'C:\Users\Administrator\Desktop\看板数据'

# Read data files
df = pd.read_excel(os.path.join(DATA_DIR, '大表.xlsx'))
print('=== 自然回收人员明细 ===')

# Extract natural recovery
natural = {}
for _, r in df.iterrows():
    g = str(r.iloc[2]).strip()
    region = str(r.iloc[3]).strip() if not pd.isna(r.iloc[3]) else ''
    if g in ('郑州', '成都', '自然回收'):
        name = str(r.iloc[0]).strip()
        p = float(r['分案剩余本金']) if not pd.isna(r['分案剩余本金']) else 0
        t = float(r['累计回退']) if not pd.isna(r['累计回退']) else 0
        target = region if g == '自然回收' else g
        if name not in natural:
            natural[name] = {'region': target, 'principal': 0, 'total': 0}
        natural[name]['principal'] += p
        natural[name]['total'] += t

total_nr_p = 0; total_nr_t = 0
cd_nr_p = 0; cd_nr_t = 0
zz_nr_p = 0; zz_nr_t = 0
all_nr_p = 0; all_nr_t = 0

for name, nr in natural.items():
    reg = nr['region']
    p = nr['principal']; t = nr['total']
    total_nr_p += p; total_nr_t += t
    if reg in ('成都',):
        cd_nr_p += p; cd_nr_t += t
    elif reg in ('郑州',):
        zz_nr_p += p; zz_nr_t += t
    else:
        all_nr_p += p; all_nr_t += t
    print(f'  {name}: region={reg}, principal={p:,.2f}, total={t:,.2f}')

print(f'\n自然回收总计: principal={total_nr_p:,.2f}, total={total_nr_t:,.2f}')
print(f'  成都专属: p={cd_nr_p:,.2f}, t={cd_nr_t:,.2f}')
print(f'  郑州专属: p={zz_nr_p:,.2f}, t={zz_nr_t:,.2f}')
print(f'  全部(待分配): p={all_nr_p:,.2f}, t={all_nr_t:,.2f}')

# Read manpower
personnel_path = os.path.join(DATA_DIR, '人员.xlsx')
p_df = pd.read_excel(personnel_path)
zz_counts = {}; cd_counts = {}; all_counts = {}
for _, r in p_df.iterrows():
    g = str(r.iloc[2]).strip()
    if g in ('ZZ-F1','ZZ-F2','ZZ-F3','ZZ-F4','ZZ-F5'):
        zz_counts[g] = zz_counts.get(g, 0) + 1
        all_counts[g] = all_counts.get(g, 0) + 1
    elif g in ('CD-F1','CD-F2'):
        cd_counts[g] = cd_counts.get(g, 0) + 1
        all_counts[g] = all_counts.get(g, 0) + 1

zz_man = sum(zz_counts.values())
cd_man = sum(cd_counts.values())
all_man = sum(all_counts.values())

print(f'\n=== 人力分布 ===')
print(f'郑州: {zz_man}人 ({zz_man/all_man*100:.1f}%)')
for g, c in zz_counts.items():
    print(f'  {g}: {c}人 (ZZ内{c/zz_man*100:.1f}%, 全局{c/all_man*100:.1f}%)')
print(f'成都: {cd_man}人 ({cd_man/all_man*100:.1f}%)')
for g, c in cd_counts.items():
    print(f'  {g}: {c}人 (CD内{c/cd_man*100:.1f}%, 全局{c/all_man*100:.1f}%)')

# Allocate "全部" natural recovery to 两地 by manpower ratio
cd_from_all_p = all_nr_p * cd_man / all_man
cd_from_all_t = all_nr_t * cd_man / all_man
zz_from_all_p = all_nr_p * zz_man / all_man
zz_from_all_t = all_nr_t * zz_man / all_man

# Total natural recovery per location
cd_total_nr_p = cd_from_all_p + cd_nr_p
cd_total_nr_t = cd_from_all_t + cd_nr_t
zz_total_nr_p = zz_from_all_p + zz_nr_p
zz_total_nr_t = zz_from_all_t + zz_nr_t

print(f'\n=== 自然回收 -> 两地分配 ===')
print(f'成都获得: principal={cd_total_nr_p:,.2f}, total={cd_total_nr_t:,.2f}')
print(f'  "全部"按人力分配: p={cd_from_all_p:,.2f}, t={cd_from_all_t:,.2f}')
print(f'  成都专属: p={cd_nr_p:,.2f}, t={cd_nr_t:,.2f}')
print(f'郑州获得: principal={zz_total_nr_p:,.2f}, total={zz_total_nr_t:,.2f}')
print(f'  "全部"按人力分配: p={zz_from_all_p:,.2f}, t={zz_from_all_t:,.2f}')
print(f'  郑州专属: p={zz_nr_p:,.2f}, t={zz_nr_t:,.2f}')

# Allocate to groups
print(f'\n=== 自然回收 -> 小组分配 ===')
cd_alloc = {}
zz_alloc = {}
for g, cnt in cd_counts.items():
    ratio = cnt / cd_man
    p = round(cd_total_nr_p * ratio, 2)
    t = round(cd_total_nr_t * ratio, 2)
    cd_alloc[g] = {'p': p, 't': t, 'ratio': ratio}
    print(f'  {g}: ratio={ratio:.4f}, +principal={p:,.2f}, +total={t:,.2f}')
for g, cnt in zz_counts.items():
    ratio = cnt / zz_man
    p = round(zz_total_nr_p * ratio, 2)
    t = round(zz_total_nr_t * ratio, 2)
    zz_alloc[g] = {'p': p, 't': t, 'ratio': ratio}
    print(f'  {g}: ratio={ratio:.4f}, +principal={p:,.2f}, +total={t:,.2f}')

# Read STAGE1 data
STAGE1_PATH = os.path.join(DATA_DIR, '一阶段小组信息.xlsx')
s1_df = pd.read_excel(STAGE1_PATH)
s1_data = {}
for _, row in s1_df.iterrows():
    g = str(row.iloc[0]).strip()
    s1_data[g] = {
        'principal': float(row.iloc[1]) if not pd.isna(row.iloc[1]) else 0,
        'total': float(row.iloc[2]) if not pd.isna(row.iloc[2]) else 0,
    }

# Read GROUP_SUMMARY (大表) data - filtered
VALID_GROUPS = ['ZZ-F1','ZZ-F2','ZZ-F3','ZZ-F4','ZZ-F5','CD-F1','CD-F2']
df_filtered = df[df.iloc[:, 2].astype(str).str.strip().isin(VALID_GROUPS)]

gs_data = {}
for _, r in df_filtered.iterrows():
    g = str(r['组别']).strip()
    if g not in gs_data:
        gs_data[g] = {'principal': 0, 'total': 0}
    gs_data[g]['principal'] += float(r['分案剩余本金']) if not pd.isna(r['分案剩余本金']) else 0
    gs_data[g]['total'] += float(r['累计回退']) if not pd.isna(r['累计回退']) else 0

print(f'\n=== 最终合并数值（用于排名计算） ===')
print(f'{"组别":<8} {"大表本金":>14} {"大表回退":>14} {"+自然回收P":>12} {"+自然回收T":>12} {"一阶段本金":>14} {"一阶段回退":>14} {"合并本金":>14} {"合并回退":>14} {"回退率":>8}')
print('-' * 130)

cd_final_p = 0; cd_final_t = 0
zz_final_p = 0; zz_final_t = 0

for g in ['CD-F1','CD-F2','ZZ-F1','ZZ-F2','ZZ-F3','ZZ-F4','ZZ-F5']:
    gs_p = gs_data.get(g, {}).get('principal', 0)
    gs_t = gs_data.get(g, {}).get('total', 0)
    nr_p = cd_alloc[g]['p'] if g in cd_alloc else zz_alloc[g]['p']
    nr_t = cd_alloc[g]['t'] if g in cd_alloc else zz_alloc[g]['t']
    # Note: gs_data already available, but natural recovery was added in Python script
    # Here we show the allocation separately for clarity
    s1_p = s1_data.get(g, {}).get('principal', 0)
    s1_t = s1_data.get(g, {}).get('total', 0)

    merged_p = gs_p + nr_p + s1_p
    merged_t = gs_t + nr_t + s1_t
    rate = merged_t / merged_p * 100 if merged_p > 0 else 0

    print(f'{g:<8} {gs_p:>14,.2f} {gs_t:>14,.2f} {nr_p:>12,.2f} {nr_t:>12,.2f} {s1_p:>14,.2f} {s1_t:>14,.2f} {merged_p:>14,.2f} {merged_t:>14,.2f} {rate:>7.2f}%')

    if g.startswith('CD'):
        cd_final_p += merged_p
        cd_final_t += merged_t
    else:
        zz_final_p += merged_p
        zz_final_t += merged_t

print(f'\n成都合计: 本金={cd_final_p:,.2f}, 回退={cd_final_t:,.2f}, 回退率={cd_final_t/cd_final_p*100:.2f}%')
print(f'郑州合计: 本金={zz_final_p:,.2f}, 回退={zz_final_t:,.2f}, 回退率={zz_final_t/zz_final_p*100:.2f}%')
print(f'全量合计: 本金={cd_final_p+zz_final_p:,.2f}, 回退={cd_final_t+zz_final_t:,.2f}')
