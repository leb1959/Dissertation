#!/usr/bin/env python3
"""
run_tukey.py — Аналіз викидів методом меж Тюкі / Tukey fences (Додаток 4.7)
Відтворює Таблицю 4.13 з тексту розділу 4.

Запуск:  python run_tukey.py
Потреби: pip install numpy pandas openpyxl
Файл:    Додаток_4_3.xlsx  (у тій самій папці)
"""
import numpy as np
import pandas as pd

# ── 1. Завантаження ───────────────────────────────────────────────────────
df = pd.read_excel('Додаток_4_3.xlsx', header=2)
df.columns = ['num','date','strok','month','TSP','PM10','ratio',
              'n_cairnet','RH','T'] + list(df.columns[10:])
df = df[pd.to_numeric(df['num'], errors='coerce').notna()].copy()
df = df.dropna(subset=['TSP','PM10'])
df = df[df['TSP'] > 0]   # TSP=0 фізично неможливий — виключаємо

month_map = {'Січень':1,'Лютий':2,'Березень':3,'Квітень':4,'Травень':5,
             'Червень':6,'Липень':7,'Серпень':8,'Вересень':9,
             'Жовтень':10,'Листопад':11,'Грудень':12}
if df['month'].dtype == object:
    df['mnum'] = df['month'].map(month_map)
else:
    df['mnum'] = df['month'].astype(int)
month_names = {v:k for k,v in month_map.items()}
print(f"Пар для аналізу: {len(df)}\n")

# ── 2. Глобальні межі Тюкі (на всій вибірці) ─────────────────────────────
def tukey_bounds(series):
    q1, q3 = series.quantile(0.25), series.quantile(0.75)
    iqr = q3 - q1
    return q1 - 1.5*iqr, q3 + 1.5*iqr

lo_tsp, hi_tsp = tukey_bounds(df['TSP'])
lo_pm,  hi_pm  = tukey_bounds(df['PM10'])

df['out_tsp'] = (df['TSP'] < lo_tsp) | (df['TSP'] > hi_tsp)
df['out_pm']  = (df['PM10'] < lo_pm)  | (df['PM10'] > hi_pm)

print(f"Межі Тюкі TSP  : [{lo_tsp:.0f} ; {hi_tsp:.0f}] мкг/м³")
print(f"Межі Тюкі PM₁₀ : [{lo_pm:.1f} ; {hi_pm:.1f}] мкг/м³\n")

# ── 3. Місячна статистика ─────────────────────────────────────────────────
print(f"{'Місяць':12} {'N':>4}  {'TSP вик.':>9}  {'TSP %':>7}  {'PM10 вик.':>10}  {'PM10 %':>7}")
print("─"*62)

results = []
for m in range(1,13):
    sub = df[df['mnum']==m]
    if len(sub)==0: continue
    nt, np10 = sub['out_tsp'].sum(), sub['out_pm'].sum()
    pt = 100*nt/len(sub);  pp = 100*np10/len(sub)
    pm_str = f"{np10} ({pp:.1f}%)" if np10>0 else "—"
    print(f"{month_names[m]:12} {len(sub):>4}   {nt:>4} ({pt:.1f}%)   {pm_str:>15}")
    results.append({'Місяць':month_names[m],'N пар':len(sub),
                    'Викидів TSP n':int(nt),'Викидів TSP %':round(pt,1),
                    'Викидів PM10 n':int(np10),'Викидів PM10 %':round(pp,1),
                    'Залишено у вибірці':'Так'})

# Загальний підсумок
print("─"*62)
nt_all = df['out_tsp'].sum(); np_all = df['out_pm'].sum(); N = len(df)
print(f"{'ЗАГАЛОМ':12} {N:>4}   {nt_all:>4} ({100*nt_all/N:.1f}%)  "
      f"{np_all:>4} ({100*np_all/N:.1f}%)")

print(f"\nВисновок: виявлені 'статистичні викиди' є реальними фізичними подіями")
print(f"(промислові емісії, ресуспензія). Всі збережено у вибірці.")

# ── 4. Збереження ─────────────────────────────────────────────────────────
pd.DataFrame(results).to_excel('tukey_results.xlsx', index=False)
print(f"\nЗбережено: tukey_results.xlsx")
