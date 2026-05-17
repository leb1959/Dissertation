# =============================================================================
# gen_graphs_chapter5.py  —  Генерація 36 графіків для Розділу 5
# =============================================================================
# Використання (запускати по 3-4 місяці за раз через обмеження часу):
#   python gen_graphs_chapter5.py 1  3    # місяці 1-3
#   python gen_graphs_chapter5.py 4  6    # місяці 4-6
#   python gen_graphs_chapter5.py 7  9    # місяці 7-9
#   python gen_graphs_chapter5.py 10 12   # місяці 10-12
#
# Результат: PNG-файли зберігаються у IMG_DIR (змінна нижче)
# Після генерації — запустити окремий скрипт вбудовування в docx.
#
# Залежності: pip install pandas openpyxl matplotlib
# =============================================================================

import os, io, sys
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import AutoMinorLocator
import numpy as np

DATA_BASE = "/sessions/admiring-sharp-ramanujan/mnt/Розіл4/Додаток_2 Оцифровані дані/Processed_CairCloud_DIG"
IMG_DIR   = "/sessions/admiring-sharp-ramanujan/mnt/outputs/graphs"
os.makedirs(IMG_DIR, exist_ok=True)

plt.rcParams.update({
    'font.family': 'DejaVu Sans', 'font.size': 9, 'axes.labelsize': 9,
    'xtick.labelsize': 8, 'ytick.labelsize': 8, 'legend.fontsize': 8,
    'axes.grid': True, 'grid.alpha': 0.3, 'grid.linestyle': '--',
    'axes.spines.top': False,
})

COL = {'pan2': '#1f77b4', 'pan5': '#2ca02c', 'pan6': '#d62728'}

def load_month(post, m):
    f = os.path.join(DATA_BASE, f"CairCloud_DIG_{post}", f"{m:02d}_CairCloud_DIG_{post}.xlsx")
    df = pd.read_excel(f, usecols=[0,1,2,3,4,5,6], engine='openpyxl')
    df.columns = ['ts','pm','qpm','temp','qt','rh','qrh']
    df['pm']   = df['pm'].where(df['qpm']==1, np.nan)
    df['temp'] = df['temp'].where(df['qt']==1, np.nan)
    df['rh']   = df['rh'].where(df['qrh']==1, np.nan)
    df['ts']   = pd.to_datetime(df['ts'])
    return df

FIG_IMG = {
    1:'image17.png', 2:'image2.png',  3:'image5.png',  4:'image26.png',
    5:'image11.png', 6:'image27.png', 7:'image22.png', 8:'image25.png',
    9:'image24.png',10:'image36.png',11:'image21.png',12:'image31.png',
   13:'image23.png',14:'image34.png',15:'image29.png',16:'image28.png',
   17:'image35.png',18:'image32.png',19:'image30.png',20:'image33.png',
   21:'image20.png',22:'image4.png', 23:'image3.png', 24:'image7.png',
   25:'image9.png', 26:'image16.png',27:'image10.png',28:'image13.png',
   29:'image18.png',30:'image19.png',31:'image8.png', 32:'image6.png',
   33:'image14.png',34:'image1.png', 35:'image12.png',36:'image15.png',
}

def fmt_ax(ax):
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d.%m'))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=0))
    ax.xaxis.set_minor_locator(mdates.DayLocator())
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha='center')
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))

month_range = range(int(sys.argv[1]), int(sys.argv[2])+1)

for m in month_range:
    print(f"Month {m}...", flush=True)
    data = {p: load_month(p, m) for p in ('pan2','pan5','pan6')}

    # --- Type 1: PM10 ---
    fig, ax = plt.subplots(figsize=(13,4))
    for post in ('pan2','pan6','pan5'):
        d = data[post]
        ax.plot(d['ts'], d['pm'], color=COL[post], lw=0.8, alpha=0.85, label=post)
    ax.set_ylabel('PM₁₀, мкг/м³')
    fmt_ax(ax)
    ax.legend(loc='upper right', framealpha=0.7)
    fig.tight_layout(pad=0.5)
    out = f"{IMG_DIR}/{FIG_IMG[m]}"
    fig.savefig(out, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  -> {out}", flush=True)

    # --- Type 2: PM10 + RH ---
    fig, ax1 = plt.subplots(figsize=(13,4.5))
    for post in ('pan2','pan6','pan5'):
        d = data[post]
        ax1.plot(d['ts'], d['pm'], color=COL[post], lw=0.8, alpha=0.85, label=f'PM₁₀ {post}')
    ax1.set_ylabel('PM₁₀, мкг/м³')
    ax2 = ax1.twinx()
    ax2.plot(data['pan2']['ts'], data['pan2']['rh'], color='#7f7f7f', lw=1.0, alpha=0.6, ls='--', label='RH, %')
    ax2.set_ylabel('RH, %', color='#555')
    ax2.tick_params(axis='y', labelcolor='#555')
    ax2.set_ylim(0,105)
    fmt_ax(ax1)
    l1,lb1 = ax1.get_legend_handles_labels()
    l2,lb2 = ax2.get_legend_handles_labels()
    ax1.legend(l1+l2, lb1+lb2, loc='upper right', framealpha=0.7, ncol=2)
    fig.tight_layout(pad=0.5)
    out = f"{IMG_DIR}/{FIG_IMG[m+12]}"
    fig.savefig(out, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  -> {out}", flush=True)

    # --- Type 3: PM10 + Temp ---
    fig, ax1 = plt.subplots(figsize=(13,4.5))
    for post in ('pan2','pan6','pan5'):
        d = data[post]
        ax1.plot(d['ts'], d['pm'], color=COL[post], lw=0.8, alpha=0.85, label=f'PM₁₀ {post}')
    ax1.set_ylabel('PM₁₀, мкг/м³')
    ax2 = ax1.twinx()
    ax2.plot(data['pan2']['ts'], data['pan2']['temp'], color='#ff7f0e', lw=1.0, alpha=0.7, ls='--', label='T, °C')
    ax2.set_ylabel('T, °C', color='#b85000')
    ax2.tick_params(axis='y', labelcolor='#b85000')
    fmt_ax(ax1)
    l1,lb1 = ax1.get_legend_handles_labels()
    l2,lb2 = ax2.get_legend_handles_labels()
    ax1.legend(l1+l2, lb1+lb2, loc='upper right', framealpha=0.7, ncol=2)
    fig.tight_layout(pad=0.5)
    out = f"{IMG_DIR}/{FIG_IMG[m+24]}"
    fig.savefig(out, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  -> {out}", flush=True)

print("DONE", flush=True)
