"""Memo exhibits for v13, drawn from NCLH_model_v6.xlsx (recalculated values).
Plain descriptive titles; black = our view / NCL, gray = comparison."""
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
import openpyxl

model, outdir = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(model, data_only=True)
for f in font_manager.findSystemFonts():
    if 'Carlito' in f:
        font_manager.fontManager.addfont(f)
plt.rcParams.update({'font.family': 'Carlito', 'font.size': 7.5, 'axes.linewidth': 0.6, 'axes.edgecolor': '#808080',
                     'xtick.color': '#404040', 'ytick.color': '#404040', 'xtick.major.width': 0.6, 'ytick.major.width': 0.6,
                     'axes.labelcolor': '#404040'})
INK, DARK, MID, LIGHT = '#000000', '#404040', '#8c8c8c', '#c8c8c8'
W = 2.5                                    # inches, as placed in the memo


def finish(fig, ax, title, source, name):
    ax.spines[['top', 'right']].set_visible(False)
    fig.text(0.01, 0.985, title, ha='left', va='top', fontsize=8, fontweight='bold')
    fig.text(0.01, 0.01, source, ha='left', va='bottom', fontsize=5.8, color='#595959')
    fig.savefig(f'{outdir}/{name}.png', dpi=300)
    plt.close(fig)


v = lambda sh, c: wb[sh][c].value

# Exhibit 1: EBITDA the price needs vs. Street vs. ours
ex = wb['Exhibit Data']
labels = [ex[f'B{r}'].value for r in range(44, 50)]
vals = [ex[f'C{r}'].value / 1000 for r in range(44, 50)]
labels = ['2025 actual', 'March 2026 guide', 'Share price needs (9%)', 'Average target needs', 'Street 2027', 'Ours 2027']
fig, ax = plt.subplots(figsize=(W, 1.9))
fig.subplots_adjust(left=0.43, right=0.9, top=0.86, bottom=0.17)
y = range(len(vals))[::-1]
cols = [LIGHT, LIGHT, MID, MID, MID, INK]
ax.barh(list(y), vals, color=cols, height=0.62)
for yi, val in zip(y, vals):
    ax.text(val + 0.02, yi, f'${val:.2f}B', va='center', fontsize=7)
ax.set_yticks(list(y), labels)
ax.set_xlim(2.0, 3.25)
ax.set_xticks([2.0, 2.5, 3.0])
ax.tick_params(axis='y', length=0)
finish(fig, ax, 'Exhibit 1: Adjusted EBITDA ($B)', 'Source: NCLH 8-Ks; Yahoo/Zacks; team model (reverse DCF at 9% WACC).', 'ex1_ebitda')

# Exhibit 2: fare change by line and region, 1H27 sailings
fp = wb['Fare Panel']
def cell(line, region, window='1H 2027'):
    for r in range(8, 115):
        if fp[f'B{r}'].value == line and fp[f'C{r}'].value == region and fp[f'D{r}'].value == window:
            return fp[f'F{r}'].value * 100
    return 0
lines = [('NCL', 'Norwegian'), ('Royal', 'Royal Caribbean'), ('Carnival', 'Carnival'), ('Celebrity', 'Celebrity')]
carib = [cell(l, 'Caribbean/Bahamas') for _, l in lines]
europe = [cell(l, 'Europe') for _, l in lines]
fig, ax = plt.subplots(figsize=(W, 1.9))
fig.subplots_adjust(left=0.1, right=0.97, top=0.78, bottom=0.2)
x = range(len(lines))
ax.bar([i - 0.19 for i in x], carib, width=0.36, color=INK, label='Caribbean')
ax.bar([i + 0.19 for i in x], europe, width=0.36, color=LIGHT, label='Europe')
for i, (c, e) in enumerate(zip(carib, europe)):
    ax.text(i - 0.19, c - 0.6, f'{c:.1f}' if c else '0', ha='center', va='top', fontsize=6.5)
    ax.text(i + 0.19, e - 0.6, f'{e:.1f}' if e else '0', ha='center', va='top', fontsize=6.5)
ax.axhline(0, color='#808080', lw=0.6)
ax.set_xticks(list(x), [n for n, _ in lines])
ax.set_ylim(-18, 1)
ax.tick_params(axis='x', length=0)
ax.spines['bottom'].set_visible(False)
ax.legend(frameon=False, fontsize=7, loc='lower right', bbox_to_anchor=(1.0, 1.0), ncol=2, handlelength=1)
finish(fig, ax, 'Exhibit 2: Fare change, 1H27 sailings (%)',
       'Median lowest fare vs. its own 90-day average. Source: allaboarddeals.com,\n4,783 itineraries, Oct 2, 2026.', 'ex2_fares')

# Exhibit 3: deposit indicator vs. net yield
qs = [ex[f'B{r}'].value for r in range(17, 26)]
dep = [ex[f'C{r}'].value for r in range(17, 26)]
ny = [ex[f'D{r}'].value for r in range(17, 26)]
qs = ['Dec-24', 'Mar-25', 'Jun-25', 'Sep-25', 'Dec-25', 'Mar-26', 'Jun-26', 'Sep-26', 'Dec-26']
fig, ax = plt.subplots(figsize=(W, 1.9))
fig.subplots_adjust(left=0.1, right=0.97, top=0.8, bottom=0.2)
xd = [i for i, d in enumerate(dep) if d is not None]
ax.plot(xd, [dep[i] * 100 for i in xd], color=INK, lw=1.6, marker='o', ms=3, label='Deposits per future berth-day')
ax.plot(range(7), [n * 100 for n in ny[:7]], color=MID, lw=1.6, marker='o', ms=3, label='Net yield')
ax.plot([6, 7, 8], [ny[6] * 100, ny[7] * 100, ny[8] * 100], color=MID, lw=1.4, ls='--')
ax.axhline(0, color='#808080', lw=0.6)
ax.set_xticks([0, 2, 4, 6, 8], [qs[i] for i in (0, 2, 4, 6, 8)])
ax.set_ylim(-11, 11)
ax.annotate(f'{dep[6] * 100:.1f}%', (6, dep[6] * 100), xytext=(4, -9), textcoords='offset points', fontsize=6.5)
ax.legend(frameon=False, fontsize=6.5, loc='lower left', bbox_to_anchor=(0.0, 1.0), ncol=2, handlelength=1.4, columnspacing=0.8)
finish(fig, ax, 'Exhibit 3: NCLH deposits vs. net yield (% y/y)',
       'Dashed: 3Q26 guide and 4Q26 implied. Source: SEC XBRL; NCLH 8-Ks.', 'ex3_deposits')

# Exhibit 4: 2027 EBITDA bridge, consensus case to ours
mc = wb['Model vs Consensus']
start = mc['C24'].value
steps = [mc[f'C{r}'].value for r in range(25, 31)]
end = mc['C31'].value
names = ['Consensus\ncase', '2H26\nexit', 'Thesis 1\nprice', 'Thesis 2\nprice', 'Thesis 2\noccupancy', 'Onboard\nspend', 'Costs\nand fuel', 'Ours']
kinds = ['total', 'other', 'thesis', 'thesis', 'thesis', 'other', 'other', 'total']
fig, ax = plt.subplots(figsize=(3.3, 2.05))
fig.subplots_adjust(left=0.1, right=0.99, top=0.88, bottom=0.25)
level = start
ax.bar(0, start, color=MID, width=0.62)
ax.text(0, start + 8, f'{start:,.0f}', ha='center', va='bottom', fontsize=6.3)
for i, s in enumerate(steps, start=1):
    color = INK if kinds[i] == 'thesis' else LIGHT
    ax.bar(i, s, bottom=level, color=color, width=0.62)
    ax.text(i, level + 8, f'{s:,.0f}', ha='center', va='bottom', fontsize=6.3)
    level += s
ax.bar(7, end, color=INK, width=0.62)
ax.text(7, end + 8, f'{end:,.0f}', ha='center', va='bottom', fontsize=6.3)
ax.set_ylim(2300, 2850)
ax.set_xticks(range(8), names, fontsize=6.3)
ax.tick_params(axis='x', length=0)
finish(fig, ax, 'Exhibit 4: 2027 EBITDA, consensus case to ours ($M)',
       'Black: the two theses. Light gray: other differences. Source: team model (Bridge Build).', 'ex4_bridge')
print('charts written')
