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

# Exhibit 3: deposits shifted forward four quarters vs. net yield (the post-COVID lead)
k2 = wb['KPI 2 Booking Build']
nyb = wb['Net Yield Build']
labels = [k2[f'I{r}'].value for r in range(70, 84)]                  # Mar-24 .. Jun-27
dep_shift = [None] * 4 + [k2[f'C{r}'].value * 100 for r in range(70, 80)]   # plotted 4 quarters later
ny_rep = [k2[f'J{r}'].value * 100 for r in range(70, 80)]            # reported Mar-24 .. Jun-26
ny_guide = [ny_rep[-1], k2['J80'].value * 100, k2['J81'].value * 100]
ours = [ny_guide[-1], nyb['N58'].value * 100, nyb['O58'].value * 100]
cons = [ny_guide[-1], nyb['N31'].value * 100, nyb['O31'].value * 100]
fig, ax = plt.subplots(figsize=(W, 1.95))
fig.subplots_adjust(left=0.1, right=0.97, top=0.76, bottom=0.25)
xs = list(range(14))
ax.plot([x for x, v in zip(xs, dep_shift) if v is not None], [v for v in dep_shift if v is not None], color=INK, lw=1.6,
        marker='o', ms=2.6, label='Deposits, 4 quarters earlier')
ax.plot(xs[4:10], ny_rep[4:], color=MID, lw=1.6, marker='o', ms=2.6, label='Net yield')
ax.plot([9, 10, 11], ny_guide, color=MID, lw=1.4, ls='--')
ax.plot([11, 12, 13], ours, color=MID, lw=1.2, ls='--', marker='o', ms=2.6, mfc='white')
ax.plot([11, 12, 13], cons, color=LIGHT, lw=1.2, ls=':', marker='o', ms=2.6, mfc='white')
ax.text(13.15, ours[-1], 'Ours', va='center', fontsize=6, color='#404040')
ax.text(13.15, cons[-1], 'Consensus', va='center', fontsize=6, color='#404040')
ax.axhline(0, color='#808080', lw=0.6)
ax.set_xticks([4, 6, 8, 10, 12], [labels[i] for i in (4, 6, 8, 10, 12)])
ax.set_xlim(3.6, 15.6)
ax.set_ylim(-11, 15)
ax.legend(frameon=False, fontsize=6.3, loc='lower left', bbox_to_anchor=(0.0, 1.0), ncol=2, handlelength=1.4, columnspacing=0.8)
finish(fig, ax, 'Exhibit 3: Deposits vs. net yield (% y/y)',
       'Deposits per future berth-day shifted 4 quarters. Dashed: 3Q-4Q26\nguide, then 1H27 model. Source: SEC XBRL; 8-Ks; team model.', 'ex3_deposits')

# Exhibit 4: 2027 EBITDA bridge, consensus case at run-rate costs (Street EBITDA) to ours
mc = wb['Model vs Consensus']
start = mc['C26'].value
steps = [mc[f'C{r}'].value for r in range(27, 32)]
end = mc['C32'].value
names = ['Consensus\n(run-rate)', '4Q26\nexit', '1H27\nprice', '2H27\nprice', 'Occupancy', 'Onboard\nmix', 'Ours']
cols = [MID, INK, INK, DARK, DARK, '#8c8c8c', INK]
fig, ax = plt.subplots(figsize=(3.3, 2.05))
fig.subplots_adjust(left=0.1, right=0.99, top=0.88, bottom=0.27)
level = start
ax.bar(0, start, color=MID, width=0.62)
ax.text(0, start + 6, f'{start:,.0f}', ha='center', va='bottom', fontsize=6.3)
for i, st_ in enumerate(steps, start=1):
    ax.bar(i, st_, bottom=level, color=cols[i], width=0.62)
    ax.text(i, level + 6, f'{st_:,.0f}', ha='center', va='bottom', fontsize=6.3)
    level += st_
ax.bar(6, end, color=INK, width=0.62)
ax.text(6, end + 6, f'{end:,.0f}', ha='center', va='bottom', fontsize=6.3)
ax.set_ylim(2350, 2720)
ax.set_xticks(range(7), names, fontsize=6.2)
ax.tick_params(axis='x', length=0)
finish(fig, ax, 'Exhibit 4: 2027 EBITDA, consensus case to ours ($M)',
       'Black: fare panel. Dark gray: deposits. Gray: guest mix (both). Unit costs are\nrun-rate in both; consensus EPS also needs them to fall (another $107M).', 'ex4_bridge')
print('charts written')
