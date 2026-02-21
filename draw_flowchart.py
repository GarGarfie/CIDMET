"""Generate data flow diagram for CIDMET software copyright listing."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import os

# --- Configuration ---
fig, ax = plt.subplots(1, 1, figsize=(10, 14.5))
ax.set_xlim(0, 10)
ax.set_ylim(0, 14.5)
ax.axis('off')
fig.patch.set_facecolor('white')

# Font settings
font_props = {'fontsize': 9, 'fontfamily': 'DejaVu Sans'}
font_title = {'fontsize': 10, 'fontfamily': 'DejaVu Sans', 'fontweight': 'bold'}
font_small = {'fontsize': 7.5, 'fontfamily': 'DejaVu Sans', 'color': '#555555'}


def draw_box(x, y, w, h, text, color='#E3F2FD', edge='#1565C0',
             subtext=None, bold=False):
    """Draw a rounded rectangle with centered text."""
    box = FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                         boxstyle="round,pad=0.1",
                         facecolor=color, edgecolor=edge, linewidth=1.5)
    ax.add_patch(box)
    fp = font_title if bold else font_props
    ax.text(x, y + (0.12 if subtext else 0), text,
            ha='center', va='center', **fp)
    if subtext:
        ax.text(x, y - 0.18, subtext,
                ha='center', va='center', **font_small)


def arrow(x1, y1, x2, y2, label=None, color='#424242', lw=1.5, ls='-'):
    """Draw an arrow with optional label."""
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=lw,
                                linestyle=ls))
    if label:
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        ax.text(mx + 0.15, my, label, ha='left', va='center', **font_small)


# =====================================================================
# Title
# =====================================================================
ax.text(5, 14.1, 'ПОТОК ДАННЫХ ПРОГРАММЫ CIDMET',
        ha='center', va='center', fontsize=13,
        fontfamily='DejaVu Sans', fontweight='bold', color='#1565C0')
ax.text(5, 13.75, 'Схема взаимодействия модулей',
        ha='center', va='center', fontsize=10,
        fontfamily='DejaVu Sans', color='#555555')

# =====================================================================
# Layer 1 — User
# =====================================================================
draw_box(5, 13.1, 2.4, 0.5, 'Пользователь',
         color='#FFF3E0', edge='#E65100', bold=True)
arrow(5, 12.84, 5, 12.46)

# =====================================================================
# Layer 2 — main.py
# =====================================================================
draw_box(5, 12.2, 2.0, 0.46, 'main.py',
         color='#F3E5F5', edge='#7B1FA2', subtext='точка входа')
arrow(5, 11.96, 5, 11.56)

# =====================================================================
# Layer 3 — gui_app.py
# =====================================================================
draw_box(5, 11.3, 3.4, 0.46, 'gui_app.py',
         color='#E8F5E9', edge='#2E7D32',
         subtext='GUI: MainWindow, ProcessingWorker, DuplicateDialog')

# Arrows to input files
arrow(3.5, 11.06, 2.2, 10.56)
arrow(5, 11.06, 5, 10.56)
arrow(6.5, 11.06, 7.8, 10.56)

# =====================================================================
# Layer 4 — Input files
# =====================================================================
draw_box(2.2, 10.3, 2.2, 0.46, 'BibTeX (.bib)',
         color='#FFFDE7', edge='#F9A825')
draw_box(5, 10.3, 2.2, 0.46, 'WoS TXT / XLS',
         color='#FFFDE7', edge='#F9A825')
draw_box(7.8, 10.3, 2.2, 0.46, 'Scopus / EI',
         color='#FFFDE7', edge='#F9A825', subtext='CSV, TXT')

# Arrows to parsers
arrow(2.2, 10.06, 3.8, 9.56)
arrow(5, 10.06, 5, 9.56)
arrow(7.8, 10.06, 6.2, 9.56)

# =====================================================================
# Layer 5 — parsers.py
# =====================================================================
draw_box(5, 9.2, 5.6, 0.65, 'parsers.py',
         color='#E3F2FD', edge='#1565C0', bold=True,
         subtext='parse_bibtex, parse_wos_txt/xls, parse_scopus_csv/txt, parse_ei_csv/txt')
arrow(5, 8.87, 5, 8.33, label='ParsedRecord[ ]')

# =====================================================================
# Layer 6 — matcher.py
# =====================================================================
draw_box(5, 8.0, 5.6, 0.6, 'matcher.py',
         color='#E3F2FD', edge='#1565C0', bold=True,
         subtext='MatchEngine: 1) DOI exact  →  2) Title exact  →  3) Fuzzy + author')
arrow(5, 7.69, 5, 7.15, label='MatchResult[ ]')

# =====================================================================
# Layer 7 — writers.py & gui_app results (split)
# =====================================================================
draw_box(3.2, 6.7, 3.4, 0.65, 'writers.py',
         color='#E3F2FD', edge='#1565C0', bold=True,
         subtext='экспорт подмножеств + объединённый экспорт')
draw_box(7.5, 6.7, 2.8, 0.65, 'gui_app.py',
         color='#E8F5E9', edge='#2E7D32',
         subtext='статистика, дубликаты')

arrow(4.2, 7.15, 3.2, 7.03)
arrow(5.8, 7.15, 7.5, 7.03)

# =====================================================================
# Layer 8 — Output formats (three boxes aligned horizontally)
# =====================================================================
out_y = 5.6
draw_box(1.8, out_y, 2.0, 0.46, 'WoS формат',
         color='#FBE9E7', edge='#D84315', subtext='TXT / XLS')
draw_box(3.5, out_y, 2.0, 0.46, 'Scopus формат',
         color='#FBE9E7', edge='#D84315', subtext='CSV / TXT')
draw_box(5.2, out_y, 2.0, 0.46, 'EI формат',
         color='#FBE9E7', edge='#D84315', subtext='CSV / TXT')

arrow(2.0, 6.37, 1.8, out_y + 0.24)
arrow(3.2, 6.37, 3.5, out_y + 0.24)
arrow(4.4, 6.37, 5.2, out_y + 0.24)

# =====================================================================
# Layer 9 — Merged export
# =====================================================================
merge_y = 4.6
draw_box(3.5, merge_y, 4.8, 0.5, 'Объединённый экспорт',
         color='#FCE4EC', edge='#C62828', bold=True,
         subtext='дедупликация + конвертация форматов авторов')

arrow(1.8, out_y - 0.24, 2.5, merge_y + 0.26)
arrow(3.5, out_y - 0.24, 3.5, merge_y + 0.26)
arrow(5.2, out_y - 0.24, 4.5, merge_y + 0.26)

# =====================================================================
# utils.py — side module (right side, spans layers 5–8)
# =====================================================================
ux, uy = 8.5, 5.6
uw, uh = 2.4, 2.0
draw_box(ux, uy, uw, uh, '', color='#ECEFF1', edge='#546E7A')
ax.text(ux, uy + 0.75, 'utils.py', ha='center', va='center', **font_title)

util_funcs = [
    'normalize_doi()',
    'normalize_title()',
    'detect_encoding()',
    'extract_first_author()',
    'detect_csv_dialect()',
    'detect_scopus_language()',
]
for i, func in enumerate(util_funcs):
    ax.text(ux, uy + 0.45 - i * 0.23, func, ha='center', va='center',
            fontsize=7, fontfamily='DejaVu Sans', color='#37474F')

# Dashed arrows from utils to processing modules
arrow(7.3, uy + 0.75, 7.05, 9.2, color='#90A4AE', lw=1.0, ls='dashed')
arrow(7.3, uy + 0.55, 7.05, 8.0, color='#90A4AE', lw=1.0, ls='dashed')
arrow(7.3, uy + 0.35, 4.9, 6.7, color='#90A4AE', lw=1.0, ls='dashed')

# =====================================================================
# Legend
# =====================================================================
leg_y = 3.5
ax.plot([0.5, 9.5], [leg_y + 0.45, leg_y + 0.45], color='#BDBDBD',
        lw=0.5, ls='-')
ax.text(5, leg_y + 0.25, 'Условные обозначения',
        ha='center', va='center', fontsize=9,
        fontfamily='DejaVu Sans', fontweight='bold')

legend_items = [
    (1.2, leg_y - 0.15, '#FFF3E0', '#E65100', 'Пользователь'),
    (3.1, leg_y - 0.15, '#F3E5F5', '#7B1FA2', 'Точка входа'),
    (5.0, leg_y - 0.15, '#E8F5E9', '#2E7D32', 'GUI интерфейс'),
    (6.9, leg_y - 0.15, '#FFFDE7', '#F9A825', 'Входные файлы'),
    (8.8, leg_y - 0.15, '#E3F2FD', '#1565C0', 'Модули'),
]
legend_items2 = [
    (1.8, leg_y - 0.65, '#FBE9E7', '#D84315', 'Выходные файлы'),
    (4.0, leg_y - 0.65, '#FCE4EC', '#C62828', 'Объединённый вывод'),
    (6.4, leg_y - 0.65, '#ECEFF1', '#546E7A', 'Утилиты'),
]

for lx, ly, fc, ec, label in legend_items + legend_items2:
    box = FancyBboxPatch((lx - 0.55, ly - 0.15), 1.1, 0.3,
                         boxstyle="round,pad=0.05",
                         facecolor=fc, edgecolor=ec, linewidth=1)
    ax.add_patch(box)
    ax.text(lx, ly, label, ha='center', va='center',
            fontsize=7, fontfamily='DejaVu Sans')

# Arrow legend
arr_y = leg_y - 1.1
ax.annotate('', xy=(2.3, arr_y), xytext=(1.3, arr_y),
            arrowprops=dict(arrowstyle='->', color='#424242', lw=1.5))
ax.text(2.5, arr_y, '— передача данных', ha='left', va='center',
        fontsize=7.5, fontfamily='DejaVu Sans')

ax.annotate('', xy=(6.3, arr_y), xytext=(5.3, arr_y),
            arrowprops=dict(arrowstyle='->', color='#90A4AE', lw=1.0,
                            linestyle='dashed'))
ax.text(6.5, arr_y, '— вызов утилит', ha='left', va='center',
        fontsize=7.5, fontfamily='DejaVu Sans')

# =====================================================================
# Save
# =====================================================================
output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           'CIDMET_dataflow.png')
plt.tight_layout()
plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
print(f'Saved: {output_path}')
