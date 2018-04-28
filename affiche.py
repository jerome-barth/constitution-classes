from IPython.display import HTML, display
from config import *

# Affichage des couleurs dans le notebook
html = '<div style="overflow: auto;">'
templ = '<div style="padding: 5pt; float: left; color: black; background-color: {};">' + \
        '<strong>Couleur classe {}</strong></div>'
for i, (coul1, coul2) in enumerate(C_CLS):
    html += templ.format(coul1, i + 1 if i + 1 != len(C_CLS) else 'NA')
    html += templ.format(coul2, i + 1 if i + 1 != len(C_CLS) else 'NA')
    if not (i + 1) % 3: html += '</div><br><div style="overflow: auto;">'
html += '</div>'
html += '<br><div style="overflow: auto;">'
templ = '<div style="padding: 5pt; float: left; color: {}; background-color: {};">' + \
        '<strong>{}</strong></div>'
for cat in C_CAT:
    coultxt, coulfnd = C_CAT[cat]
    html += templ.format(coultxt, coulfnd, str(cat))
html += '</div>'
display(HTML(html))