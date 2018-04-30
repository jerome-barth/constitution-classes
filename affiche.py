import builtins
from IPython.lib import deepreload
builtins.reload = deepreload.reload


import config
reload(config, exclude=['math', 'datetime', 'time', 'collections'])

from IPython.display import HTML, display
from config import *

# Affichage des couleurs dans le notebook
html = '<div style="overflow: auto;">'
templ = '<div style="padding: 5pt; float: left; color: black; background-color: {};">' + \
        '<strong>Couleur {}</strong></div>'
for i, (coul1, coul2) in enumerate(C_CLS):
    if i < NB_DIVS: n_cls = 'classe '+str(NOM_DIVS[i])
    elif NB_DIVS<=i<len(C_CLS)-1: n_cls ='supplémentaire'
    else: n_cls = 'élèves NA'
    html += templ.format(coul1, n_cls)
    html += templ.format(coul2, n_cls)
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