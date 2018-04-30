import sys
ipython = get_ipython()


def hide_traceback(exc_tuple=None,
                   filename=None,
                   tb_offset=None,
                   exception_only=False,
                   running_compiled_code=False):
    etype, value, tb = sys.exc_info()
    return ipython._showtraceback(etype, value,
                                  ipython.InteractiveTB.get_exception_only(
                                      etype, value))


ipython.showtraceback = hide_traceback

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
    if i < NB_DIVS: n_cls = 'classe ' + str(NOM_DIVS[i])
    elif NB_DIVS <= i < len(C_CLS) - 1: n_cls = 'supplémentaire'
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
#display(HTML(html))

LV2S_VRAIES = LV2S[:-1]
OPTIONS_UNIQUES = [opt for opt in OPTIONS if len(OPTIONS[opt]) == 1]

from datetime import datetime
try:
    ANNEE
except NameError:
    ANNEE = datetime.today().year % 100  # automatique

try:
    if 'xlsm' != NOM_FICHIER.split('.')[-1]: NOM_FICHIER += '.xlsm'
except NameError:
    NOM_FICHIER = 'R' + str(ANNEE) + '-Répart-' + CLASSES + '.xlsm'

html2 = '''
<center>
 <h2>Récapitulatif des paramètres choisis</h2>
 <table>
  <tr>
   <th style="text-align:center">Paramètre</th>
   <th style="text-align:center">Valeur<th>
  </tr>
'''
templ = '''
  <tr>
   <td style="text-align:left">&nbsp;<strong>{}</strong>&nbsp;</td>
   <td style="text-align:center">&nbsp;{}&nbsp;</td>
  </tr>
'''
templ_span = '''
  <tr>
   <td style="text-align:center" colspan="2"><strong>{}</strong></td>
  </tr>
'''
templ_div = '''
  <tr>
   <td style="text-align:center"><div style="padding: 3px; background-color: {0};"><strong>''' + CLASSES + ''' {2}</strong></div></td>
   <td style="text-align:center"><div style="padding: 3px; background-color: {1};"><strong>''' + CLASSES + ''' {2}</strong></div></td>
  </tr>
'''
templ_niv = '<span style="padding: 3px; color: {}; background-color: {};"><strong>&nbsp;{}&nbsp;</strong></span>'
templ_coul = '''
  <tr>
   <td style="text-align:left">&nbsp;<strong>{}</strong>&nbsp;</td>
   <td style="text-align:center"><div style="padding: 3px; color: {}; background-color: {};"><strong>{}</strong></div></td>
  </tr>
'''
html2 += templ.format('Établissement :',
                      '<font size="4">' + ETABLISSEMENT + '</font>')
html2 += templ.format('Ville :', '<font size="3">' + VILLE + '</font>')
html2 += templ_coul.format('Classes :', C_CAT['CLS'][0], C_CAT['CLS'][1],
                           CLASSES)
html2 += templ.format(
    'Nombre de divisions :',
    '<font size="3"><strong>' + str(NB_DIVS) + '</strong></font>')
html2 += templ.format('Nom et couleurs des divisions :', '')
for i, div in enumerate(NOM_DIVS):
    html2 += templ_div.format(C_CLS[i][0], C_CLS[i][1], div)

for i, lv2 in enumerate(LV2S):
    ent = 'LV2 :' if i == 0 else ''
    t_lv = 'LV2'
    if i == len(LV2S) - 1: t_lv = 's' + t_lv
    html2 += templ_coul.format(ent, C_CAT[t_lv][0], C_CAT[t_lv][1], lv2)
c_op_un = {}
for i, opt in enumerate(OPTIONS_UNIQUES):
    ent = 'Options :' if i == 0 else ''
    c_op_un = {**c_op_un, opt: C_CAT['opt' + str(i % 3 + 1)]}
    html2 += templ_coul.format(ent, c_op_un[opt][0], c_op_un[opt][1], opt)

compat = []
for i, opt in enumerate(OPTIONS):
    if len(OPTIONS[opt]) > 1:
        compat.append(opt)
if compat:
    html2 += templ.format('Options compatibles :', '')
    for opt in compat:
        for i, opt2 in enumerate(OPTIONS[opt]):
            ent = opt if i == 0 else ''
            html2 += templ_coul.format(ent, c_op_un[opt2][0], c_op_un[opt2][1],
                                       opt2)
        if opt != compat[-1]: html2 += templ_span.format('* *')
    html2 += templ_span.format('* * * * *')
for i, niv in enumerate(NIVEAUX):
    ent = 'Niveaux :' if i == 0 else ''
    html2 += templ.format(ent,
                          templ_niv.format(C_CAT[niv][0], C_CAT[niv][1], niv))

html2 += templ_span.format('* * * * *')
p_ent = True
tx_coul = {
    'F': 'Filles',
    'G': 'Garçons',
    '%F': 'Pourcentage de filles',
    'R': 'Retard',
    'TOT': 'Colonne Total (clair)',
    'TOT2': 'Colonne Total (foncé)',
    'Reste1': 'Colonne Reste (clair)',
    'Reste2': 'Colonne Reste (foncé)',
    'ptR': 'Comportement ' + NIVEAUX[-3],
    'moyR': 'Comportement ' + NIVEAUX[-2],
    'grR': 'Comportement ' + NIVEAUX[-1],
    'ERR': 'Erreur',
    'ERRP': 'Erreur de structure',
}
for i, coul in enumerate(C_CAT):
    if p_ent:
        ent = 'Couleurs :'
        p_ent = False
    else:
        ent = ''
    if coul not in NIVEAUX + LV2S + [
            'LV2', 'sLV2', 'opt1', 'opt2', 'opt3', 'CLS'
    ]:
        html2 += templ_coul.format(ent, C_CAT[coul][0], C_CAT[coul][1],
                                   tx_coul[coul])
html2 += '''</table></center>'''
display(HTML(html2))