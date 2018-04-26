
DEBUG = True

from datetime import datetime
from collections import OrderedDict
from IPython.display import HTML, display

##### PARAMETRAGE UTILISATEUR
ANNEE = datetime.today().year % 100  # automatique
# ANNEE = 18                         # manuel : décommenter
ETABLISSEMENT = 'Collège Marie Curie'
VILLE = 'Troyes'
CLASSES = '3e'

NB_DIVS = 4
NOM_DIVS = [2, 3, 4, 5]  # [i+1 for i in range(NB_DIVS)] # Pour 1, 2,...

assert (NB_DIVS == len(NOM_DIVS)
        ), "Il faut autant de noms que de divisions prévues"

NB_ELV = 25 * NB_DIVS + 20  # nb de lignes à prévoir dans la liste d'élèves

LV2S = ['All2', 'Ita2', 'Esp2', 'Sans LV2']
OPTIONS = OrderedDict([
    ('Sans opt', []),
    ('Sport', ['Sport']),
    ('Latin', ['Latin']),
    ('Sp-lat', ['Sport', 'Latin']),
])
OPTIONS_UNIQUES = [opt for opt in OPTIONS if len(OPTIONS[opt]) == 1]

NIVEAUX = ['A', 'B', 'C', 'D', 'E']

NOM_FICHIER = 'R' + str(ANNEE) + '-Repart-' + CLASSES
NOM_FICHIER += '.xls' + ('m' if not DEBUG else 'x')
# NOM_FICHIER = 'R18-Repart-3e.xlsm' # Attention de ne pas mettre '*.xlsx'

##### DEFINITION DES COULEURS
# Pour la couleur de fond des classes (clair, foncé)
C_CLS = [
    ('#66ff99', '#00cc33'),  # vert clair
    ('#99ffff', '#00cccc'),  # cyan
    ('#ff99ff', '#ff00ff'),  # magenta
    ('#ffcc00', '#ff9900'),  # orange
    ('#ffff66', '#ffcc00'),  # jaune clair
    ('#00ccff', '#3399ff'),  # bleu cobalt
    ('#ffcccc', '#cc9999'),  # rose
    ('#99ff66', '#00cc00'),  # vert lime
    ('#ccff00', '#99cc00'),  # jaune citron
    ('#cccccc', '#999999'),  # gris (pour NA)
]

assert (len(C_CLS) > NB_DIVS), "Trop de classes, pas assez de couleurs"

# Pour les étiquettes 'Etiquette': (txt, fond)
C_CAT = {
    'F': ('#990000', '#ff6666'),
    'G': ('#0000cc', '#66ccff'),
    '%F': ('#990099', '#ffccff'),
    'opt1': ('#000099', '#00ffff'),
    'opt2': ('#330066', '#cc99ff'),
    'opt3': ('#660033', '#ff99cc'),
    'LV2': ('#000099', '#99ccff'),
    'A': ('#009900', '#00ff00'),
    'B': ('#669900', '#99ff33'),
    'C': ('#999900', '#ffff00'),
    'D': ('#993300', '#ff6600'),
    'E': ('#990000', '#ff0000'),
    'R': ('#333333', '#999999'),
    'TOT': ('#990000', '#cccccc'),
    'TOT2': ('#990000', '#999999'),
    'CLS': ('#ffffff', '#000000'),
    'Reste1': ('#ff0000', '#ffffcc'),
    'Reste2': ('#ff0000', '#cccc99'),
}

for niv in NIVEAUX:
    assert (niv in C_CAT), "Pb définition des niveaux"

TX_YA = 'Placés'  # 'Il y a'
TX_FAUT = 'Prévus'  # 'Il faut'


# Affichage des couleurs dans le notebook
def jupyter():
    try:
        shell = get_ipython().__class__.__name__
        return shell == 'ZMQInteractiveShell'
    except NameError:
        return False


if jupyter():
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