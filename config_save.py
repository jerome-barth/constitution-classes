from collections import OrderedDict

##### PARAMÉTRAGE UTILISATEUR
### Laisser le # pour avoir l'année en cours
# ANNEE = 18

ETABLISSEMENT = 'Collège Marie Curie'
VILLE = 'Troyes'

CLASSES = '3e'

NB_DIVS = 5
### Exemples de noms de divisions
# NOM_DIVS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'][:NB_DIVS]
# NOM_DIVS = ['α', 'β', 'γ', 'δ', 'ε', 'ζ', 'η', 'θ', 'ι', 'κ', 'λ', 'μ'][:NB_DIVS]
# NOM_DIVS = [10 + i for i in range(NB_DIVS)]  # Pour 10, 11, ...
NOM_DIVS = [i + 1 for i in range(NB_DIVS)]  # Pour 1, 2,...

# Nombre d'élèves prévus (approximatif, des lignes peuvent être ajoutées ou retirées)
NB_ELV = 24 * NB_DIVS + 10
# NB_ELV = 130

# LV2 envisagées : la dernière ('Sans LV2') est traitée de manière spécifique
LV2S = ['All2', 'Ita2', 'Esp2', 'Sans LV2']

# Options compatibles du type : ('Nom', ['opt1', 'opt2',...] )
OPTIONS = OrderedDict([
    ('Sans opt', []), ('Sport', ['Sport']),
    ('Latin', ['Latin']),
    ('Sp-lat', ['Sport', 'Latin']),
#    ('Tricot', ['Tricot']),
#    ('Origami', ['Origami']),
#    ('Tricorilatin', ['Tricot', 'Origami', 'Latin']),
])

# Options pour lesquelles il faut 2 colonnes (typiquement Sport-Étude avec la Section)
OPTIONS_CAT = {'Sport': 'Section'}

# Pour classer les élèves (scolaire et comportement)
NIVEAUX = ['A', 'B', 'C', 'D', 'E']

### Laisser le # pour avoir le nom de fichier par défaut (du type 'R18-Répart-3e.xlsm')
# NOM_FICHIER = 'R18-Répart-3e' # Attention de ne pas mettre d'extension'

##### DÉFINITION DES COULEURS
# Pour la couleur de fond pour les classes : (clair, foncé)
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

# Pour les étiquettes : 'Etiquette': (txt, fond)
C_CAT = {
    'F':      ('#990000', '#ff6666'),  # filles
    'G':      ('#0000cc', '#66ccff'),  # garçons
    '%F':     ('#990099', '#ffccff'),  # pourcentage de filles
    'opt1':   ('#000099', '#00ffff'),  # cycle de 3 couleurs
    'opt2':   ('#330066', '#cc99ff'),  #     pour les différentes
    'opt3':   ('#660033', '#ff99cc'),  #     options uniques
    'LV2':    ('#000099', '#99ccff'),  # couleur pour les options de lv2
    'sLV2':   ('#000066', '#6699cc'),  # couleur pour "Sans LV2"
    'A':      ('#003300', '#00ff00'),  # couleurs pour les différents niveaux
    'B':      ('#003300', '#99ff33'),  #      il faut adapter ceci à la liste
    'C':      ('#333300', '#ffff00'),  #      des niveaux ci-dessus
    'D':      ('#330000', '#ff6600'),  #
    'E':      ('#330000', '#ff0000'),  #
    'R':      ('#333333', '#999999'),  #
    'TOT':    ('#3333cc', '#99cccc'),  # couleur claire pour les totaux
    'TOT2':   ('#3333cc', '#669999'),  # couleur foncée pour les totaux
    'CLS':    ('#ffffff', '#000000'),  # couleur d'entête pour la division
    'Reste1': ('#990000', '#ffffcc'),  # couleur colonne 'Reste' (clair)
    'Reste2': ('#990000', '#cccc99'),  # couleur colonne 'Reste' (foncé)
    'ptR':    ('#ff0000', '#ffffff'),  # couleur pour comportement (avant-avant-dernier)
    'moyR':   ('#660000', '#ff6600'),  # couleur pour comportement (avant-dernier)
    'grR':    ('#660000', '#ff0000'),  # couleur pour comportement (dernier)
    'ERR':    ('#ff0000', '#000000'),  # Disparité prévision/répartition
    'ERRP':   ('#ffffff', '#ff0000'),  # Erreur de structure
}