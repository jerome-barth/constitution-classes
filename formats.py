# Formats de base

from config import C_CAT

F_DEF = {'align': 'center', 'valign': 'vcenter', 'font_size': 11}
F_PETIT = {'font_size': 10}
F_MOYEN = {'font_size': 12}
F_GROS = {'font_size': 16}
F_TGROS = {'font_size': 24}
F_HAUT = {'top': 2}
F_BAS = {'bottom': 2}
F_HB = {**F_HAUT, **F_BAS}
F_GAUCHE = {'left': 2}
F_DROITE = {'right': 2}
F_COTES = {**F_GAUCHE, **F_DROITE}
F_BORD = {'border': 2}
F_GRAS = {'bold': True}
F_UNL = {'locked': False}
#
# Formats pour les cellules (composition de formats)
F_TITRE = {**F_DEF, **F_GRAS, **F_TGROS}
F_TOTAUX = {
    **F_DEF,
    **F_GRAS, 'font_color': C_CAT['TOT'][0],
    'bg_color': C_CAT['TOT'][1]
}
F_TOTAUX2 = {
    **F_TOTAUX, 'font_color': C_CAT['TOT2'][0],
    'bg_color': C_CAT['TOT2'][1]
}
F_RENTREE = {**F_DEF, **F_GRAS, **F_GROS}
F_CLS = {**F_DEF}
F_YA = {**F_DEF, **F_PETIT}
F_FAUT = {**F_DEF, **F_PETIT}
F_LST = {**F_DEF, **F_PETIT}
F_LV2 = {**F_DEF, **F_GRAS, **F_BORD, 'bg_color': '#999999'}
F_OPT = {**F_DEF, **F_COTES, 'bg_color': '#cccccc'}
F_ETAB = {**F_DEF, **F_GRAS}
F_ENT = {**F_DEF, **F_GRAS, 'bg_color': C_CAT['TOT'][1]}
F_RESTE1 = {
    **F_DEF, 'font_color': C_CAT['Reste1'][0],
    'bg_color': C_CAT['Reste1'][1]
}
F_RESTE2 = {
    **F_DEF, 'font_color': C_CAT['Reste2'][0],
    'bg_color': C_CAT['Reste2'][1]
}
F_F = {**F_DEF, 'font_color': C_CAT['F'][0], 'bg_color': C_CAT['F'][1]}
F_G = {**F_DEF, 'font_color': C_CAT['G'][0], 'bg_color': C_CAT['G'][1]}
F_PF = {
    **F_DEF,
    **F_GRAS, 'font_color': C_CAT['%F'][0],
    'bg_color': C_CAT['%F'][1]
}
F_NIV = {
    'A': {
        **F_DEF, 'font_color': C_CAT['A'][0],
        'bg_color': C_CAT['A'][1]
    },
    'B': {
        **F_DEF, 'font_color': C_CAT['B'][0],
        'bg_color': C_CAT['B'][1]
    },
    'C': {
        **F_DEF, 'font_color': C_CAT['C'][0],
        'bg_color': C_CAT['C'][1]
    },
    'D': {
        **F_DEF, 'font_color': C_CAT['D'][0],
        'bg_color': C_CAT['D'][1]
    },
    'E': {
        **F_DEF, 'font_color': C_CAT['E'][0],
        'bg_color': C_CAT['E'][1]
    }
}
F_R = {**F_DEF, 'font_color': C_CAT['R'][0], 'bg_color': C_CAT['R'][1]}
F_OPT3 = [
    {
        **F_DEF, 'font_color': C_CAT['opt1'][0],
        'bg_color': C_CAT['opt1'][1]
    },
    {
        **F_DEF, 'font_color': C_CAT['opt2'][0],
        'bg_color': C_CAT['opt2'][1]
    },
    {
        **F_DEF, 'font_color': C_CAT['opt3'][0],
        'bg_color': C_CAT['opt3'][1]
    },
]
F_LV = {**F_DEF, 'font_color': C_CAT['LV2'][0], 'bg_color': C_CAT['LV2'][1]}
F_SLV = {'font_color': C_CAT['sLV2'][0], 'bg_color': C_CAT['sLV2'][1]}
