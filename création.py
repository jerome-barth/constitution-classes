# coding: utf-8
from config import *
from formats import *

import xlsxwriter
from datetime import datetime

assert (NB_DIVS == len(NOM_DIVS)
        ), "Il faut autant de noms que de divisions prévues"
assert ('NA' not in NOM_DIVS), "Le nom 'NA' ne peut pas être choisi"
assert (len(set(NOM_DIVS)) == NB_DIVS
        ), "Les noms de division doivent tous être différents"
assert (len(C_CLS) > NB_DIVS), "Trop de classes, pas assez de couleurs"
for niv in NIVEAUX:
    assert (
        niv in C_CAT
    ), "Une ou plusieurs couleurs pour les niveaux ne sont pas définies"

LV2S_VRAIES = LV2S[:-1]
OPTIONS_UNIQUES = [opt for opt in OPTIONS if len(OPTIONS[opt]) == 1]

try:
    ANNEE
except NameError:
    ANNEE = datetime.today().year % 100  # automatique

try:
    if 'xlsm' != NOM_FICHIER.split('.')[-1]: NOM_FICHIER += '.xlsm'
except NameError:
    NOM_FICHIER = 'R' + str(ANNEE) + '-Répart-' + CLASSES + '.xlsm'

TX_YA = 'Placés'  # 'Il y a'
TX_FAUT = 'Prévus'  # 'Il faut'


def lig_col(lig, col, lig_abs=False, col_abs=False):
    """Retourne la chaîne 'A1' pour une cellule représentée par 0, 0
         lig_abs et col_abs rajoutent des '$' pour un adressage absolu"""

    lig_abs = '$' if lig_abs else ''
    col_abs = '$' if col_abs else ''

    lig += 1
    col += 1
    col_str = ''
    while col:
        reste = col % 26
        if reste == 0: reste = 26
        col_lettre = chr(ord('A') + reste - 1)
        col_str = col_lettre + col_str
        col = (col - 1) // 26
    return col_abs + col_str + lig_abs + str(lig)


def jupyter():
    try:
        shell = get_ipython().__class__.__name__
        return shell == 'ZMQInteractiveShell'
    except NameError:
        return False


if jupyter():
    import hashlib
    if hashlib.sha224(input("Mot de passe ? ").encode('utf-8')).hexdigest(
    ) != '9a48a0a06af408a3200c12938ac9267da16ba1080268603604082c4b':
        raise ValueError('Mot de passe erroné !')

with xlsxwriter.Workbook(NOM_FICHIER) as workbook:

    def pat(lig, col, txt, form):
        """Fct pour écrire dans 'patates' en passant le format en dict"""
        patates.write(lig, col, txt, workbook.add_format(form))

    def pat_merge(l1, c1, l2, c2, txt, form):
        """Fct pour fusionner et écrire dans 'patates' en passant le format"""
        patates.merge_range(l1, c1, l2, c2, txt, workbook.add_format(form))

    def rep(lig, col, txt, form):
        """Fct pour écrire dans '3e-2018-19' avec le format en dict"""
        liste.write(lig, col, txt, workbook.add_format(form))

    # Inclure le VBA ?
    workbook.set_vba_name('ThisWorkbook')
    workbook.add_vba_project('./vbaProject.bin')

    # Propriétés du fichier Excel
    workbook.set_properties({
        'title': 'Répartition ' + CLASSES,
        'subject': 'Rentrée R' + str(ANNEE),
        'author': 'Jérôme BARTH',
        'company': 'Lycée Marie de Champagne',
        'created': datetime.utcnow(),
        'comments': 'Créé avec Python and XlsxWriter'
    })

    # Ajout des feuilles 'Patates' et 'Xe-20XX-XX'
    patates = workbook.add_worksheet('Patates')
    patates.outline_settings(symbols_below=False)
    nom_liste = CLASSES + ' ' + str(2000 + ANNEE) + '-' + str(ANNEE + 1)
    liste = workbook.add_worksheet(nom_liste)
    liste.outline_settings(symbols_below=False)
    liste.set_vba_name('Feuil2')
    ##############################
    ###                        ###
    ###  Feuille 'Xe-20XX-XX'  ###
    ###                        ###
    ##############################
    # Repère dans les lignes
    lig_opt = 6 + len(NIVEAUX)  # ligne des options
    lig_lv2 = lig_opt + len(OPTIONS_UNIQUES)  # ligne des lv2
    dern_recap = lig_lv2 + len(LV2S)  # dernière ligne avant la liste
    premier_el = dern_recap + 2  # ligne du 1er élève
    dernier_el = premier_el + NB_ELV - 1  # ligne du dernier élève

    # Création des noms de plage
    def plage_col(nom, col):
        plage = lig_col(
            premier_el, col, lig_abs=True, col_abs=True) + ":" + lig_col(
                dernier_el, col, lig_abs=True, col_abs=True)
        workbook.define_name('_' + nom, "='" + nom_liste + "'!" + plage)

    plage_col('Nom', 0)
    plage_col('Prénom', 1)
    plage_col('Sexe', 2)
    plage_col('Retard', 3)
    plage_col('Niveau', 4)
    plage_col('Comportement', 5)
    plage_col('Classe', 6)
    plage_col('DivOrig', 7)
    nom_lv2 = []
    sans_lv2 = ''
    for i, lv2 in enumerate(LV2S_VRAIES):
        nosp = ''.join(lv2.split())
        nom_lv2.append('_' + nosp)
        sans_lv2 += '_' + nosp + ',"",'
        plage_col(nosp, 8 + i)
    sans_lv2 = sans_lv2[:-1]
    nom_opt = []
    compt = 0
    for i, opt in enumerate(OPTIONS_UNIQUES):
        nosp = ''.join(opt.split())
        nom_opt.append('_' + nosp)
        plage_col(nosp, 8 + len(LV2S_VRAIES) + i + compt)
        if opt in OPTIONS_CAT: compt += 1

    liste.set_column(0, 0, 20)
    liste.set_column(1, 1, 15)
    liste.set_column(2, 7 + len(LV2S_VRAIES), 6)
    rep(0, 0, 'R' + str(ANNEE), F_RENTREE)
    etab = ETABLISSEMENT.split(' ', maxsplit=1)
    rep(1, 0, etab[0], F_ETAB)
    rep(2, 0, etab[1], F_ETAB)
    rep(3, 0, VILLE, F_ETAB)
    rep(
        0, 1, CLASSES, {
            **F_CLS,
            **F_GRAS,
            **F_GROS,
            **F_BORD, 'font_color': C_CAT['CLS'][0],
            'bg_color': C_CAT['CLS'][1]
        })

    # Tab récap étiquettes
    rep(1, 1, 'Effectif', {**F_ENT, **F_BORD})
    rep(2, 1, 'F', {**F_F, **F_GRAS, **F_HAUT, **F_COTES})
    rep(3, 1, 'G', {**F_G, **F_GRAS, **F_BAS, **F_COTES})
    rep(4, 1, '%F', {**F_PF, **F_BORD, **F_COTES})
    liste.set_row(2, None, None, {'level': 1})
    liste.set_row(3, None, None, {'level': 1})
    liste.set_row(4, None, None, {'level': 1})

    # Tab récap étiquettes niveaux
    for i, niv in enumerate(NIVEAUX):
        bordure = {**F_COTES}
        if i == 0: bordure = {**bordure, **F_HAUT}
        if i == len(NIVEAUX) - 1: bordure = {**bordure, **F_BAS}
        rep(5 + i, 1, niv, {**F_NIV[niv], **F_GRAS, **bordure})
        if i != 0: liste.set_row(5 + i, None, None, {'level': 1})
    rep(lig_opt - 1, 1, 'R', {**F_R, **F_GRAS, **F_BORD})

    # Tab récap étiquettes options
    for i, opt in enumerate(OPTIONS_UNIQUES):
        bordure = {**F_COTES}
        if i == 0: bordure = {**bordure, **F_HAUT}
        if i == len(OPTIONS_UNIQUES) - 1: bordure = {**bordure, **F_BAS}
        rep(lig_opt + i, 1, opt, {**F_OPT3[i % 3], **F_GRAS, **bordure})
        liste.set_row(lig_opt + i, None, None, {'level': 1})

    # Tab récap étiquettes LV2
    for i, lv2 in enumerate(LV2S):
        bordure = {**F_COTES}
        if i == 0: bordure = {**bordure, **F_HAUT}
        if i == len(LV2S) - 1: bordure = {**F_SLV, **bordure, **F_HB}
        rep(lig_lv2 + i, 1, lv2, {**F_LV, **F_GRAS, **bordure})
        if i != 0: liste.set_row(lig_lv2 + i, None, None, {'level': 1})

    # Dernière ligne avant la liste
    liste.set_row(dern_recap, 5)
    for i in range(NB_DIVS + 4):
        rep(dern_recap, i + 1, None, F_HAUT)

    # Tableau récap : colonnes des classes
    for div, nom_div in enumerate(NOM_DIVS):
        bordure = {}
        if div == 0: bordure = {**F_GAUCHE}
        if div == NB_DIVS - 1: bordure = {**bordure, **F_DROITE}
        rep(
            0, 2 + div, nom_div, {
                **F_CLS,
                **F_GRAS,
                **F_GROS,
                **F_HB,
                **bordure, 'bg_color': C_CLS[div][0]
            })
        rep(1, 2 + div, '=COUNTIF(_Classe,' + lig_col(0, 2 + div) + ')', {
            **F_CLS,
            **F_GRAS,
            **F_HB,
            **bordure, 'bg_color': C_CLS[div][1]
        })
        rep(2, 2 + div,
            '=COUNTIFS(_Classe,' + lig_col(0, 2 + div) + ',_Sexe,B3)', {
                **F_CLS,
                **F_HAUT,
                **bordure, 'bg_color': C_CLS[div][0]
            })
        rep(3, 2 + div,
            '=COUNTIFS(_Classe,' + lig_col(0, 2 + div) + ',_Sexe,B4)', {
                **F_CLS,
                **F_BAS,
                **bordure, 'bg_color': C_CLS[div][0]
            })
        rep(4, 2 + div, '=IF(' + lig_col(1, 2 + div) + '<>0,ROUND(' +
            lig_col(2, 2 + div) + '/' + lig_col(1, 2 + div) + '*100,0),"*")', {
                **F_CLS,
                **F_GRAS,
                **F_HB,
                **bordure, 'bg_color': C_CLS[div][1]
            })

        for i, niv in enumerate(NIVEAUX):
            bordure2 = {}
            if i == 0: bordure2 = F_HAUT
            if i == len(NIVEAUX) - 1: bordure2 = {**bordure2, **F_BAS}
            rep(5 + i, 2 + div, '=COUNTIFS(_Classe,' + lig_col(0, 2 + div) +
                ',_Niveau,' + lig_col(5 + i, 1) + ')', {
                    **F_CLS,
                    **bordure,
                    **bordure2, 'bg_color': C_CLS[div][0]
                })
        rep(5 + len(NIVEAUX), 2 + div,
            '=COUNTIFS(_Classe,' + lig_col(0, 2 + div) + ',_Retard,"R")', {
                **F_CLS,
                **F_HB,
                **bordure, 'bg_color': C_CLS[div][1]
            })
        for i, opt in enumerate(OPTIONS_UNIQUES):
            bordure2 = {}
            if i == 0: bordure2 = {**bordure2, **F_HAUT}
            if i == len(OPTIONS_UNIQUES) - 1: bordure2 = {**bordure2, **F_BAS}
            rep(lig_opt + i, 2 + div, '=COUNTIFS(_Classe,' +
                lig_col(0, 2 + div) + ',' + nom_opt[i] + ',1)', {
                    **F_CLS,
                    **bordure,
                    **bordure2, 'bg_color': C_CLS[div][0]
                })
        for i, lv2 in enumerate(LV2S):
            bordure2 = {}
            if i == 0: bordure2 = {**bordure2, **F_HAUT}
            if i == len(LV2S) - 1:
                bordure2 = {**bordure2, **F_HB}
                rep(lig_lv2 + i, 2 + div, '=COUNTIFS(_Classe,' +
                    lig_col(0, 2 + div) + ',' + sans_lv2 + ')', {
                        **F_CLS,
                        **bordure,
                        **bordure2, 'bg_color': C_CLS[div][1]
                    })
            else:
                rep(lig_lv2 + i, 2 + div, '=COUNTIFS(_Classe,' +
                    lig_col(0, 2 + div) + ',' + nom_lv2[i] + ',1)', {
                        **F_CLS,
                        **bordure,
                        **bordure2, 'bg_color': C_CLS[div][0]
                    })
        liste.conditional_format(
            1, 2 + div, 1, 2 + div, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][1]
                })
            })
        liste.conditional_format(
            2, 2 + div, 3, 2 + div, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][0]
                })
            })
        liste.conditional_format(
            5, 2 + div, 4 + len(NIVEAUX), 2 + div, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][0]
                }),
                'stop_if_true': True,
            })
        liste.conditional_format(
            5, 2 + div, 4 + len(NIVEAUX), 2 + div, {
                'type': 'data_bar',
                'bar_solid': True,
                'bar_color': C_CLS[div][1],
                'bar_axis_position': 'none',
            })
        liste.conditional_format(
            5 + len(NIVEAUX), 2 + div, 5 + len(NIVEAUX), 2 + div, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][1]
                }),
                'stop_if_true': True,
            })
        liste.conditional_format(
            6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
            2 + div,
            6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
            2 + div, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][1]
                }),
                'stop_if_true': True,
            })
        liste.conditional_format(
            5 + len(NIVEAUX) + 1, 2 + div,
            5 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
            2 + div, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][0]
                }),
                'stop_if_true': True,
            })

    liste.conditional_format(
        4, 2, 4, NB_DIVS + 4, {
            'type': 'data_bar',
            'bar_color': '#cc66cc',
            'bar_border_color': '#990099',
            'bar_axis_position': 'none',
            'min_type': 'num',
            'max_type': 'num',
            'min_value': -2,
            'max_value': 102,
        })

    # Colonne NA
    rep(0, 2 + NB_DIVS, 'NA', {
        **F_CLS,
        **F_GRAS,
        **F_GROS,
        **F_BORD, 'bg_color': C_CLS[-1][0]
    })
    rep(1, 2 + NB_DIVS, '=COUNTIF(_Classe,"NA")', {
        **F_CLS,
        **F_GRAS,
        **F_BORD,
        **bordure, 'bg_color': C_CLS[-1][1]
    })
    rep(2, 2 + NB_DIVS, '=COUNTIFS(_Classe,"NA",_Sexe,B3)', {
        **F_CLS,
        **F_HAUT,
        **F_COTES, 'bg_color': C_CLS[-1][0]
    })
    rep(3, 2 + NB_DIVS, '=COUNTIFS(_Classe,"NA",_Sexe,B4)', {
        **F_CLS,
        **F_BAS,
        **F_COTES, 'bg_color': C_CLS[-1][0]
    })
    rep(4, 2 + NB_DIVS,
        '=IF(' + lig_col(1, 2 + NB_DIVS) + '<>0,ROUND(' + lig_col(
            2, 2 + NB_DIVS) + '/' + lig_col(1, 2 + NB_DIVS) + '*100,0),"*")', {
                **F_CLS,
                **F_GRAS,
                **F_BORD, 'bg_color': C_CLS[-1][1]
            })
    for i, niv in enumerate(NIVEAUX):
        bordure2 = {}
        if i == 0: bordure2 = F_HAUT
        if i == len(NIVEAUX) - 1: bordure2 = {**bordure2, **F_BAS}
        rep(5 + i, 2 + NB_DIVS,
            '=COUNTIFS(_Classe,"NA",_Niveau,' + lig_col(5 + i, 1) + ')', {
                **F_CLS,
                **F_COTES,
                **bordure2, 'bg_color': C_CLS[-1][0]
            })
    rep(5 + len(NIVEAUX), 2 + NB_DIVS, '=COUNTIFS(_Classe,"NA",_Retard,"R")', {
        **F_CLS,
        **F_HB,
        **F_COTES, 'bg_color': C_CLS[-1][1]
    })
    for i, opt in enumerate(OPTIONS_UNIQUES):
        bordure2 = {}
        if i == 0: bordure2 = {**bordure2, **F_HAUT}
        if i == len(OPTIONS_UNIQUES) - 1: bordure2 = {**bordure2, **F_BAS}
        rep(lig_opt + i, 2 + NB_DIVS,
            '=COUNTIFS(_Classe,"NA",' + nom_opt[i] + ',1)', {
                **F_CLS,
                **F_COTES,
                **bordure2, 'bg_color': C_CLS[-1][0]
            })
    for i, lv2 in enumerate(LV2S):
        bordure2 = {}
        if i == 0: bordure2 = {**bordure2, **F_HAUT}
        if i == len(LV2S) - 1:
            bordure2 = {**bordure2, **F_HB}
            rep(lig_lv2 + i, 2 + NB_DIVS,
                '=COUNTIFS(_Classe,"NA",' + sans_lv2 + ')', {
                    **F_CLS,
                    **F_COTES,
                    **bordure2, 'bg_color': C_CLS[-1][1]
                })
        else:
            rep(lig_lv2 + i, 2 + NB_DIVS,
                '=COUNTIFS(_Classe,"NA",' + nom_lv2[i] + ',1)', {
                    **F_CLS,
                    **F_COTES,
                    **bordure2, 'bg_color': C_CLS[-1][0]
                })
    liste.conditional_format(
        1, 2 + NB_DIVS, 1, 2 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CLS[-1][1]
            })
        })
    liste.conditional_format(
        2, 2 + NB_DIVS, 3, 2 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CLS[-1][0]
            })
        })
    liste.conditional_format(
        5, 2 + NB_DIVS, 4 + len(NIVEAUX), 2 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CLS[-1][0]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        5, 2 + NB_DIVS, 4 + len(NIVEAUX), 2 + NB_DIVS, {
            'type': 'data_bar',
            'bar_solid': True,
            'bar_color': C_CLS[-1][1],
            'bar_axis_position': 'none',
        })
    liste.conditional_format(
        5 + len(NIVEAUX), 2 + NB_DIVS, 5 + len(NIVEAUX), 2 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CLS[-1][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        2 + NB_DIVS,
        6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        2 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CLS[-1][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        5 + len(NIVEAUX) + 1, 2 + NB_DIVS,
        5 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        2 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CLS[-1][0]
            }),
            'stop_if_true': True,
        })

    # Colonne Reste
    rep(0, 3 + NB_DIVS, 'Reste', {
        **F_RESTE1,
        **F_GRAS,
        **F_BORD,
    })
    rep(1, 3 + NB_DIVS, '=COUNTBLANK(_Classe)', {
        **F_RESTE2,
        **F_GRAS,
        **F_BORD,
    })

    rep(2, 3 + NB_DIVS, '=COUNTIFS(_Classe,"",_Sexe,B3)', {
        **F_RESTE1,
        **F_HAUT,
        **F_COTES,
    })
    rep(3, 3 + NB_DIVS, '=COUNTIFS(_Classe,"",_Sexe,B4)', {
        **F_RESTE1,
        **F_BAS,
        **F_COTES,
    })
    rep(4, 3 + NB_DIVS,
        '=IF(' + lig_col(1, 3 + NB_DIVS) + '<>0,ROUND(' + lig_col(
            2, 3 + NB_DIVS) + '/' + lig_col(1, 3 + NB_DIVS) + '*100,0),"*")', {
                **F_RESTE2,
                **F_GRAS,
                **F_BORD,
            })
    for i, niv in enumerate(NIVEAUX):
        bordure2 = {}
        if i == 0: bordure2 = F_HAUT
        if i == len(NIVEAUX) - 1: bordure2 = {**bordure2, **F_BAS}
        rep(5 + i, 3 + NB_DIVS,
            '=COUNTIFS(_Classe,"",_Niveau,' + lig_col(5 + i, 1) + ')', {
                **F_RESTE1,
                **F_COTES,
                **bordure2,
            })
    rep(5 + len(NIVEAUX), 3 + NB_DIVS, '=COUNTIFS(_Classe,"",_Retard,"R")', {
        **F_RESTE2,
        **F_HB,
        **F_COTES,
    })
    for i, opt in enumerate(OPTIONS_UNIQUES):
        bordure2 = {}
        if i == 0: bordure2 = {**bordure2, **F_HAUT}
        if i == len(OPTIONS_UNIQUES) - 1: bordure2 = {**bordure2, **F_BAS}
        rep(lig_opt + i, 3 + NB_DIVS,
            '=COUNTIFS(_Classe,"",' + nom_opt[i] + ',1)', {
                **F_RESTE1,
                **F_COTES,
                **bordure2,
            })
    for i, lv2 in enumerate(LV2S):
        bordure2 = {}
        if i == 0: bordure2 = {**bordure2, **F_HAUT}
        if i == len(LV2S) - 1:
            bordure2 = {**bordure2, **F_HB}
            rep(lig_lv2 + i, 3 + NB_DIVS,
                '=COUNTIFS(_Classe,"",' + sans_lv2 + ')', {
                    **F_RESTE2,
                    **F_COTES,
                    **bordure2,
                })
        else:
            rep(lig_lv2 + i, 3 + NB_DIVS,
                '=COUNTIFS(_Classe,"",' + nom_lv2[i] + ',1)', {
                    **F_RESTE1,
                    **F_COTES,
                    **bordure2,
                })
    liste.conditional_format(
        1, 3 + NB_DIVS, 1, 3 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['Reste2'][1]
            })
        })
    liste.conditional_format(
        2, 3 + NB_DIVS, 3, 3 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['Reste1'][1]
            })
        })
    liste.conditional_format(
        5, 3 + NB_DIVS, 4 + len(NIVEAUX), 3 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['Reste1'][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        5, 3 + NB_DIVS, 4 + len(NIVEAUX), 3 + NB_DIVS, {
            'type': 'data_bar',
            'bar_solid': True,
            'bar_color': C_CAT['Reste2'][1],
            'bar_axis_position': 'none',
        })
    liste.conditional_format(
        5 + len(NIVEAUX), 3 + NB_DIVS, 5 + len(NIVEAUX), 3 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['Reste2'][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        3 + NB_DIVS,
        6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        3 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['Reste2'][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        5 + len(NIVEAUX) + 1, 3 + NB_DIVS,
        5 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        3 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['Reste1'][1]
            }),
            'stop_if_true': True,
        })

    # Colonne Totaux
    rep(0, 4 + NB_DIVS, 'Totaux', {
        **F_TOTAUX,
        **F_BORD,
    })
    for i in range(5 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S)):
        l_retard = 4 + len(NIVEAUX)
        l_lv2 = l_retard + len(OPTIONS_UNIQUES)
        l_sans_lv2 = l_lv2 + len(LV2S)
        bordure = F_COTES
        if i in [0, l_retard, l_sans_lv2]:
            bordure = F_BORD
        if i in [1, 4, l_retard + 1, l_lv2 + 1]:
            bordure = {**bordure, **F_HAUT}
        if i in [2, l_retard - 1, l_lv2, l_sans_lv2 - 1]:
            bordure = {**bordure, **F_BAS}
        if i == 0: bordure = {**bordure, **F_GRAS}
        coul = F_TOTAUX2 if i in [0, l_retard, l_sans_lv2] else F_TOTAUX
        if i == 3:
            rep(i + 1, 4 + NB_DIVS, '=IF(' + lig_col(1, 4 + NB_DIVS) +
                '<>0,ROUND(' + lig_col(2, 4 + NB_DIVS) + '/' +
                lig_col(1, 4 + NB_DIVS) + '*100,0),"*")', {
                    **F_TOTAUX2,
                    **F_BORD,
                })
        else:
            rep(i + 1, 4 + NB_DIVS, '=SUM(' + lig_col(i + 1, 2) + ':' +
                lig_col(i + 1, 3 + NB_DIVS) + ')', {
                    **coul,
                    **bordure,
                })
    liste.conditional_format(
        1, 4 + NB_DIVS, 1, 4 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['TOT2'][1]
            })
        })
    liste.conditional_format(
        2, 4 + NB_DIVS, 3, 4 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['TOT'][1]
            })
        })
    liste.conditional_format(
        5, 4 + NB_DIVS, 4 + len(NIVEAUX), 4 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['TOT'][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        5, 4 + NB_DIVS, 4 + len(NIVEAUX), 4 + NB_DIVS, {
            'type': 'data_bar',
            'bar_solid': True,
            'bar_color': C_CAT['TOT2'][1],
            'bar_axis_position': 'none',
        })
    liste.conditional_format(
        5 + len(NIVEAUX), 4 + NB_DIVS, 5 + len(NIVEAUX), 4 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['TOT2'][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        4 + NB_DIVS,
        6 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        4 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['TOT2'][1]
            }),
            'stop_if_true': True,
        })
    liste.conditional_format(
        5 + len(NIVEAUX) + 1, 4 + NB_DIVS,
        5 + len(NIVEAUX) + len(OPTIONS_UNIQUES) + len(LV2S_VRAIES),
        4 + NB_DIVS, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['TOT'][1]
            }),
            'stop_if_true': True,
        })

    # -~- Liste -~-
    # Entête
    rep(dern_recap + 1, 0, 'Nom', {**F_ENT, **F_HB, **F_GAUCHE, **F_UNL})
    fmt = workbook.add_format({**F_ENT, **F_HB, **F_UNL})
    liste.write_row(
        dern_recap + 1, 1,
        ['Prénom', 'Sexe', 'Retard', 'Niv', 'Comp', CLASSES, 'Cls orig'], fmt)
    rep(
        dern_recap + 1, 6, CLASSES, {
            **F_DEF,
            **F_HB,
            **F_UNL,
            'font_color': C_CAT['CLS'][0],
            'bg_color': C_CAT['CLS'][1],
        })
    for i, lv2 in enumerate(LV2S_VRAIES):
        rep(dern_recap + 1, 8 + i, lv2, {**F_LV, **F_GRAS, **F_HB, **F_UNL})
    compt = 0
    for i, opt in enumerate(OPTIONS_UNIQUES):
        col = 8 + len(LV2S_VRAIES) + i + compt
        rep(dern_recap + 1, col, opt, {
            **F_OPT3[i % 3],
            **F_GRAS,
            **F_HB,
            **F_UNL,
        })
        if opt in OPTIONS_CAT:
            compt += 1
            liste.set_column(col, col, 6)
            liste.set_column(col + 1, col + 1, 15)
            rep(dern_recap + 1, col + 1, OPTIONS_CAT[opt], {
                **F_OPT3[i % 3],
                **F_GRAS,
                **F_HB,
                **F_UNL,
            })
        else:
            liste.set_column(col, col, 6)
    col = 8 + len(LV2S_VRAIES) + compt + len(OPTIONS_UNIQUES)
    liste.set_column(col, col, 25)
    rep(dern_recap + 1, col, 'Observations', {
        **F_ENT,
        **F_HB,
        **F_DROITE,
        **F_UNL,
    })
    for i in range(col + 1):
        rep(dernier_el + 1, i, None, F_HAUT)

    # Déverrouille les lignes entête et élèves
    liste.set_row(premier_el - 1, 20, workbook.add_format(F_UNL))
    for el in range(NB_ELV):
        liste.set_row(premier_el + el, None, workbook.add_format(F_UNL))
    liste.autofilter(premier_el - 1, 0, dernier_el, col)

    # Validation de données et formatage
    # NA -> ligne grisée
    liste.conditional_format(
        premier_el, 0, dernier_el,
        len(LV2S_VRAIES) + len(OPTIONS_UNIQUES) + len(OPTIONS_CAT)+8, {
            'type':
            'formula',
            'criteria':
            '=' + lig_col(premier_el, 6, False, True) + '="NA"',
            'stop_if_true':
            True,
            'format':
            workbook.add_format({
                'font_color': C_CLS[-1][1],
                'bg_color': C_CLS[-1][0],
            })
        })
    # Erreurs de lv2
    liste.conditional_format(
        premier_el, 0, dernier_el,
        len(LV2S_VRAIES) + len(OPTIONS_UNIQUES) + len(OPTIONS_CAT)+8, {
            'type':
            'formula',
            'criteria':
            '=SUM(' + lig_col(premier_el, 8, False, True) + ':' +
            lig_col(premier_el, 7 + len(LV2S_VRAIES), False, True) + ')>1',
            'stop_if_true':
            True,
            'format':
            workbook.add_format(F_ERR),
        })

    #   Sexe : F ou G ou ''
    liste.data_validation(premier_el, 2, dernier_el, 2, {
        'validate': 'list',
        'source': '=$B$3:$B$4'
    })

    liste.conditional_format(
        premier_el, 2, dernier_el, 2, {
            'type':
            'cell',
            'criteria':
            'equal to',
            'value':
            '"F"',
            'format':
            workbook.add_format({
                'font_color': C_CAT['F'][0],
                'bg_color': C_CAT['F'][1]
            })
        })
    liste.conditional_format(
        premier_el, 2, dernier_el, 2, {
            'type':
            'cell',
            'criteria':
            'equal to',
            'value':
            '"G"',
            'format':
            workbook.add_format({
                'font_color': C_CAT['G'][0],
                'bg_color': C_CAT['G'][1]
            })
        })
    #   Retard: R ou ''
    liste.data_validation(premier_el, 3, dernier_el, 3, {
        'validate': 'list',
        'source': '=' + lig_col(lig_opt - 1, 1, True, True)
    })
    liste.conditional_format(
        premier_el, 3, dernier_el, 3, {
            'type':
            'cell',
            'criteria':
            'equal to',
            'value':
            '"R"',
            'format':
            workbook.add_format({
                'font_color': C_CAT['R'][0],
                'bg_color': C_CAT['R'][1]
            })
        })

    #   Niv / Comp
    liste.data_validation(
        premier_el, 4, dernier_el, 5, {
            'validate':
            'list',
            'source':
            '=' + lig_col(5, 1, True, True) + ':' +
            lig_col(4 + len(NIVEAUX), 1, True, True)
        })
    for i, niv in enumerate(NIVEAUX):
        liste.conditional_format(
            premier_el, 4, dernier_el, 5, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': '"' + niv + '"',
                'format': workbook.add_format(F_NIV[niv])
            })

    #   Classe
    liste.data_validation(
        premier_el, 6, dernier_el, 6, {
            'validate':
            'list',
            'source':
            '=' + lig_col(0, 2, True, True) + ':' +
            lig_col(0, 2 + NB_DIVS, True, True)
        })
    #    coloration des noms suivant le comportement
    #     -> petit rouge : comportement C
    if len(NIVEAUX) > 4:
        liste.conditional_format(
            premier_el, 0, dernier_el, 1, {
                'type':
                'formula',
                'criteria':
                '=' + lig_col(premier_el, 5, False, True) + '="' +
                str(NIVEAUX[-3] + '"'),
                'format':
                workbook.add_format({
                    'font_color': C_CAT['ptR'][0],
                    'bold': True,
                    'italic': True,
                })
            })
    #     -> moyen rouge : comportement D
    if len(NIVEAUX) > 3:
        liste.conditional_format(
            premier_el, 0, dernier_el, 1, {
                'type':
                'formula',
                'criteria':
                '=' + lig_col(premier_el, 5, False, True) + '="' +
                str(NIVEAUX[-2] + '"'),
                'format':
                workbook.add_format({
                    'font_color': C_CAT['moyR'][0],
                    'bg_color': C_CAT['moyR'][1],
                    'bold': True,
                    'italic': True,
                })
            })
    #     -> gros rouge  : comportement E
    if len(NIVEAUX) > 2:
        liste.conditional_format(
            premier_el, 0, dernier_el, 1, {
                'type':
                'formula',
                'criteria':
                '=' + lig_col(premier_el, 5, False, True) + '="' +
                str(NIVEAUX[-1] + '"'),
                'format':
                workbook.add_format({
                    'font_color': C_CAT['grR'][0],
                    'bg_color': C_CAT['grR'][1],
                    'bold': True,
                    'italic': True,
                })
            })
    #    coloration des noms/prénoms/classes suivant les classes
    for i, classe in enumerate(NOM_DIVS):
        if not type(classe) in [int, float]: classe = '"' + classe + '"'
        liste.conditional_format(
            premier_el, 6, dernier_el, 6, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': classe,
                'format': workbook.add_format({
                    'bg_color': C_CLS[i][0]
                })
            })
        liste.conditional_format(
            premier_el, 0, dernier_el, 1, {
                'type':
                'formula',
                'criteria':
                '=' + lig_col(premier_el, 6, False, True) + '=' + str(classe),
                'format':
                workbook.add_format({
                    'bg_color': C_CLS[i][0]
                })
            })
    liste.conditional_format(
        premier_el, 6, dernier_el, 6, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': '"NA"',
            'format': workbook.add_format({
                'bg_color': C_CLS[-1][0],
            })
        })
    liste.conditional_format(
        premier_el, 6, dernier_el, 6, {
            'type':
            'cell',
            'criteria':
            'equal to',
            'value':
            '""',
            'format':
            workbook.add_format({
                'font_color': C_CAT['Reste1'][1],
                'bg_color': C_CAT['Reste1'][0],
            })
        })

    #   LV2
    liste.data_validation(premier_el, 8, dernier_el, 7 + len(LV2S_VRAIES), {
        'validate': 'integer',
        'criteria': 'equal to',
        'value': 1
    })
    liste.conditional_format(
        premier_el, 8, dernier_el, 7 + len(LV2S_VRAIES), {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 1,
            'format': workbook.add_format(F_LV)
        })
    #   Options
    compt = 0
    for i, opt in enumerate(OPTIONS_UNIQUES):
        liste.data_validation(premier_el, 8 + len(LV2S_VRAIES) + i + compt,
                              dernier_el, 8 + len(LV2S_VRAIES) + i + compt, {
                                  'validate': 'integer',
                                  'criteria': 'equal to',
                                  'value': 1
                              })
        deux_col = opt in OPTIONS_CAT
        liste.conditional_format(
            premier_el, 8 + len(LV2S_VRAIES) + i + compt, dernier_el,
            8 + len(LV2S_VRAIES) + i + compt + deux_col, {
                'type': 'cell',
                'criteria': 'not equal to',
                'value': '""',
                'format': workbook.add_format(F_OPT3[i % 3])
            })
        if opt in OPTIONS_CAT: compt += 1

    # coloration ligne sans classe
    liste.conditional_format(
        premier_el, 0, dernier_el,
        len(LV2S_VRAIES) + len(OPTIONS_UNIQUES) + len(OPTIONS_CAT)+8, {
            'type':
            'formula',
            'criteria':
            '=' + lig_col(premier_el, 6, False, True) + '=""',
            'format':
            workbook.add_format({
                'font_color': C_CAT['Reste1'][0],
                'bg_color': C_CAT['Reste1'][1],
                'bold': True,
            })
        })
    ### TESTS ###
    # insertion de données de test
    test_nom = [
        'AGOSTINHO', 'ARMATA', 'ARNON', 'ATIFI', 'AUNOS', 'BARTH', 'BELNOT',
        'BERTHEUX', 'BIANCHI', 'CANDERAN', 'CARDONNE', 'CLAUSSE', 'CONVERT',
        'CROZAT', 'CRÉVOULIN', 'DE MARCH', 'DEFRANCE', 'DEHRI', 'DUS',
        'FERREIRA', 'GARNIER', 'GHARBI', 'GOTTRANT', 'HALATA', 'HAUMESSER',
        'HURPOIL', 'LABILLE', 'LLATY', 'LOISEAU', 'NOEL', 'PANTALEON',
        'PIERRE', 'POCHOT', 'POINTU', 'PRÉVOT', 'REGNIER', 'RIVIÈRE',
        'RONDEAU', 'ROUVROY', 'RUBASZEWSKI', 'SERVAS', 'SIMON', 'SIMONNOT',
        'SPAY', 'TILLIET', 'TRINQUET', 'TROCHAIN', 'VANHAREN', 'WOJSZVZYK'
    ]
    test_prenom = {
        'Alexandre': 'G',
        'Aline': 'F',
        'Anne': 'F',
        'Carole': 'F',
        'Cassandre': 'F',
        'Clothilde': 'F',
        'Cécile': 'F',
        'Cédric': 'G',
        'Céline': 'F',
        'Daniel': 'G',
        'Didier': 'G',
        'Dominique': 'G',
        'Fanny': 'F',
        'Florence': 'F',
        'Franck': 'G',
        'Frédéric': 'G',
        'Frédérique': 'F',
        'Guillaume': 'G',
        'Isabelle': 'F',
        'Jean-Pierre': 'G',
        'Julie': 'F',
        'Jérémie': 'G',
        'Jérôme': 'G',
        'Katia': 'F',
        'Mangni': 'F',
        'Marie': 'F',
        'Marie-Louis': 'G',
        'Marie-Lourdes': 'F',
        'Mathieu': 'G',
        'Mohammed': 'G',
        'Nathalie': 'F',
        'Nelly': 'F',
        'Nicolas': 'G',
        'Rachida': 'F',
        'Romain': 'G',
        'Sandra': 'F',
        'Sandrine': 'F',
        'Stéphane': 'G',
        'Séverine': 'F',
        'Virginie': 'F',
        'Émilie': 'F',
        'Évelyne': 'F'
    }
    test_retard = [''] * 10 + ['R']
    test_div = ['4e' + str(i + 1) for i in range(8)] + ['Nv3e']
    test_lv2 = []
    for i in range(len(LV2S_VRAIES)):
        lg=[]
        for j in range(len(LV2S_VRAIES)):
            if i==j:
                lg.append(1)
            else:
                lg.append('')
        test_lv2.append(lg)
    test_sp = [['', '']] * 10 + [[1, 'FOOTBALL'], [1, 'GYMNASTIQUE'],
                                 [1, 'BASKET']]
    test_lat = [''] * 10 + [1]
    test_obs = [''] * 5 + [
        'Teigne', 'Peste', 'Pénible', 'Insupportable', 'Très fort',
        'Bon en maths', 'Bon en français', 'Bon en sport'
    ]

    from random import choice
    for el in range(NB_ELV):
        nom = choice(test_nom)
        prenom, sexe = choice(list(test_prenom.items()))
        retard = choice(test_retard)
        niv = choice(NIVEAUX[:-2] * 5 + NIVEAUX[-2:] * 2)
        comp = choice(NIVEAUX[:-3] * 5 + NIVEAUX[-3:])
        classe = choice(NOM_DIVS * 5 + [''] * NB_DIVS * 10 + ['NA'] * NB_DIVS)
        div_orig = choice(test_div)
        lv2 = choice(test_lv2)
        opts = []
        for opt in OPTIONS_UNIQUES:
            if opt in OPTIONS_CAT:
                opts +=choice(test_sp)
            else:
                opts.append(choice(test_lat))
        obs = choice(test_obs)
        rep(premier_el + el, 0, nom, {
            **F_DEF,
            **F_GAUCHE,
            **F_UNL, 'align': 'left'
        })
        rep(premier_el + el, 1, prenom, {**F_DEF, **F_UNL, 'align': 'left'})
        rep(premier_el + el, 2, sexe, {**F_DEF, **F_UNL})
        liste.write_row(premier_el + el, 2,
                        [sexe, retard, niv, comp, classe, div_orig] + lv2+opts, workbook.add_format({
                            **F_DEF,
                            **F_UNL
                        }))
        rep(premier_el + el,
            len(LV2S_VRAIES) + len(OPTIONS_UNIQUES) + len(OPTIONS_CAT)+8,
             obs, {
                **F_DEF,
                **F_DROITE,
                **F_UNL, 'align': 'left'
            })

    ###########################
    ###                     ###
    ###  Feuille 'Patates'  ###
    ###                     ###
    ###########################
    patates.set_row(0, 42)
    pat_merge(0, 0, 0, 2 * NB_DIVS + 4,
              ETABLISSEMENT + ' - ' + VILLE + ' - ' + 'Rentrée R' + str(ANNEE),
              F_TITRE)
    patates.set_column(0, 0, 10)
    patates.set_column(1, 2 * NB_DIVS + 4, 7)
    pat(1, 0, 'R' + str(ANNEE), F_RENTREE)
    pat(
        2, 0, CLASSES, {
            **F_CLS,
            **F_GRAS,
            **F_GROS,
            **F_BORD, 'font_color': C_CAT['CLS'][0],
            'bg_color': C_CAT['CLS'][1]
        })

    # Nombre de lignes par lv2
    nbl = len(OPTIONS) + 1
    # N° de la ligne après la dernière LV2
    dern = len(LV2S) * (len(OPTIONS) + 1) + 3
    # N° de la ligne après la liste des options
    dern2 = dern + len(OPTIONS) + 2

    for div, nom_div in enumerate(NOM_DIVS):  # boucle sur les classes
        pat(1, 2 * div + 1, TX_YA, F_YA)
        pat(1, 2 * div + 2, TX_FAUT, {
            **F_FAUT,
            **F_BORD, 'bg_color': C_CLS[div][1]
        })
        pat_merge(2, 2 * div + 1, 2, 2 * div + 2, nom_div, {
            **F_CLS,
            **F_GRAS,
            **F_GROS,
            **F_BORD, 'bg_color': C_CLS[div][0]
        })
        col = 2 * div + 2
        for i, lv2 in enumerate(LV2S):  # boucles sur les différentes LV2
            lig = i * nbl + 3
            if i == len(LV2S) - 1:
                pat(
                    lig, col - 1, '=COUNTIFS(_Classe,' + lig_col(2, col - 1) +
                    ',' + sans_lv2 + ')', {
                        **F_CLS,
                        **F_GRAS,
                        **F_MOYEN,
                        **F_HB,
                        **F_GAUCHE, 'bg_color': C_CLS[div][1]
                    })
            else:
                pat(
                    lig, col - 1, '=COUNTIFS(_Classe,' + lig_col(2, col - 1) +
                    ',' + nom_lv2[i] + ',1)', {
                        **F_CLS,
                        **F_GRAS,
                        **F_MOYEN,
                        **F_HB,
                        **F_GAUCHE, 'bg_color': C_CLS[div][1]
                    })
            pat(
                lig, col, '=SUM(' + lig_col(lig + 1, col) + ':' +
                lig_col(lig + nbl - 1, col) + ')', {
                    **F_CLS,
                    **F_GRAS,
                    **F_MOYEN,
                    **F_HB,
                    **F_DROITE, 'bg_color': C_CLS[div][1]
                })
            for j, opt in enumerate(OPTIONS):  # boucles sur les options
                bordure = {}
                if j == 0: bordure = F_HAUT
                if j == len(OPTIONS) - 1: bordure = {**bordure, **F_BAS}
                cond = []
                for option in OPTIONS_UNIQUES:
                    nosp = '_' + ''.join(option.split())
                    cond.append(nosp)
                    if option in OPTIONS[opt]: cond.append('1')
                    else: cond.append('""')
                    txt_cond = ','.join(cond)
                if i == len(LV2S) - 1:
                    pat(
                        lig + j + 1, col - 1,
                        '=COUNTIFS(_Classe,' + lig_col(2, col - 1) + ',' +
                        sans_lv2 + ',' + txt_cond + ')', {
                            **F_CLS,
                            **F_PETIT,
                            **F_GAUCHE,
                            **bordure, 'bg_color': C_CLS[div][0]
                        })
                else:
                    pat(
                        lig + j + 1, col - 1,
                        '=COUNTIFS(_Classe,' + lig_col(2, col - 1) + ',' +
                        nom_lv2[i] + ',1,' + txt_cond + ')', {
                            **F_CLS,
                            **F_PETIT,
                            **F_GAUCHE,
                            **bordure, 'bg_color': C_CLS[div][0]
                        })
                pat(
                    lig + j + 1, col, '', {
                        **F_CLS,
                        **F_PETIT,
                        **F_DROITE,
                        **F_UNL,
                        **bordure, 'bg_color': C_CLS[div][0]
                    })
        ### tab récap du bas
        lignes_lv2 = [i * nbl + 3 for i in range(len(LV2S))]
        txt_somme = '=SUM(' + ','.join(
            [lig_col(l, col - 1) for l in lignes_lv2]) + ')'
        txt_somme2 = '=SUM(' + ','.join([lig_col(l, col)
                                         for l in lignes_lv2]) + ')'
        pat(
            dern + 1, col - 1, txt_somme, {
                **F_CLS,
                **F_GRAS,
                **F_MOYEN,
                **F_HB,
                **F_GAUCHE, 'bg_color': C_CLS[div][1]
            })
        pat(
            dern + 1, col, txt_somme2, {
                **F_CLS,
                **F_GRAS,
                **F_MOYEN,
                **F_HB,
                **F_DROITE, 'bg_color': C_CLS[div][1]
            })
        for i, _ in enumerate(OPTIONS):
            lignes_opt = [l + i + 1 for l in lignes_lv2]
            txt_somme_opt = '=SUM(' + ','.join(
                [lig_col(l, col - 1) for l in lignes_opt]) + ')'
            txt_somme_opt2 = '=SUM(' + ','.join(
                [lig_col(l, col) for l in lignes_opt]) + ')'
            bordure = F_BAS if i == len(OPTIONS) - 1 else {}
            pat(
                dern + i + 2, col - 1, txt_somme_opt, {
                    **F_CLS,
                    **F_PETIT,
                    **F_GAUCHE, 'bg_color': C_CLS[div][0],
                    **bordure
                })
            pat(
                dern + i + 2, col, txt_somme_opt2, {
                    **F_CLS,
                    **F_PETIT,
                    **F_DROITE, 'bg_color': C_CLS[div][0],
                    **bordure
                })
        for i, opt in enumerate(OPTIONS_UNIQUES):
            lignes_opt = []
            for j, option in enumerate(OPTIONS):
                if opt in OPTIONS[option]:
                    lignes_opt.append(j)
            txt = '=SUM(' + ','.join(
                [lig_col(dern + 2 + l, col - 1) for l in lignes_opt]) + ')'
            txt2 = '=SUM(' + ','.join(
                [lig_col(dern + 2 + l, col) for l in lignes_opt]) + ')'
            bordure = {}
            if i == 0: bordure = F_HAUT
            if i == len(OPTIONS_UNIQUES) - 1: bordure = {**bordure, **F_BAS}
            pat(
                dern2 + i, col - 1, txt, {
                    **F_CLS,
                    **F_PETIT,
                    **F_GAUCHE, 'bg_color': C_CLS[div][0],
                    **bordure
                })
            pat(
                dern2 + i, col, txt2, {
                    **F_CLS,
                    **F_PETIT,
                    **F_DROITE, 'bg_color': C_CLS[div][0],
                    **bordure
                })
###

# dernière catégorie : NA
    pat(1, 2 * NB_DIVS + 1, TX_YA, F_YA)
    pat(2, 2 * NB_DIVS + 1, 'NA', {
        **F_CLS,
        **F_GRAS,
        **F_GROS,
        **F_BORD, 'bg_color': C_CLS[-1][0]
    })
    col = 2 * NB_DIVS + 1
    for i, lv2 in enumerate(LV2S):  # boucles sur les différentes LV2
        lig = i * nbl + 3
        if i == len(LV2S) - 1:
            pat(lig, col,
                '=COUNTIFS(_Classe,' + lig_col(2, col) + ',' + sans_lv2 + ')',
                {
                    **F_CLS,
                    **F_GRAS,
                    **F_MOYEN,
                    **F_BORD, 'bg_color': C_CLS[-1][1]
                })
        else:
            pat(lig, col, '=COUNTIFS(_Classe,' + lig_col(2, col) + ',' +
                nom_lv2[i] + ',1)', {
                    **F_CLS,
                    **F_GRAS,
                    **F_MOYEN,
                    **F_BORD, 'bg_color': C_CLS[-1][1]
                })
        for j, opt in enumerate(OPTIONS):  # boucles sur les options
            bordure = {}
            if j == 0: bordure = F_HAUT
            if j == len(OPTIONS) - 1: bordure = {**bordure, **F_BAS}
            cond = []
            for option in OPTIONS_UNIQUES:
                nosp = '_' + ''.join(option.split())
                cond.append(nosp)
                if option in OPTIONS[opt]: cond.append('1')
                else: cond.append('""')
                txt_cond = ','.join(cond)
            if i == len(LV2S) - 1:
                pat(
                    lig + j + 1, col, '=COUNTIFS(_Classe,' + lig_col(2, col) +
                    ',' + sans_lv2 + ',' + txt_cond + ')', {
                        **F_CLS,
                        **F_PETIT,
                        **F_COTES,
                        **bordure, 'bg_color': C_CLS[-1][0]
                    })
            else:
                pat(
                    lig + j + 1, col, '=COUNTIFS(_Classe,' + lig_col(2, col) +
                    ',' + nom_lv2[i] + ',1,' + txt_cond + ')', {
                        **F_CLS,
                        **F_PETIT,
                        **F_COTES,
                        **bordure, 'bg_color': C_CLS[-1][0]
                    })
    lignes_lv2 = [i * nbl + 3 for i in range(len(LV2S))]
    txt_somme = '=SUM(' + ','.join([lig_col(l, col) for l in lignes_lv2]) + ')'
    pat(dern + 1, col, txt_somme, {
        **F_CLS,
        **F_GRAS,
        **F_MOYEN,
        **F_BORD, 'bg_color': C_CLS[-1][1]
    })
    for i, _ in enumerate(OPTIONS):
        lignes_opt = [l + i + 1 for l in lignes_lv2]
        txt_somme_opt = '=SUM(' + ','.join(
            [lig_col(l, col) for l in lignes_opt]) + ')'
        bordure = F_BAS if i == len(OPTIONS) - 1 else {}
        pat(dern + i + 2, col, txt_somme_opt, {
            **F_CLS,
            **F_PETIT,
            **F_COTES, 'bg_color': C_CLS[-1][0],
            **bordure
        })
    for i, opt in enumerate(OPTIONS_UNIQUES):
        lignes_opt = []
        for j, option in enumerate(OPTIONS):
            if opt in OPTIONS[option]:
                lignes_opt.append(j)
        txt = '=SUM(' + ','.join(
            [lig_col(dern + 2 + l, col) for l in lignes_opt]) + ')'
        bordure = {}
        if i == 0: bordure = F_HAUT
        if i == len(OPTIONS_UNIQUES) - 1: bordure = {**bordure, **F_BAS}
        pat(dern2 + i, col, txt, {
            **F_CLS,
            **F_PETIT,
            **F_COTES, 'bg_color': C_CLS[-1][0],
            **bordure
        })
###
# Totaux
    pat(1, 2 * NB_DIVS + 2, TX_YA, F_YA)
    pat(1, 2 * NB_DIVS + 3, TX_FAUT, {
        **F_FAUT, 'italic': False,
        'bold': False
    })
    pat(1, 2 * NB_DIVS + 4, 'Liste', F_LST)
    pat_merge(2, 2 * NB_DIVS + 2, 2, 2 * NB_DIVS + 4, 'TOTAUX', {
        **F_TOTAUX,
        **F_BORD,
        **F_GROS,
    })
    col = 2 * NB_DIVS + 4
    for i, lv2 in enumerate(LV2S):  # boucles sur les différentes LV2
        lig = i * nbl + 3

        if i == len(LV2S) - 1:
            pat(lig, col, '=COUNTIFS(' + sans_lv2 + ')', {
                **F_TOTAUX2,
                **F_GRAS,
                **F_MOYEN,
                **F_HB,
                **F_DROITE,
            })
        else:
            pat(lig, col, '=COUNTIFS(' + nom_lv2[i] + ',1)', {
                **F_TOTAUX2,
                **F_GRAS,
                **F_MOYEN,
                **F_HB,
                **F_DROITE,
            })
        for j, opt in enumerate(OPTIONS):  # boucles sur les options
            bordure = {}
            if j == 0: bordure = F_HAUT
            if j == len(OPTIONS) - 1: bordure = {**bordure, **F_BAS}
            cond = []
            for option in OPTIONS_UNIQUES:
                nosp = '_' + ''.join(option.split())
                cond.append(nosp)
                if option in OPTIONS[opt]: cond.append('1')
                else: cond.append('""')
                txt_cond = ','.join(cond)
            if i == len(LV2S) - 1:
                pat(lig + j + 1, col,
                    '=COUNTIFS(' + sans_lv2 + ',' + txt_cond + ')', {
                        **F_TOTAUX,
                        **F_PETIT,
                        **F_DROITE,
                        **bordure,
                    })
            else:
                pat(lig + j + 1, col,
                    '=COUNTIFS(' + nom_lv2[i] + ',1,' + txt_cond + ')', {
                        **F_TOTAUX,
                        **F_PETIT,
                        **F_DROITE,
                        **bordure,
                    })
###
    lignes_lv2 = [i * nbl + 3 for i in range(len(LV2S))]
    txt_somme = '=SUM(' + ','.join([lig_col(l, col - 2)
                                    for l in lignes_lv2]) + ')'
    txt_somme2 = '=SUM(' + ','.join([lig_col(l, col - 1)
                                     for l in lignes_lv2]) + ')'
    txt_somme3 = '=SUM(' + ','.join([lig_col(l, col)
                                     for l in lignes_lv2]) + ')'

    pat(dern + 1, col - 2, txt_somme, {
        **F_TOTAUX2,
        **F_GRAS,
        **F_MOYEN,
        **F_HB,
        **F_GAUCHE
    })
    pat(dern + 1, col - 1, txt_somme2, {
        **F_TOTAUX2,
        **F_GRAS,
        **F_MOYEN,
        **F_HB
    })
    pat(dern + 1, col, txt_somme3, {
        **F_TOTAUX2,
        **F_GRAS,
        **F_MOYEN,
        **F_HB,
        **F_DROITE
    })

    for i, _ in enumerate(OPTIONS):
        lignes_opt = [l + i + 1 for l in lignes_lv2]
        txt_somme_opt = '=SUM(' + ','.join(
            [lig_col(l, col - 2) for l in lignes_opt]) + ')'
        txt_somme_opt2 = '=SUM(' + ','.join(
            [lig_col(l, col - 1) for l in lignes_opt]) + ')'
        txt_somme_opt3 = '=SUM(' + ','.join(
            [lig_col(l, col) for l in lignes_opt]) + ')'
        bordure = F_BAS if i == len(OPTIONS) - 1 else {}
        pat(dern + i + 2, col - 2, txt_somme_opt, {
            **F_TOTAUX,
            **F_PETIT,
            **F_GAUCHE,
            **bordure
        })
        pat(dern + i + 2, col - 1, txt_somme_opt2, {
            **F_TOTAUX,
            **F_PETIT,
            **bordure
        })
        pat(dern + i + 2, col, txt_somme_opt3, {
            **F_TOTAUX,
            **F_PETIT,
            **F_DROITE,
            **bordure
        })

    for i, opt in enumerate(OPTIONS_UNIQUES):
        lignes_opt = []
        for j, option in enumerate(OPTIONS):
            if opt in OPTIONS[option]:
                lignes_opt.append(j)
        txt = '=SUM(' + ','.join(
            [lig_col(dern + 2 + l, col - 2) for l in lignes_opt]) + ')'
        txt2 = '=SUM(' + ','.join(
            [lig_col(dern + 2 + l, col - 1) for l in lignes_opt]) + ')'
        txt3 = '=SUM(' + ','.join(
            [lig_col(dern + 2 + l, col) for l in lignes_opt]) + ')'
        bordure = {}
        if i == 0: bordure = F_HAUT
        if i == len(OPTIONS_UNIQUES) - 1: bordure = {**bordure, **F_BAS}
        pat(dern2 + i, col - 2, txt, {
            **F_TOTAUX,
            **F_PETIT,
            **F_GAUCHE,
            **bordure
        })
        pat(dern2 + i, col - 1, txt2, {**F_TOTAUX, **F_PETIT, **bordure})
        pat(dern2 + i, col, txt3, {
            **F_TOTAUX,
            **F_PETIT,
            **F_DROITE,
            **bordure
        })
    #### formatage conditionnel
    for div in range(NB_DIVS):
        col = 2 * div + 1
        patates.conditional_format(
            3, col, dern2 + len(OPTIONS_UNIQUES) - 1, col, {
                'type': 'cell',
                'criteria': 'not equal to',
                'value': '=' + lig_col(3, col + 1),
                'format': workbook.add_format(F_ERR)
            })
    patates.conditional_format(
        3, 2 * NB_DIVS + 2, dern2 + len(OPTIONS_UNIQUES) - 1, 2 * NB_DIVS + 2,
        {
            'type': 'cell',
            'criteria': 'not equal to',
            'value': '=' + lig_col(3, 2 * NB_DIVS + 3),
            'format': workbook.add_format(F_ERR)
        })
    patates.conditional_format(
        3, 2 * NB_DIVS + 3, dern2 + len(OPTIONS_UNIQUES) - 1, 2 * NB_DIVS + 3,
        {
            'type': 'cell',
            'criteria': 'not equal to',
            'value': '=' + lig_col(3, 2 * NB_DIVS + 4),
            'format': workbook.add_format(F_ERRP)
        })

    lignes_lv2 = [i * nbl + 3 for i in range(len(LV2S))] + [dern + 1]
    for div in range(NB_DIVS):
        col = 2 * div + 1
        for j in lignes_lv2:
            patates.conditional_format(
                j, col, j, col + 1, {
                    'type': 'cell',
                    'criteria': 'equal to',
                    'value': 0,
                    'format': workbook.add_format({
                        'font_color': C_CLS[div][1]
                    })
                })
        patates.conditional_format(
            3, col, dern2 + len(OPTIONS_UNIQUES) - 1, col, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][0]
                })
            })
        patates.conditional_format(
            dern + 1, col + 1, dern2 + len(OPTIONS_UNIQUES) - 1, col + 1, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[div][0]
                })
            })
    #NA
    col = 2 * NB_DIVS + 1
    for j in lignes_lv2:
        patates.conditional_format(
            j, col, j, col, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CLS[-1][1]
                })
            })
    patates.conditional_format(
        3, col, dern2 + len(OPTIONS_UNIQUES) - 1, col, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CLS[-1][0]
            })
        })
    # Totaux
    col = 2 * NB_DIVS + 2
    for j in lignes_lv2:
        patates.conditional_format(
            j, col, j, col + 2, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': 0,
                'format': workbook.add_format({
                    'font_color': C_CAT['TOT2'][1]
                })
            })
    patates.conditional_format(
        3, col, dern2 + len(OPTIONS_UNIQUES) - 1, col + 2, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': workbook.add_format({
                'font_color': C_CAT['TOT'][1]
            })
        })

    # Création des formules dans les colonnes
    for i, lv2 in enumerate(LV2S):  # sommes dans les totaux
        lig = i * nbl + 3
        somme = '=SUM(' + ','.join(
            [lig_col(lig, 2 * div + 1) for div in range(NB_DIVS + 1)]) + ')'
        somme2 = '=SUM(' + ','.join(
            [lig_col(lig, 2 * div + 2) for div in range(NB_DIVS)]) + ')'
        pat(lig, 2 * NB_DIVS + 2, somme, {
            **F_TOTAUX2,
            **F_MOYEN,
            **F_HB,
            **F_GAUCHE,
        })
        pat(lig, 2 * NB_DIVS + 3, somme2, {
            **F_TOTAUX2,
            **F_MOYEN,
            **F_HB,
        })

        for j, opt in enumerate(OPTIONS):  # lignes des options
            somme = '=SUM(' + ','.join([
                lig_col(lig + j + 1, 2 * div + 1)
                for div in range(NB_DIVS + 1)
            ]) + ')'
            somme2 = '=SUM(' + ','.join(
                [lig_col(lig + j + 1, 2 * div + 2)
                 for div in range(NB_DIVS)]) + ')'

            bordure = {}
            if j == 0: bordure = F_HAUT
            if j == len(OPTIONS) - 1: bordure = {**bordure, **F_BAS}
            pat(lig + j + 1, 2 * NB_DIVS + 2, somme, {
                **F_TOTAUX,
                **F_PETIT,
                **F_GAUCHE,
                **bordure,
            })
            pat(lig + j + 1, 2 * NB_DIVS + 3, somme2, {
                **F_TOTAUX,
                **F_PETIT,
                **bordure,
            })

    # Création de la première colonne
    for i, lv2 in enumerate(LV2S):
        pat(i * nbl + 3, 0, lv2, F_LV2)
        for j, opt in enumerate(OPTIONS):
            bordure = F_BAS if i == len(LV2S) - 1 and j == len(
                OPTIONS) - 1 else {}
            pat(i * nbl + j + 4, 0, opt, {**F_OPT, **bordure})
            patates.set_row(i * nbl + j + 4, None, None, {'level': 1})
    patates.set_row(dern, 10)
    pat(dern + 1, 0, 'Effectifs', {**F_LV2, **F_MOYEN})
    for i, opt in enumerate(OPTIONS):
        bordure = F_BAS if i == len(OPTIONS) - 1 else {}
        pat(dern + i + 2, 0, opt, {**F_OPT, **bordure})
        patates.set_row(dern + i + 2, None, None, {'level': 1})

    for i, opt in enumerate(OPTIONS_UNIQUES):
        bordure = {}
        if i == 0: bordure = F_HAUT
        if i == len(OPTIONS_UNIQUES) - 1: bordure = {**bordure, **F_BAS}
        pat(dern2 + i, 0, opt, {**F_OPT, **bordure})
        patates.set_row(dern2 + i, None, None, {'level': 1})

    patates.freeze_panes(3, 1)
    liste.freeze_panes(premier_el, 0)

    # zones d'impression
    patates.set_landscape()
    patates.set_paper(9)  # A4
    patates.center_horizontally()
    patates.center_vertically()
    patates.set_margins(0.4, 0.4, 0.4, 0.4)
    patates.set_header('', {'margin': 0})
    patates.set_footer('', {'margin': 0})
    patates.hide_gridlines()
    patates.print_area(0, 0, dern2 + len(OPTIONS_UNIQUES), 2 * NB_DIVS + 4)
    patates.fit_to_pages(1, 1)

    liste.set_landscape()
    liste.set_paper(9)  # A4
    liste.center_horizontally()
    #liste.center_vertically()
    liste.set_margins(0.4, 0.4, 0.4, 0.4)
    liste.set_header('', {'margin': 0})
    liste.set_footer('', {'margin': 0})
    liste.hide_gridlines()
    liste.repeat_rows(premier_el - 1)
    liste.print_area(0, 0, dernier_el,
                     max(
                         len(LV2S_VRAIES) + len(OPTIONS_UNIQUES) + len(OPTIONS_CAT)+8,
                         4 + NB_DIVS))
    liste.fit_to_pages(1, 0)
    liste.set_h_pagebreaks(
        [premier_el - 1] +
        [premier_el + (i + 1) * NB_ELV // NB_DIVS for i in range(NB_DIVS)])
    liste.set_v_pagebreaks(
        [max(len(LV2S_VRAIES) + len(OPTIONS_UNIQUES) + len(OPTIONS_CAT)+8, 4 + NB_DIVS) + 1])

    patates.protect()
    liste.protect(
        options={
            'insert_rows': True,
            'delete_rows': True,
            'sort': True,
            'autofilter': True,
        })

from IPython.display import HTML, display
if jupyter():
    display(
        HTML(
            '<p>Création du fichier terminée. Cliquer le lien ci-dessous pour le télécharger.</p><hr><h1><a href="./'
            + NOM_FICHIER +
            '" target="_blank">Lien vers le fichier</a></h1><br>' +
            '<h2 align="center">Pour se faire un peu de place dans Excel :</h2><img src="xl_place.png">'
        ))
else:
    import os
    os.startfile(NOM_FICHIER)
    print("Création du fichier '{}'".format(NOM_FICHIER))
    print("Fini !")
