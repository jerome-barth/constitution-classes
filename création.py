# coding: utf-8
from config import *
from formats import *

import xlsxwriter
from datetime import datetime


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
    if not DEBUG:
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
    liste = workbook.add_worksheet(
        CLASSES + ' ' + str(2000 + ANNEE) + '-' + str(ANNEE + 1))
    liste.outline_settings(symbols_below=False)

    ##############################
    ###                        ###
    ###  Feuille 'Xe-20XX-XX'  ###
    ###                        ###
    ##############################
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
    rep(3, 1, 'G', {**F_G, **F_GRAS, **F_COTES})
    rep(4, 1, '%F', {**F_PF, **F_BAS, **F_COTES})
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
    lig_opt = 6 + len(NIVEAUX)
    rep(lig_opt - 1, 1, 'R', {**F_R, **F_GRAS, **F_BORD})
    # Tab récap étiquettes options
    for i, opt in enumerate(OPTIONS_UNIQUES):
        bordure = {**F_COTES}
        if i == 0: bordure = {**bordure, **F_HAUT}
        if i == len(OPTIONS_UNIQUES) - 1: bordure = {**bordure, **F_BAS}
        rep(lig_opt + i, 1, opt, {**F_OPT3[i % 3], **F_GRAS, **bordure})
        liste.set_row(lig_opt + i, None, None, {'level': 1})
    # Tab récap étiquettes LV2
    lig_lv2 = lig_opt + len(OPTIONS_UNIQUES)
    for i, lv2 in enumerate(LV2S):
        bordure = {**F_COTES}
        if i == 0: bordure = {**bordure, **F_HAUT}
        if i == len(LV2S) - 1: bordure = {**bordure, **F_BAS}
        rep(lig_lv2 + i, 1, lv2, {**F_LV, **F_GRAS, **bordure})
        if i != 0: liste.set_row(lig_lv2 + i, None, None, {'level': 1})
    # Dernière ligne avant la liste
    dern_recap = lig_lv2 + len(LV2S)
    liste.set_row(dern_recap, 5)
    for i in range(NB_DIVS + 4):
        rep(dern_recap, i + 1, None, F_HAUT)
    # Tableau récap : colonnes des classes
    for div, nom_div in enumerate(NOM_DIVS):
        bordure = {}
        if div == 0: bordure = {**F_GAUCHE}
        if div == NB_DIVS - 1: bordure = {**bordure, **F_HB, **F_DROITE}
        rep(
            0, 2 + div, nom_div, {
                **F_CLS,
                **F_GRAS,
                **F_GROS,
                **F_HB,
                **bordure, 'bg_color': C_CLS[div][0]
            })
    rep(0, 2 + NB_DIVS, 'NA', {
        **F_CLS,
        **F_GRAS,
        **F_GROS,
        **F_BORD, 'bg_color': C_CLS[-1][0]
    })
    rep(0, 3 + NB_DIVS, 'Reste', {
        **F_RESTE1,
        **F_GRAS,
        **F_BORD,
    })
    rep(0, 4 + NB_DIVS, 'Totaux', {
        **F_TOTAUX,
        **F_BORD,
    })
    # -~- Liste -~-
    # Entête
    rep(dern_recap + 1, 0, 'Nom', {**F_ENT, **F_HB, **F_GAUCHE})
    fmt = workbook.add_format({
        **F_ENT,
        **F_HB,
    })
    liste.write_row(
        dern_recap + 1, 1,
        ['Prénom', 'Sexe', 'Retard', 'Niv', 'Comp', CLASSES, 'Cls orig'], fmt)
    for i, lv2 in enumerate(LV2S_VRAIES):
        rep(dern_recap + 1, 8 + i, lv2, {**F_ENT, **F_HB})
    compt = 0
    for i, opt in enumerate(OPTIONS_UNIQUES):
        col = 8 + len(LV2S_VRAIES) + i + compt
        rep(dern_recap + 1, col, opt, {**F_ENT, **F_HB})
        if opt in OPTIONS_CAT:
            compt += 1
            liste.set_column(col, col, 6)
            liste.set_column(col + 1, col + 1, 15)
            rep(dern_recap + 1, col + 1, OPTIONS_CAT[opt], {**F_ENT, **F_HB})
        else:
            liste.set_column(col, col, 6)
    col = 8 + len(LV2S_VRAIES) + compt + len(OPTIONS_UNIQUES)
    liste.set_column(col, col, 25)
    rep(dern_recap + 1, col, 'Observations', {**F_ENT, **F_HB, **F_DROITE})
    ###TESTS###

    if DEBUG:
        donnees = [[
            'AUBLE', 'Ayyoub', 'G', '', 'A', 'A', '', '5E2', 1, '', '', 1,
            'GYMNASTIQUE G', '', 'Observation'
        ], [
            'BELMELIANI', 'Elias', 'G', '', 'B', 'B', '', '5E2', '', 1, '', '',
            '', 1, 'Obs'
        ]]
        for i, ligne in enumerate(donnees):
            liste.write_row(dern_recap + 5 + i, 0, ligne)
    ###########################
    ###                     ###
    ###  Feuille 'Patates'  ###
    ###                     ###
    ###########################

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
    dern = (len(LV2S) + 1) * (len(OPTIONS) + 1) - 2
    # N° de la ligne après la liste des options
    dern2 = dern + len(OPTIONS) + 2

    for div, nom_div in enumerate(NOM_DIVS):  # boucle sur les classes
        pat(1, 2 * div + 1, TX_YA, F_YA)
        pat(1, 2 * div + 2, TX_FAUT, F_FAUT)
        pat_merge(2, 2 * div + 1, 2, 2 * div + 2, nom_div, {
            **F_CLS,
            **F_GRAS,
            **F_GROS,
            **F_BORD, 'bg_color': C_CLS[div][0]
        })
        for i, lv2 in enumerate(LV2S):  # boucles sur les différentes LV2
            lig = i * nbl + 3
            col = 2 * div + 2
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
                pat(
                    lig + j + 1, col, '*', {
                        **F_CLS,
                        **F_PETIT,
                        **F_DROITE,
                        **F_UNL,
                        **bordure, 'bg_color': C_CLS[div][0]
                    })

    # dernière catégorie : NA
    pat(1, 2 * NB_DIVS + 1, TX_YA, F_YA)
    pat(2, 2 * NB_DIVS + 1, 'NA', {
        **F_CLS,
        **F_GRAS,
        **F_BORD, 'bg_color': C_CLS[-1][0]
    })

    # Totaux
    pat(1, 2 * NB_DIVS + 2, TX_YA, F_YA)
    pat(1, 2 * NB_DIVS + 3, TX_FAUT, F_FAUT)
    pat(1, 2 * NB_DIVS + 4, 'Liste', F_LST)
    pat_merge(2, 2 * NB_DIVS + 2, 2, 2 * NB_DIVS + 4, 'TOTAUX', {
        **F_TOTAUX,
        **F_BORD,
        **F_GROS,
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

    pat(dern + 1, 0, 'Effectifs', F_LV2)
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

    #     patates.protect()
    #     liste.protect()
    liste.activate()


def jupyter():
    try:
        shell = get_ipython().__class__.__name__
        return shell == 'ZMQInteractiveShell'
    except NameError:
        return False


if jupyter():
    display(
        HTML('<a href="./' + NOM_FICHIER +
             '" target="_blank">Lien vers le fichier</a>'))
else:
    import os
    os.startfile(NOM_FICHIER)
    print("Fini !")
