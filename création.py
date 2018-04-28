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
    nom_liste = CLASSES + ' ' + str(2000 + ANNEE) + '-' + str(ANNEE + 1)
    liste = workbook.add_worksheet(nom_liste)
    liste.outline_settings(symbols_below=False)

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

    # Colonne NA
    rep(0, 2 + NB_DIVS, 'NA', {
        **F_CLS,
        **F_GRAS,
        **F_GROS,
        **F_BORD, 'bg_color': C_CLS[-1][0]
    })
    rep(1, 2 + NB_DIVS, '=COUNTIF(_Classe,"NA")', {
        **F_CLS,
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

    # Colonne Reste
    rep(0, 3 + NB_DIVS, 'Reste', {
        **F_RESTE1,
        **F_GRAS,
        **F_BORD,
    })
    rep(1, 3 + NB_DIVS, '=COUNTBLANK(_Classe)', {
        **F_RESTE2,
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
            **F_UNL
        })
        if opt in OPTIONS_CAT:
            compt += 1
            liste.set_column(col, col, 6)
            liste.set_column(col + 1, col + 1, 15)
            rep(dern_recap + 1, col + 1, OPTIONS_CAT[opt], {
                **F_OPT3[i % 3],
                **F_GRAS,
                **F_HB,
                **F_UNL
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

    ### TESTS ###
    # insertion de données de test
    test_nom = [
        'BARTH', 'DEHRI', 'LOISEAU', 'NOEL', 'RUBASZEWSKI', 'TILLIET',
        'VANHAREN'
    ]
    test_prenom = {
        'Jérôme': 'G',
        'Mohammed': 'G',
        'Franck': 'G',
        'Fanny': 'F',
        'Julie': 'F',
        'Céline': 'F',
        'Marie Louis': 'G',
    }
    test_retard = [''] * 10 + ['R']
    test_div = ['4e' + str(i + 1) for i in range(8)] + ['Nv3e']
    test_lv2 = [[1, '', ''], ['', 1, ''], ['', '', 1]] * 5 + [['', '', '']]
    test_sp = [['', '']] * 10 + [[1, 'FOOTBALL'], [1, 'GYMNASTIQUE'],
                                 [1, 'BASKET']]
    test_lat = [''] * 10 + [1]
    test_obs = [''] * 5 + ['Teigne', 'Peste', 'Bon en maths']

    from random import choice
    for el in range(NB_ELV):
        nom = choice(test_nom)
        prenom, sexe = choice(list(test_prenom.items()))
        retard = choice(test_retard)
        niv = choice(NIVEAUX)
        comp = choice(NIVEAUX)
        classe = choice(NOM_DIVS + [''])
        div_orig = choice(test_div)
        lv2 = choice(test_lv2)
        sport = choice(test_sp)
        lat = choice(test_lat)
        obs = choice(test_obs)
        rep(premier_el + el, 0, nom, {
            **F_DEF,
            **F_GAUCHE,
            **F_UNL, 'align': 'left'
        })
        rep(premier_el + el, 1, prenom, {**F_DEF, **F_UNL, 'align': 'left'})
        rep(premier_el + el, 2, sexe, {**F_DEF, **F_UNL})
        liste.write_row(
            premier_el + el, 2,
            [sexe, retard, niv, comp, classe, div_orig] + lv2 + sport + [lat],
            workbook.add_format({
                **F_DEF,
                **F_UNL
            }))
        rep(premier_el + el, 14, obs, {
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
    liste.freeze_panes(premier_el,0)

    if not DEBUG:
        patates.protect()
        liste.protect(
            options={
                'insert_rows': True,
                'delete_rows': True,
                'sort': True,
                'autofilter': True,
            })
    if DEBUG: liste.activate()


def jupyter():
    try:
        shell = get_ipython().__class__.__name__
        return shell == 'ZMQInteractiveShell'
    except NameError:
        return False


from IPython.display import HTML, display
if jupyter():
    display(
        HTML('<a href="./' + NOM_FICHIER +
             '" target="_blank">Lien vers le fichier</a>'))
else:
    import os
    os.startfile(NOM_FICHIER)
    print("Création du fichier '{}'".format(NOM_FICHIER))
    print("Fini !")
