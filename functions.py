import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from datetime import datetime, time, date


def getNomsSalles():
    df = pd.read_csv("/parameters/salles.csv")
    noms = []
    for d in df.iterrows():
        noms.append(d[1][0].upper())
    return noms


def getColorsBySalleName(nom_salle):
    df = pd.read_csv("/parameters/salles.csv")
    for d in df.iterrows():
        if d[1][0].upper() == nom_salle.upper():
            splitted_bcolor = d[1][1].split("/")
            splitted_fcolor = d[1][2].split("/")
            bcolor = colors.Color(red=(float(splitted_bcolor[0]) / 255), green=(float(splitted_bcolor[1]) / 255),
                                  blue=(float(splitted_bcolor[2]) / 255))
            fcolor = colors.Color(red=(float(splitted_fcolor[0]) / 255), green=(float(splitted_fcolor[1]) / 255),
                                  blue=(float(splitted_fcolor[2]) / 255))
            return bcolor, fcolor


def genPlanningEtr(input_file, with_title, with_srm, all_days, jours_entrainements, salles, sous_titre, title):
    saison = sous_titre
    canva = input_file
    df = pd.read_excel(canva)
    bcouleurs_salles = []
    fcouleurs_salles = []
    for salle in salles:
        bcouleur, fcouleur = getColorsBySalleName(salle)
        bcouleurs_salles.append(bcouleur)
        fcouleurs_salles.append(fcouleur)
    dic_categorie = {}
    dic_salle = {}

    for d in df.iterrows():
        # d[1][0] : Salle
        # d[1][1] : Jour
        # d[1][2] : heure début
        # d[1][3] : heure fin
        # d[1][4] : catégorie

        # Scinder par catégorie
        dic_temp = {}
        dic_temp["salle"] = d[1][0].upper()
        dic_temp["jour"] = d[1][1].upper()
        dic_temp["heure_deb"] = d[1][2]
        dic_temp["heure_fin"] = d[1][3]
        for cat in d[1][4].split("/"):
            if cat.upper() in dic_categorie:
                dic_categorie[cat.upper()].append(dic_temp)
            else:
                dic_categorie[cat.upper()] = [dic_temp]
        # Scinder par salle/jour
        dic_temp = {}
        dic_temp["categorie"] = d[1][4].upper()
        dic_temp["heure_fin"] = d[1][3]
        dic_temp["heure_deb"] = d[1][2]
        if d[1][0].upper() in dic_salle:
            if d[1][1].upper() in dic_salle[d[1][0].upper()]:
                dic_salle[d[1][0].upper()][d[1][1].upper()].append(dic_temp)
            else:
                dic_salle[d[1][0].upper()][d[1][1].upper()] = [dic_temp]
        else:
            dic_salle[d[1][0].upper()] = {}
            dic_salle[d[1][0].upper()][d[1][1].upper()] = [dic_temp]

    # Fichier planning par Equipe
    entete = ["ÉQUIPES"]
    for j in jours_entrainements:
        entete.append(j)
    tab = [entete]
    tab_sty = [
        ('ALIGN', (0, 0), (-1, -1), 'CENTRE'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('SIZE', (0, 0), (-1, -1), 7),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('FONT', (0, 0), (-1, 0), "Helvetica-Bold"),
        ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ('FONT', (0, 0), (0, -1), "Helvetica-Bold")]
    for i, equipe in enumerate(sorted(dic_categorie)):
        row = [equipe]
        for j in jours_entrainements:
            row.append("")
        ids_col = []
        stored_salles = []
        for creneau in dic_categorie[equipe]:
            if creneau["salle"] in salles and creneau["jour"] in jours_entrainements:
                str_to_print = creneau["heure_deb"].strftime("%HH%M") + " - " + creneau["heure_fin"].strftime(
                    "%HH%M") + "\n" + \
                               creneau["salle"]
                row[jours_entrainements.index(creneau["jour"]) + 1] = str_to_print
                ids_col.append(jours_entrainements.index(creneau["jour"]) + 1)
                stored_salles.append(creneau["salle"])
        tab.append(row)
        for j, id in enumerate(ids_col):
            tab_sty.append(('BACKGROUND', (id, i + 1), (id, i + 1), bcouleurs_salles[salles.index(stored_salles[j])]))
            tab_sty.append(('TEXTCOLOR', (id, i + 1), (id, i + 1), fcouleurs_salles[salles.index(stored_salles[j])]))
    t = Table(tab, colWidths='*')
    sty = TableStyle(tab_sty)
    t.setStyle(sty)
    elements = []
    if with_title:
        sty = ParagraphStyle("style", fontSize=13, alignment=1, spaceAfter=5)
        p = Paragraph(title, style=sty)
        elements.append(p)
        sty = ParagraphStyle("style", fontSize=11, alignment=1, spaceAfter=3)
        p = Paragraph(saison, style=sty)
        elements.append(p)
    sty = ParagraphStyle("style", fontSize=5, alignment=1, spaceAfter=8)
    if with_srm:
        p = Paragraph("*Sous réserve de modifications (version du " + datetime.now().strftime("%d/%m/%Y") + ")",
                      style=sty)
        elements.append(p)
    elements.append(t)
    # write the document to disk

    doc = SimpleDocTemplate("/outputs/planning_entrainements_parEquipe.pdf", title="Planning des entrainements",
                            pagesize=A4,
                            rightMargin=5, leftMargin=5, topMargin=0, bottomMargin=0)
    doc.build(elements)

    # Fichier par Salle
    elements = []
    nb_pages = 1
    if with_title:
        sty = ParagraphStyle("style", fontSize=12, alignment=1, spaceAfter=5)
        p = Paragraph(title, style=sty)
        elements.append(p)
        sty = ParagraphStyle("style", fontSize=10, alignment=1, spaceAfter=3)
        p = Paragraph(saison, style=sty)
        elements.append(p)
    sty = ParagraphStyle("style", fontSize=5, alignment=1, spaceAfter=1)
    if with_srm:
        p = Paragraph("*Sous réserve de modifications (version du " + datetime.now().strftime("%d/%m/%Y") + ")",
                      style=sty)
        elements.append(p)

    for s, salle in enumerate(salles):
        if salle in dic_salle:
            sty = ParagraphStyle("style", fontSize=10, leading=16, alignment=1, spaceAfter=1, spaceBefore=6,
                                 backColor=bcouleurs_salles[s], textColor=fcouleurs_salles[s])
            p = Paragraph(salle, style=sty)
            elements.append(p)
            tab = []
            min_heure_deb = time(hour=23, minute=59)
            max_heure_fin = time(hour=0, minute=1)
            for jour in dic_salle[salle]:
                if all_days:
                    tab = jours_entrainements
                else:
                    tab.append(jour)
                for creneau in dic_salle[salle][jour]:
                    if creneau["heure_deb"] < min_heure_deb:
                        min_heure_deb = creneau["heure_deb"]
                    if creneau["heure_fin"] > max_heure_fin:
                        max_heure_fin = creneau["heure_fin"]
            d = date(1, 1, 1)
            datetime1 = datetime.combine(d, max_heure_fin)
            datetime2 = datetime.combine(d, min_heure_deb)
            time_elapsed = datetime1 - datetime2
            nb_row = time_elapsed.seconds / (5 * 60)
            tab = [tab]
            for i in range(0, int(nb_row)):
                tab_temp = []
                if all_days:
                    for jour in jours_entrainements:
                        tab_temp.append("")
                else:
                    for jour in dic_salle[salle]:
                        tab_temp.append("")
                tab.append(tab_temp)
            tab_style = [
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('SIZE', (0, 0), (-1, -1), 7),
                ('FONT', (0, 0), (-1, 0), "Helvetica-Bold"),
                ('GRID', (0, 0), (-1, 0), 0.25, colors.black),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                ('LINEBEFORE', (0, 0), (-1, -1), 0.25, colors.black)]
            for i, jour in enumerate(dic_salle[salle]):
                if jour in jours_entrainements:
                    if all_days:
                        i = jours_entrainements.index(jour)
                    for j, creneau in enumerate(dic_salle[salle][jour]):
                        datetime3 = datetime.combine(d, creneau["heure_deb"])
                        time_elapsed = datetime3 - datetime2
                        id_row = time_elapsed.seconds / (5 * 60)
                        str_to_print = creneau["categorie"] + "\n" + creneau["heure_deb"].strftime("%HH%M") + " - " + \
                                       creneau[
                                           "heure_fin"].strftime("%HH%M")
                        tab[int(id_row) + 1][i] = str_to_print
                        datetime4 = datetime.combine(d, creneau["heure_fin"])
                        time_elapsed = datetime4 - datetime3
                        delta_time = int(time_elapsed.seconds / (5 * 60))
                        tab_style.append(('SPAN', (i, int(id_row) + 1), (i, int(id_row) + delta_time)))
                        # tab_style.append(
                        #    ('BACKGROUND', (i, int(id_row) + 1), (i, int(id_row) + delta_time), couleurs_creneaux[j % 2]))
                        # tab_style.append(('GRID', (i, int(id_row) + 1), (i, int(id_row) + delta_time), 0.25, colors.black))
                        tab_style.append(('BOX', (i, int(id_row) + 1), (i, int(id_row) + delta_time), 2, colors.black))
            nrows = len(tab)
            rowHeights = nrows * [2]
            rowHeights[0] = 11
            if all_days:
                t = Table(tab, rowHeights=rowHeights, colWidths='*', hAlign='CENTER')
            else:
                t = Table(tab, rowHeights=rowHeights, hAlign='CENTER')
            sty = TableStyle(tab_style)
            t.setStyle(sty)
            elements.append(t)
    doc = SimpleDocTemplate("/outputs/planning_entrainements_parSalle.pdf", title="Planning des entrainements",
                            pagesize=A4,
                            rightMargin=10, leftMargin=10, topMargin=0, bottomMargin=0)
    doc.build(elements)
