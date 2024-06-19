from tkinter import filedialog, messagebox
from tkinter import *
import functions as f

root_file = ""
jours_entrainements = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]


def fileChoice():
    global root_file
    root = filedialog.askopenfile(initialdir="")
    root_file = root.name


def generatePlannings():
    global root_file
    if root_file.split('.')[-1]!="xlsx":
        messagebox.showerror("ERROR -- Mauvais fichier canva", "Veuillez sélectionner le bon fichier Canva au format .xlsx")
    else:
        jours = []
        for j in cbJoursVals:
            if cbJoursVals[j].get():
                jours.append(j)
        salles = []
        for s in cbSallesVals:
            if cbSallesVals[s].get():
                salles.append(s)
        f.genPlanningEtr(with_srm=varSRMGr.get(), with_title=vartitleSubGr.get(), input_file=root_file,
                         all_days=varsevenDaysGr.get(), salles=salles, jours_entrainements=jours,
                         sous_titre=sous_titre.get(), title=doc_titre.get())
        messagebox.showinfo('Plannings générés', "Planning générés dans le dossier outputs")

if __name__ == "__main__":
    win = Tk()
    win.title("Générateur de planning d'entrainements")
    frame_title = Frame(win)
    frame_title.pack(pady=10)
    Label(frame_title, text="Générateur de planning d'entrainements", font=("", 20)).pack()
    frame_file = Frame(win)
    frame_file.pack(pady=5)
    file_choice = Button(frame_file, text="Sélectionner le fichier Excel", command=fileChoice)
    file_choice.pack()
    frame_doctitle = Frame(win)
    frame_doctitle.pack(pady=5)
    Label(frame_doctitle, text="Titre :").pack(side=LEFT)
    doc_titre = Entry(frame_doctitle)
    doc_titre.insert(END, "Planning des entrainements")
    doc_titre.pack(side=LEFT)
    frame_subtitle = Frame(win)
    frame_subtitle.pack(pady=5)
    Label(frame_subtitle, text="Sous-Titre :").pack(side=LEFT)
    sous_titre = Entry(frame_subtitle)
    sous_titre.insert(END, "Saison 2024-2025")
    sous_titre.pack(side=LEFT)
    frame_srdm = Frame(win)
    frame_srdm.pack(pady=5)
    Label(frame_srdm, text="Afficher sous réserve de modifications + date ?").pack(side=LEFT)
    vals = [True, False]
    etiqs = ['Oui', 'Non']
    varSRMGr = BooleanVar()
    varSRMGr.set(vals[0])
    for i, e in enumerate(etiqs):
        b = Radiobutton(frame_srdm, variable=varSRMGr, text=e, value=vals[i])
        b.pack(side='left', expand=1)
    frame_titleSub = Frame(win)
    frame_titleSub.pack(pady=5)
    Label(frame_titleSub, text="Afficher titre + sous-titre ?").pack(side=LEFT)
    vartitleSubGr = BooleanVar()
    vartitleSubGr.set(vals[0])
    for i, e in enumerate(etiqs):
        b = Radiobutton(frame_titleSub, variable=vartitleSubGr, text=e, value=vals[i])
        b.pack(side='left', expand=1)
    frame_sevenDays = Frame(win)
    frame_sevenDays.pack(pady=5)
    Label(frame_sevenDays, text="Affichage des tous les jours sélectionnés pour chaque salles ?").pack(side=LEFT)
    varsevenDaysGr = BooleanVar()
    varsevenDaysGr.set(vals[0])
    for i, e in enumerate(etiqs):
        b = Radiobutton(frame_sevenDays, variable=varsevenDaysGr, text=e, value=vals[i])
        b.pack(side='left', expand=1)
    frame_salles = Frame(win)
    frame_salles.pack(pady=5)
    Label(frame_salles, text="Salles : ").pack(side=LEFT)
    cbSallesVals = {}
    for noms in f.getNomsSalles():
        cbSallesVals[noms] = BooleanVar()
        cb = Checkbutton(frame_salles, text=noms, variable=cbSallesVals[noms])
        cb.toggle()
        cb.pack(side=LEFT)

    frame_jours = Frame(win)
    frame_jours.pack(pady=5)
    Label(frame_jours, text="Jours : ").pack(side=LEFT)
    cbJoursVals = {}
    for i, noms in enumerate(jours_entrainements):
        cbJoursVals[noms] = BooleanVar()
        cb = Checkbutton(frame_jours, text=noms, variable=cbJoursVals[noms])
        if i != len(jours_entrainements) - 1:
            cb.toggle()
        cb.pack(side=LEFT)
    Button(win, text="Générer les plannings", command=generatePlannings).pack(pady=10)
    win.mainloop()
