import tkinter
from tkinter import *
from tkinter import messagebox
import sqlite3
from tkinter import ttk
import fileinput
import sys
import csv
from pprint import pprint
import numpy as np
import pandas as pd
import xlwings as xw
#import gspread
#from oauth2client.service_account import ServiceAccountCredentials

#wb = xw.Book(r'C:\Users\pasca\OneDrive\Compétitions GR\Résultats.xlsx') # uniquement quand le fichier excel est fermé

"""
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open("Résultats")
"""

# créeation des feuilles de résultats pour Google sheet (a faire 1x au début)
# worksheet = sh.add_worksheet("Résultats passages", rows="200", cols="11")
# worksheet = sh.add_worksheet("Résultats championnat", rows="200", cols="7")
# worksheet = sh.add_worksheet("Résultats coupe", rows="200", cols="5")

# création fenêtre principale
mainapp = tkinter.Tk()
mainapp.title("Tableau de bord")
mainapp.geometry("1080x720")
mainapp.minsize(480, 360)
mainapp.config(background='sandybrown')
mainapp.iconbitmap("Logo.ico")

# création image
long = 400
haut = 300
image = tkinter.PhotoImage(file="LogoGR1tr.png").zoom(40).subsample(30)
cadre = tkinter.Canvas(mainapp, width=long, height=haut, bg='sandybrown', bd=0, highlightthickness=0)
cadre.create_image(long / 2, haut / 2, image=image)
cadre.place(x=400, y=180)

def liste_passage():  # création fenêtre liste de passage complète
    fenetrepassage = tkinter.Toplevel()
    fenetrepassage.title("Liste de passage")
    fenetrepassage.geometry("750x330")
    fenetrepassage.config(background='yellow')
    fenetrepassage.iconbitmap("Logo.ico")

    def fermer_fenetre():
        fenetrepassage.destroy()

    def select_pass_gymn(event):  # sélectionne passage gymnaste
        # Assigner variables de la gymnaste sélectionnée depuis la fenêtre liste de passage complète
        item = tree.selection()
        nr_gymnaste_select = tree.item(item)['values'][0]
        nom_select = tree.item(item)['values'][1]
        prenom_select = tree.item(item)['values'][2]
        club_select = tree.item(item)['values'][3]
        cat_select = tree.item(item)['values'][4]
        engin_select = tree.item(item)['values'][5]

        def fermer_fenetre():
            fenetrenotes.destroy()

        def base_donnees():  # enregistrement des notes dans base de données

            noted1d2 = (d1d2.get())
            noted3d4 = (d3d4.get())
            notee1e2 = (e1e2.get())
            notee3 = (e3.get())
            notee4 = (e4.get())
            notee5 = (e5.get())
            notee6 = (e6.get())
            noteded = (ded.get())

            if noted1d2 == "":
                noted1d2 = 0;noted1d2= float(noted1d2)
            if noted3d4 == "":
                noted3d4 = 0;noted3d4= float(noted3d4)
            if notee1e2 == "":
                notee1e2= 0;notee1e2= float(notee1e2)
            if notee3 == "":
                notee3 = 0;notee3= float(notee3)
            if notee4 == "":
                notee4= 0;notee4= float(notee4)
            if notee5 == "":
                notee5 = 0;notee5= float(notee5)
            if notee6 == "":
                notee6 = 0;notee6= float(notee6)
            if noteded == "" :
                noteded = 0 ; noteded=float(noteded)


            # récupérer valeur départ note exe en indiv
            with open("depart_note exe_indiv.txt", 'r') as fichier:
                base_exe_indiv = fichier.readline()
                base_exe_indiv = int(base_exe_indiv)
                # récupérer valeur départ note exe en group
            with open("depart_note exe_group.txt", 'r') as fichier:
                base_exe_group = fichier.readline()
                base_exe_group = int(base_exe_group)

            connection = sqlite3.connect("BD GRS.db")
            cursor = connection.cursor()

            type_pass = (nr_gymnaste_select,)
            cursor.execute ("select * FROM PP_Inscriptions WHERE Nr_gymnaste = ?",type_pass)
            type_gymnaste = cursor.fetchone()[4]

            # calcul note exé individuelle
            if type_gymnaste == "I" :
                if  notee4 == 0 and notee5 == 0 and notee6 == 0:
                    noteE = base_exe_indiv - float(notee3) - float(notee1e2)
                if  notee4 != 0 and notee5 == 0 and notee6 == 0:
                    noteE= base_exe_indiv - ((float(notee3) + float(notee4)) /2) - float(notee1e2)
                if  notee4 != 0 and notee5 != 0 and notee6 == 0:
                    noteE = float(base_exe_indiv)- ((float(notee3) + float(notee4) + float(notee5)) / 3) - float(notee1e2)
                if  notee4 != 0 and notee5 != 0 and notee6 != 0 :
                    liste_notes = [notee3, notee4, notee5, notee6] # créer liste des 4 notes exé
                    liste_notes.sort() # trier liste par valeur des notes
                    del liste_notes[3] # enlever la note la plus basse
                    del liste_notes[0] # enlever la note la plus haute
                    note1= liste_notes[0]
                    note2= liste_notes[1]
                    noteE = float(base_exe_indiv) - ((float(note1)+float(note2))/2) - float(notee1e2)

            # calcul note exé groupe
            if type_gymnaste == "G" :
                if  notee4 == 0 and notee5 == 0 and notee6 == 0:
                    noteE = base_exe_group - float(notee3) - float(notee1e2)
                if  notee4 != 0 and notee5 == 0 and notee6 == 0:
                    noteE= base_exe_group - ((float(notee3) + float(notee4)) /2) - float(notee1e2)
                if  notee4 != 0 and notee5 != 0 and notee6 == 0:
                    noteE = float(base_exe_group)- ((float(notee3) + float(notee4) + float(notee5)) / 3) - float(notee1e2)
                if  notee4 != 0 and notee5 != 0 and notee6 != 0 :
                    liste_notes = [notee3, notee4, notee5, notee6] # créer liste des 4 notes exé
                    liste_notes.sort() # trier liste par valeur des notes
                    del liste_notes[3] # enlever la note la plus basse
                    del liste_notes[0] # enlever la note la plus haute
                    note1= liste_notes[0]
                    note2= liste_notes[1]
                    noteE = float(base_exe_group) - ((float(note1)+float(note2))/2) - float(notee1e2)

            # Calcul des notes finales du passage
            noteD = float(noted1d2) + float(noted3d4)
            notepassage = noteE + noteD - float(noteded)
            notepassage = float("%.03f" % (notepassage))
            noteD = float("%.03f" % (noteD))
            noteE = float("%.03f" % (noteE))
            noteded = float(noteded)

            inser_noted1d2 = (noted1d2, nr_gymnaste_select, engin_select)
            inser_noted3d4 = (noted3d4, nr_gymnaste_select, engin_select)
            inser_notee1e2 = (notee1e2, nr_gymnaste_select, engin_select)
            inser_notee3 = (notee3, nr_gymnaste_select, engin_select)
            inser_notee4 = (notee4, nr_gymnaste_select, engin_select)
            inser_notee5 = (notee5, nr_gymnaste_select, engin_select)
            inser_notee6 = (notee6, nr_gymnaste_select, engin_select)
            inser_noteded = (noteded, nr_gymnaste_select, engin_select)
            inser_noteD = (noteD, nr_gymnaste_select, engin_select)
            inser_noteE = (noteE, nr_gymnaste_select, engin_select)
            inser_notepassage = (notepassage, nr_gymnaste_select, engin_select)
            cursor.execute("UPDATE PP_Liste_passages set D1_D2 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_noted1d2)
            cursor.execute("UPDATE PP_Liste_passages set D3_D4 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_noted3d4)
            cursor.execute("UPDATE PP_Liste_passages set E1_E2 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_notee1e2)
            cursor.execute("UPDATE PP_Liste_passages set E3 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_notee3)
            cursor.execute("UPDATE PP_Liste_passages set E4 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_notee4)
            cursor.execute("UPDATE PP_Liste_passages set E5 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_notee5)
            cursor.execute("UPDATE PP_Liste_passages set E6 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_notee6)
            cursor.execute("UPDATE PP_Liste_passages set Déduction = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)",inser_noteded)
            cursor.execute("UPDATE PP_Liste_passages set Note_D = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_noteD)
            cursor.execute("UPDATE PP_Liste_passages set Note_E = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_noteE)
            cursor.execute("UPDATE PP_Liste_passages set Note_passage = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_notepassage)

            connection.commit()
            connection.close()

            """
            # Enregistrer dans le fichier texte le résultat du passage
            data = [nr_gymnaste_select, nom_select, prenom_select, club_select, cat_select, engin_select,str(noteD),str(noteE),str(noteded),str(notepassage)]
            with open ('Résultats.csv', 'a') as fichier :
                ecrire = csv.writer(fichier,delimiter =';')
                ecrire.writerow(data)
            # effacer les lignes vides
            for line_number, line in enumerate(fileinput.input('Résultats.csv', inplace=1)):
                if len(line) > 1:
                    sys.stdout.write(line)
            fichier.close()
            """

            wb = xw.Book('Résultats.xlsx')  # uniquement quand le fichier excel est déjà ouvert
            sht = wb.sheets['Résultats passages']
            nouv_notes_base = [nr_gymnaste_select, nom_select, prenom_select, club_select, cat_select, engin_select,noteD, noteE, noteded, notepassage]
            rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
            rownum = rownum + 1
            rownum = str(rownum)
            print(rownum)
            sht.range('A' + rownum).value = nouv_notes_base

            """
            # Résultas notes passage envoyés dans Google Sheet
            worksheet = spreadsheet.worksheet("Résultats passages")  # ouverture de la feuille de calcul résultats des passages
            nouv_notes_base = [nr_gymnaste_select, nom_select, prenom_select, club_select, cat_select, engin_select, noteD, noteE, noteded, notepassage]
            worksheet.append_row(nouv_notes_base)
            #worksheet.append_row(data)
            """

            command = fermer_fenetre()

        # création de la fenêtre des notes
        fenetrenotes = tkinter.Toplevel()
        fenetrenotes.title("Enregistrement notes")
        fenetrenotes.geometry("1030x700")
        #fenetrenotes.minsize(850, 300)
        fenetrenotes.config(background='cadetblue')
        fenetrenotes.iconbitmap("Logo.ico")

        labelpassage = tkinter.Label(fenetrenotes, text="Passage à enregistrer :", width=18, height=3, bg='cadetblue',font=("Courrier", 12));
        labelpassage.place(x=1, y=10)
        labelpassage = tkinter.Label(fenetrenotes, text="Historique des passages :", width=30, height=3, fg='yellow',bg='cadetblue', font=("Courrier", 12));
        labelpassage.place(x=420, y=120)

        # fenêtre passages effectués gymnaste avec Treeview
        tree_pass = ttk.Treeview(fenetrenotes, columns = (1, 2, 3, 4, 5, 6, 7), height=5, show ="headings")
        tree_pass.place(x=250, y=170, width=750, height=110)
        tree_pass.heading(1, text="Nr. gymnaste");
        tree_pass.column(1, width=80, anchor='center')
        tree_pass.heading(2, text="Nom");
        tree_pass.column(2, width=140)
        tree_pass.heading(3, text="Prénom");
        tree_pass.column(3, width=120)
        tree_pass.heading(4, text="Club");
        tree_pass.column(4, width=120)
        tree_pass.heading(5, text="Catégorie");
        tree_pass.column(5, width=60, anchor='center')
        tree_pass.heading(6, text="Passage engin");
        tree_pass.column(6, width=85)
        tree_pass.heading(7, text="Note engin");
        tree_pass.column(7, width=60)

        # lire dans BD les passages effectués de la gymnaste
        conn = sqlite3.connect("BD GRS.db")
        cur = conn.cursor()
        select_effectue = cur.execute ("select * FROM PP_Gymnastes, PP_Inscriptions, PP_Liste_passages WHERE PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste "
                                       "AND PP_Liste_passages.Nr_gymnaste = PP_Inscriptions.Nr_gymnaste")

        # afficher liste de passages effectués dans Treeview de la gymnaste sélectionnée
        for ligne_effect in select_effectue:
            if ligne_effect[18] != 0 and ligne_effect[7] == nr_gymnaste_select :
            #if ligne_effect[7] == nr_gymnaste_select and ligne_effect[15] != "" or ligne_effect[15] != "0":
                tree_pass.insert('', END, values=(ligne_effect[7], ligne_effect[1], ligne_effect[2], ligne_effect[4], ligne_effect[8], ligne_effect[17], ligne_effect[18]))

        # afficher le passage en cours de la gymnaste sélectionnée (si il y en a encore un !)
        select_effectue = cur.execute("select * FROM PP_Gymnastes, PP_Inscriptions, PP_Liste_passages WHERE PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste "
                                      "AND PP_Liste_passages.Nr_gymnaste = PP_Inscriptions.Nr_gymnaste")

        # champs+étiquettes pour enregistrer notes et afficher info du passage en cours avec contrôle que le passage n'aie pas été déjà fait !
        for ligne_effect in select_effectue:
            if ligne_effect[7] == nr_gymnaste_select and ligne_effect[17] == engin_select and ligne_effect[18] == 0 : # or ligne_effect[12] == None:
            # if ligne_effect[7] == nr_gymnaste_select and ligne_effect[11] == engin_select and ligne_effect[12] == None:
                labelnr_gymn = tkinter.Label(fenetrenotes, text="Nr. gymnaste", width=10, height=1, bg='cadetblue',font=("Courrier", 12));
                labelnr_gymn.place(x=180, y=10)
                info_nr_gymn = tkinter.Message(fenetrenotes, text=nr_gymnaste_select, width=30, fg='yellow',bg='cadetblue', font=("Courrier", 12));
                info_nr_gymn.place(x=210, y=35)
                labelnr_gymn = tkinter.Label(fenetrenotes, text="Nom", width=10, height=1, bg='cadetblue',font=("Courrier", 12));
                labelnr_gymn.place(x=280, y=10)
                info_nr_gymn = tkinter.Message(fenetrenotes, text=nom_select, width=110, fg='yellow', bg='cadetblue',font=("Courrier", 12));
                info_nr_gymn.place(x=303, y=35)
                labelnr_gymn = tkinter.Label(fenetrenotes, text="Prénom", width=10, height=1, bg='cadetblue',font=("Courrier", 12));
                labelnr_gymn.place(x=420, y=10)
                info_nr_gymn = tkinter.Message(fenetrenotes, text=prenom_select, width=110, fg='yellow', bg='cadetblue',font=("Courrier", 12));
                info_nr_gymn.place(x=433, y=35)
                labelnr_gymn = tkinter.Label(fenetrenotes, text="Club", width=10, height=1, bg='cadetblue',font=("Courrier", 12));
                labelnr_gymn.place(x=545, y=10)
                info_nr_gymn = tkinter.Message(fenetrenotes, text=club_select, width=140, fg='yellow', bg='cadetblue',font=("Courrier", 12));
                info_nr_gymn.place(x=570, y=35)
                labelnr_gymn = tkinter.Label(fenetrenotes, text="Catégorie", width=10, height=1, bg='cadetblue',font=("Courrier", 12));
                labelnr_gymn.place(x=720, y=10)
                info_nr_gymn = tkinter.Message(fenetrenotes, text=cat_select, width=30, fg='yellow', bg='cadetblue',font=("Courrier", 12));
                info_nr_gymn.place(x=725, y=35)
                labelnr_gymn = tkinter.Label(fenetrenotes, text="Engin", width=10, height=1, bg='cadetblue',font=("Courrier", 12));
                labelnr_gymn.place(x=830, y=10)
                info_nr_gymn = tkinter.Message(fenetrenotes, text=engin_select, width=110, fg='yellow', bg='cadetblue',font=("Courrier", 12));
                info_nr_gymn.place(x=850, y=35)
                labeld1d2 = tkinter.Label(fenetrenotes, text="Note D1-D2", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labeld1d2.place(x=10, y=110)
                d1d2 = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                d1d2.place(x=20, y=160)
                labeld3d4 = tkinter.Label(fenetrenotes, text="Note D3-D4", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labeld3d4.place(x=10, y=180)
                d3d4 = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                d3d4.place(x=20, y=230)
                labele1e2 = tkinter.Label(fenetrenotes, text="Note E1-E2", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labele1e2.place(x=10, y=250)
                e1e2 = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                e1e2.place(x=20, y=300)
                labele3 = tkinter.Label(fenetrenotes, text="Note E3", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labele3.place(x=10, y=320)
                e3 = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                e3.place(x=20, y=370)
                labele4 = tkinter.Label(fenetrenotes, text="Note E4", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labele4.place(x=10, y=390)
                e4 = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                e4.place(x=20, y=440)
                labele5 = tkinter.Label(fenetrenotes, text="Note E5", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labele5.place(x=10, y=460)
                e5 = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                e5.place(x=20, y=510)
                labele6 = tkinter.Label(fenetrenotes, text="Note E6", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labele6.place(x=10, y=530)
                e6 = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                e6.place(x=20, y=580)
                labelded = tkinter.Label(fenetrenotes, text="Déduction", width=11, height=3, bg='cadetblue',font=("Courrier", 12));
                labelded.place(x=10, y=600)
                ded = tkinter.Entry(fenetrenotes, font=("Courrier", 12), width=10);
                ded.place(x=20, y=650)
                # bouton enregistrer notes
                enr = tkinter.Button(fenetrenotes, text="ENREGISTRER", width=15, height=3, font=("Courrier", 12),bg='cornflowerblue', command=base_donnees);
                enr.place(x=200, y=610)

        # affiche info passage déjà enegistré
        select_effectue = cur.execute("select * FROM PP_Gymnastes, PP_Inscriptions, PP_Liste_passages WHERE PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste "
                                      "AND PP_Liste_passages.Nr_gymnaste = PP_Inscriptions.Nr_gymnaste")

        # informer passage sélectionné déjà enregistré
        for ligne_effect in select_effectue:
            if ligne_effect[7] == nr_gymnaste_select and ligne_effect[17] == engin_select and ligne_effect[18] != 0:
                labelfini = tkinter.Label(fenetrenotes, text="Passage déjà enregistré !", width=120, height=3,fg='yellow', bg='cadetblue', font=("Courrier", 12));
                labelfini.place(x=0, y=10)

        conn.close()

        # créer fonction corrective des notes avec nouvelle fenêtre
        def correctif ():

            def fermer_fenetre():
                fenetrecorr.destroy()

            # création de la fenêtre des corrections
            fenetrecorr = tkinter.Toplevel()
            fenetrecorr.title("Corrections notes")
            fenetrecorr.geometry("950x350")
            fenetrecorr.minsize(500, 300)
            fenetrecorr.config(background='deeppink')
            fenetrecorr.iconbitmap("Logo.ico")

            conn = sqlite3.connect("BD GRS.db")
            cur = conn.cursor()
            # afficher le passage en cours de la gymnaste sélectionnée (pour effectuer correction)
            select_effectue = cur.execute("select * FROM PP_Gymnastes, PP_Inscriptions, PP_Liste_passages WHERE PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste "
                                          "AND PP_Liste_passages.Nr_gymnaste = PP_Inscriptions.Nr_gymnaste")

            labelpassage = tkinter.Label(fenetrecorr, text="Passage à corriger :", width=18, height=3,bg='deeppink', font=("Courrier", 12));
            labelpassage.place(x=1, y=10)

            # champs+étiquettes pour enregistrer notes et afficher info du passage en cours avec contrôle que le passage n'aie pas été déjà fait !
            for ligne_effect in select_effectue:
                if ligne_effect[7] == nr_gymnaste_select and ligne_effect[17] == engin_select:
                    labelnr_gymn = tkinter.Label(fenetrecorr, text="Nr. gymnaste", width=10, height=1, bg='deeppink',font=("Courrier", 12));
                    labelnr_gymn.place(x=180, y=10)
                    info_nr_gymn = tkinter.Message(fenetrecorr, text=nr_gymnaste_select, width=30, fg='yellow',bg='deeppink', font=("Courrier", 12));
                    info_nr_gymn.place(x=210, y=35)
                    labelnr_gymn = tkinter.Label(fenetrecorr, text="Nom", width=10, height=1, bg='deeppink',font=("Courrier", 12));
                    labelnr_gymn.place(x=280, y=10)
                    info_nr_gymn = tkinter.Message(fenetrecorr, text=nom_select, width=110, fg='yellow',bg='deeppink', font=("Courrier", 12));
                    info_nr_gymn.place(x=300, y=35)
                    labelnr_gymn = tkinter.Label(fenetrecorr, text="Prénom", width=10, height=1, bg='deeppink',font=("Courrier", 12));
                    labelnr_gymn.place(x=420, y=10)
                    info_nr_gymn = tkinter.Message(fenetrecorr, text=prenom_select, width=110, fg='yellow',bg='deeppink', font=("Courrier", 12));
                    info_nr_gymn.place(x=430, y=35)
                    labelnr_gymn = tkinter.Label(fenetrecorr, text="Club", width=10, height=1, bg='deeppink',font=("Courrier", 12));
                    labelnr_gymn.place(x=545, y=10)
                    info_nr_gymn = tkinter.Message(fenetrecorr, text=club_select, width=140, fg='yellow',bg='deeppink', font=("Courrier", 12));
                    info_nr_gymn.place(x=570, y=35)
                    labelnr_gymn = tkinter.Label(fenetrecorr, text="Catégorie", width=10, height=1, bg='deeppink',font=("Courrier", 12));
                    labelnr_gymn.place(x=713, y=10)
                    info_nr_gymn = tkinter.Message(fenetrecorr, text=cat_select, width=30, fg='yellow', bg='deeppink',font=("Courrier", 12));
                    info_nr_gymn.place(x=720, y=35)
                    labelnr_gymn = tkinter.Label(fenetrecorr, text="Engin", width=10, height=1, bg='deeppink',font=("Courrier", 12));
                    labelnr_gymn.place(x=810, y=10)
                    info_nr_gymn = tkinter.Message(fenetrecorr, text=engin_select, width=110, fg='yellow',bg='deeppink', font=("Courrier", 12));
                    info_nr_gymn.place(x=828, y=35)
                    # entrées et label pour notes corrigées
                    labeld1d2 = tkinter.Label(fenetrecorr, text="Note D1-D2", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labeld1d2.place(x=15, y=80)
                    d1d2 = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    d1d2.place(x=20, y=160)
                    labeld3d4 = tkinter.Label(fenetrecorr, text="Note D3-D4", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labeld3d4.place(x=125, y=80)
                    d3d4 = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    d3d4.place(x=130, y=160)
                    labele1e2 = tkinter.Label(fenetrecorr, text="Note E1-E2", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labele1e2.place(x=235, y=80)
                    e1e2 = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    e1e2.place(x=240, y=160)
                    labele3 = tkinter.Label(fenetrecorr, text="Note E3", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labele3.place(x=345, y=80)
                    e3 = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    e3.place(x=350, y=160)
                    labele4 = tkinter.Label(fenetrecorr, text="Note E4", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labele4.place(x=455, y=80)
                    e4 = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    e4.place(x=460, y=160)
                    labele5 = tkinter.Label(fenetrecorr, text="Note E5", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labele5.place(x=565, y=80)
                    e5 = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    e5.place(x=570, y=160)
                    labele6 = tkinter.Label(fenetrecorr, text="Note E6", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labele6.place(x=675, y=80)
                    e6 = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    e6.place(x=680, y=160)
                    labelded = tkinter.Label(fenetrecorr, text="Déduction", width=11, height=3, bg='deeppink',font=("Courrier", 12));
                    labelded.place(x=785, y=80)
                    ded = tkinter.Entry(fenetrecorr, font=("Courrier", 12), width=10);
                    ded.place(x=790, y=160)

            #créer liste des notes à corriger depuis la BD
            list_notes = []
            select = cur.execute ("select * FROM PP_Liste_passages ")
            for x in cur.fetchall():
                if x[1] == nr_gymnaste_select and x[2] == engin_select:
                    list_notes.append(x[4]);list_notes.append(x[5]);list_notes.append(x[6]);list_notes.append(x[7]);list_notes.append(x[8]);list_notes.append(x[9]);list_notes.append(x[10]);list_notes.append(x[11])

            conn.close()

            # afficher les notes du passage enregistré
            noted1d2 = tkinter.Message(fenetrecorr, text=list_notes[0], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            noted1d2.place(x=25, y=125)
            noted3d4 = tkinter.Message(fenetrecorr, text=list_notes[1], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            noted3d4.place(x=135, y=125)
            notee1e2 = tkinter.Message(fenetrecorr, text=list_notes[2], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            notee1e2.place(x=245, y=125)
            notee3 = tkinter.Message(fenetrecorr, text=list_notes[3], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            notee3.place(x=355, y=125)
            notee4 = tkinter.Message(fenetrecorr, text=list_notes[4], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            notee4.place(x=465, y=125)
            notee5 = tkinter.Message(fenetrecorr, text=list_notes[5], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            notee5.place(x=575, y=125)
            notee6 = tkinter.Message(fenetrecorr, text=list_notes[6], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            notee6.place(x=685, y=125)
            noteded = tkinter.Message(fenetrecorr, text=list_notes[7], width=40, fg='yellow', bg='deeppink', font=("Courrier", 12));
            noteded.place(x=795, y=125)

            # enregistrer les corrections
            def base_donnees_bis () :

                noted1d2 = (d1d2.get())
                noted3d4 = (d3d4.get())
                notee1e2 = (e1e2.get())
                notee3 = (e3.get())
                notee4 = (e4.get())
                notee5 = (e5.get())
                notee6 = (e6.get())
                noteded = (ded.get())

                if noted1d2 == "":
                    noted1d2 = list_notes[0];noted1d2 = float(noted1d2)
                if noted3d4 == "":
                    noted3d4 = list_notes[1];noted3d4 = float(noted3d4)
                if notee1e2 == "":
                    notee1e2 = list_notes[2];notee1e2 = float(notee1e2)
                if notee3 == "":
                    notee3 = list_notes[3];notee3 = float(notee3)
                if notee4 == "":
                    notee4 = list_notes[4];notee4 = float(notee4)
                if notee5 == "":
                    notee5 = list_notes[5];notee5 = float(notee5)
                if notee6 == "":
                    notee6 = list_notes[6];notee6 = float(notee6)
                if noteded == "":
                    noteded = list_notes[7];noteded = float(noteded)

                # récupérer valeur départ note exe en indiv
                with open("depart_note exe_indiv.txt", 'r') as fichier:
                    base_exe_indiv = fichier.readline()
                    base_exe_indiv = int(base_exe_indiv)
                    # récupérer valeur départ note exe en group
                with open("depart_note exe_group.txt", 'r') as fichier:
                    base_exe_group = fichier.readline()
                    base_exe_group = int(base_exe_group)

                connection = sqlite3.connect("BD GRS.db")
                cursor = connection.cursor()

                type_pass = (nr_gymnaste_select,)
                cursor.execute("select * FROM PP_Inscriptions WHERE Nr_gymnaste = ?", type_pass)
                type_gymnaste = cursor.fetchone()[4]

                # calcul note exé individuelle
                if type_gymnaste == "I":
                    if notee4 == 0 and notee5 == 0 and notee6 == 0:
                        noteE = base_exe_indiv - float(notee3) - float(notee1e2)
                    if notee4 != 0 and notee5 == 0 and notee6 == 0:
                        noteE = base_exe_indiv - ((float(notee3) + float(notee4)) / 2) - float(notee1e2)
                    if notee4 != 0 and notee5 != 0 and notee6 == 0:
                        noteE = float(base_exe_indiv) - ((float(notee3) + float(notee4) + float(notee5)) / 3) - float(notee1e2)
                    if notee4 != 0 and notee5 != 0 and notee6 != 0:
                        liste_notes = [float(notee3), float(notee4), float(notee5), float(notee6)]  # créer liste des 4 notes exé
                        liste_notes.sort()  # trier liste par valeur des notes
                        del liste_notes[3]  # enlever la note la plus basse
                        del liste_notes[0]  # enlever la note la plus haute
                        note1 = liste_notes[0]
                        note2 = liste_notes[1]
                        noteE = float(base_exe_indiv) - ((float(note1) + float(note2)) / 2) - float(notee1e2)

                # calcul note exé groupe
                if type_gymnaste == "G":
                    if notee4 == 0 and notee5 == 0 and notee6 == 0:
                        noteE = base_exe_group - float(notee3) - float(notee1e2)
                    if notee4 != 0 and notee5 == 0 and notee6 == 0:
                        noteE = base_exe_group - ((float(notee3) + float(notee4)) / 2) - float(notee1e2)
                    if notee4 != 0 and notee5 != 0 and notee6 == 0:
                        noteE = float(base_exe_group) - ((float(notee3) + float(notee4) + float(notee5)) / 3) - float(notee1e2)
                    if notee4 != 0 and notee5 != 0 and notee6 != 0:
                        liste_notes = [notee3, notee4, notee5, notee6]  # créer liste des 4 notes exé
                        liste_notes.sort()  # trier liste par valeur des notes
                        del liste_notes[3]  # enlever la note la plus basse
                        del liste_notes[0]  # enlever la note la plus haute
                        note1 = liste_notes[0]
                        note2 = liste_notes[1]
                        noteE = float(base_exe_group) - ((float(note1) + float(note2)) / 2) - float(notee1e2)

                # Calcul des notes finales du passage
                noteD = float(noted1d2) + float(noted3d4)
                notepassage = noteE + noteD - float(noteded)
                notepassage = float("%.03f" % (notepassage))
                noteD = float("%.03f" % (noteD))
                noteE = float("%.03f" % (noteE))
                noteded=float(noteded)

                inser_noted1d2 = (noted1d2, nr_gymnaste_select, engin_select)
                inser_noted3d4 = (noted3d4, nr_gymnaste_select, engin_select)
                inser_notee1e2 = (notee1e2, nr_gymnaste_select, engin_select)
                inser_notee3 = (notee3, nr_gymnaste_select, engin_select)
                inser_notee4 = (notee4, nr_gymnaste_select, engin_select)
                inser_notee5 = (notee5, nr_gymnaste_select, engin_select)
                inser_notee6 = (notee6, nr_gymnaste_select, engin_select)
                inser_noteded = (noteded, nr_gymnaste_select, engin_select)
                inser_noteD = (noteD, nr_gymnaste_select, engin_select)
                inser_noteE = (noteE, nr_gymnaste_select, engin_select)
                inser_notepassage = (notepassage, nr_gymnaste_select, engin_select)
                cursor.execute("UPDATE PP_Liste_passages set D1_D2 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_noted1d2)
                cursor.execute("UPDATE PP_Liste_passages set D3_D4 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_noted3d4)
                cursor.execute("UPDATE PP_Liste_passages set E1_E2 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_notee1e2)
                cursor.execute("UPDATE PP_Liste_passages set E3 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_notee3)
                cursor.execute("UPDATE PP_Liste_passages set E4 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_notee4)
                cursor.execute("UPDATE PP_Liste_passages set E5 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_notee5)
                cursor.execute("UPDATE PP_Liste_passages set E6 = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_notee6)
                cursor.execute("UPDATE PP_Liste_passages set Déduction = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_noteded)
                cursor.execute("UPDATE PP_Liste_passages set Note_D = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_noteD)
                cursor.execute("UPDATE PP_Liste_passages set Note_E = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_noteE)
                cursor.execute("UPDATE PP_Liste_passages set Note_passage = (?) WHERE Nr_gymnaste = (?) and Engin_passage = (?)", inser_notepassage)

                wb = xw.Book('Résultats.xlsx')  # uniquement quand le fichier excel est déjà ouvert
                sht = wb.sheets['Résultats passages']
                nouv_notes = [nr_gymnaste_select, nom_select, prenom_select, club_select, cat_select, engin_select,noteD, noteE, noteded, notepassage, "CORRECTIF"]
                rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
                rownum = rownum + 1
                rownum = str(rownum)
                print(rownum)
                sht.range('A' + rownum).value = nouv_notes

                """
                worksheet = spreadsheet.worksheet("Résultats passages")  # ouverture de la feuille de calcul résultats des passages
                nouv_notes = [nr_gymnaste_select, nom_select,prenom_select,club_select,cat_select,engin_select,noteD,noteE,noteded,notepassage, "CORRECTIF"]
                worksheet.append_row(nouv_notes)
                """

                connection.commit()
                connection.close()

                fenetrecorr.destroy()

            # bouton enregistrer notes
            enr = tkinter.Button(fenetrecorr, text="ENREGISTRER", width=15, height=3, font=("Courrier", 12),bg='cornflowerblue', command=base_donnees_bis)
            enr.place(x=200, y=270)

            # bouton retour
            retour = tkinter.Button(fenetrecorr, text="RETOUR", width=12, height=2, font=("Courrier", 10),bg='cornflowerblue', command=fermer_fenetre)
            retour.place(x=800, y=300)

        # bouton retour
        retour = tkinter.Button(fenetrenotes, text="RETOUR", width=12, height=2, font=("Courrier", 10),bg='cornflowerblue', command = fermer_fenetre)
        retour.place(x=800, y=635)
        # bouton correction
        corr = tkinter.Button(fenetrenotes, text="CORRECTIF", width=15, height=3, font=("Courrier", 12),bg='cornflowerblue',command = correctif)
        corr.place(x=500, y=610)

     # création fenêtre liste de passage complète avec Treeview
    scrollbar = Scrollbar(fenetrepassage)
    scrollbar.pack(side=RIGHT, fill=Y)
    tree = ttk.Treeview(fenetrepassage, columns=(1, 2, 3, 4, 5, 6), height=5, show="headings",yscrollcommand=scrollbar.set)
    scrollbar.config(command=tree.yview);tree.place(x=50, y=60, width=650, height=175)
    tree.heading(1, text="Nr. gymnaste");tree.column(1, width=70, anchor='center')
    tree.heading(2, text="Nom");tree.column(2, width=150)
    tree.heading(3, text="Prénom");tree.column(3, width=100)
    tree.heading(4, text="Club");tree.column(4, width=130)
    tree.heading(5, text="Catégorie");tree.column(5, width=60, anchor='center')
    tree.heading(6, text="Passage engin");tree.column(6, width=80)
    labelpass = tkinter.Label(fenetrepassage, text="Ordre de passage", width=20, height=3, bg='yellow',font=("Courrier", 12));labelpass.place(x=200, y=0)

    # créer liste de passage complète depuis base de donnée
    conn = sqlite3.connect("BD GRS.db")
    cur = conn.cursor()

    selection = cur.execute("select * FROM PP_Gymnastes, PP_Inscriptions, PP_Liste_passages WHERE PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste AND PP_Liste_passages.Nr_gymnaste = PP_Inscriptions.Nr_gymnaste ORDER BY ID_liste_passage")

    for ligne in selection:  # afficher liste de passage complète dans Treeview
        tree.insert('', END, values=(ligne[7], ligne[1], ligne[2], ligne[4], ligne[8], ligne[17]))
    conn.close()

    # crée la fonction (ouvrir fenêtre notes) lorsque je sélectionne une ligne de liste de passage
    tree.bind("<<TreeviewSelect>>", select_pass_gymn)

    def out():
        messagebox.showerror("OUT")

    # bouton retour
    retour = tkinter.Button(fenetrepassage, text="RETOUR", width=12, height=2, font=("Courrier", 10),bg='cornflowerblue', command=fermer_fenetre);retour.place(x=400, y=260)
    # bouton resultats
    resultats = tkinter.Button(fenetrepassage, text="RESULTATS", width=12, height=2, font=("Courrier", 10),bg='cornflowerblue', command=resultats_complets);resultats.place(x=100, y=260)
    # bouton recherche
    recherche = tkinter.Button(fenetrepassage, text="RECHERCHE", width=12, height=2, font=("Courrier", 10),bg='cornflowerblue', command=out);recherche.place(x=250, y=260)

def resultats_complets():
    connection = sqlite3.connect("BD GRS.db")
    cursor = connection.cursor()

    # créer liste des nr. gymnaste
    cursor.execute("select Nr_gymnaste FROM PP_Inscriptions " )
    liste = cursor.fetchall()
    n=len(liste)

    liste_a=[] #créer liste pour les catégories inscrites
    x=1
    while x <= n :
        cursor.execute ("select * FROM PP_Inscriptions WHERE ID_inscription = ? ",(x,))
        item = cursor.fetchone()[2]
        liste_a.append(item)
        x+=1

    # calculer et noter dans BD les notes finales
    for item in liste_a:
        liste_b = [] # créer liste des notes de passage de la gymnaste sélectionnée
        cursor.execute("select * FROM PP_Liste_passages WHERE Nr_gymnaste = ? ", (item,))
        for x in cursor.fetchall():
            liste_b.append(x[3])
            note_tot = sum(liste_b) # calculer la note totale
            note_tot = float("%.03f" % (note_tot))
            note_finale =(note_tot, item)
            cursor.execute ("UPDATE PP_Inscriptions set Note_totale = (?) WHERE Nr_gymnaste= (?) ", note_finale) # noter la note finale ds BD

    def aff_résultats_finaux_champ():
        connection = sqlite3.connect("BD GRS.db")
        cursor = connection.cursor()
        # sélection et aplatir liste des passages de la gymnaste sélectionnée
        with open("donnees_cat.txt", 'r') as fichier:
            caté = fichier.readlines()
        catégorie = []
        for item in caté:
            catégorie.append(item.rstrip('\n'))  # liste des catégories enregistrées pour le concours
        liste_gymn = []
        selection = cursor.execute("Select * From PP_Inscriptions")
        for ligne in selection:
            liste_gymn.append(ligne[2])  # liste des gymanstes inscrite pour le concours

        """
        worksheet = spreadsheet.worksheet("Résultats championnat")  # ouverture de la feuille de calcul résultats championnat
        worksheet.format("A2:A200", {"textFormat": {"fontSize": 11, "bold": True}})
        worksheet.format("G2:G200", {"textFormat": {"fontSize": 11, "bold": True}})
        # worksheet.format ("A3", {"backgroundColor": {"red": 0.0,"green": 0.0,"blue": 0.0 },"horizontalAlignment": "CENTER","textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 1.0},"fontSize": 12,"bold": True}})
        # boucle avec les  catégories, dans l'ordre des rangs des gymnastes
        """

        wb = xw.Book('Résultats.xlsx')
        sht = wb.sheets['Résultats finaux']

        for item in catégorie:
            rang = [item]
            rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
            rownum = rownum + 1;
            rownum = str(rownum)
            sht.range('A' + rownum).value = rang
            #worksheet.append_row(rang)
            gymn = []
            selection = cursor.execute("select * FROM PP_Gymnastes, PP_Inscriptions  WHERE "
                                       "PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste "
                                       "AND PP_Inscriptions.Catégorie = ? "
                                       "AND PP_Inscriptions.Invité = 'non' "
                                       "AND PP_Inscriptions.ID_gymnaste ORDER BY Note_totale DESC ", (item,))
            for ligne in selection:
                gymn.append((ligne[12], ligne[7], ligne[1], ligne[2], ligne[4], ligne[3], ligne[11]))
            for x in gymn:
                #worksheet.append_row(x)
                rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
                rownum = rownum + 1;
                rownum = str(rownum)
                sht.range('A' + rownum).value = x

            espace = [' ']
            #worksheet.append_row(espace)
            rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
            rownum = rownum + 1;
            rownum = str(rownum)
            sht.range('A' + rownum).value = espace

        connection.close()

    # créer nouvelle fenêtre pour afficher les résultats par catégorie
    fenetrechampionnat = tkinter.Toplevel()
    fenetrechampionnat.title("Résultats concours")
    fenetrechampionnat.geometry("800x800")
    fenetrechampionnat.config(background='blue')
    fenetrechampionnat.iconbitmap("Logo.ico")

    def fermer_fenetre_champ():
        fenetrechampionnat.destroy()

    # création fenêtre résultat avec Treeview
    scrollbar = Scrollbar(fenetrechampionnat)
    scrollbar.pack(side=RIGHT, fill=Y)
    tree = ttk.Treeview(fenetrechampionnat, columns=(1, 2, 3, 4, 5, 6, 7), height=5, show="headings",yscrollcommand=scrollbar.set)
    scrollbar.config(command=tree.yview);tree.place(x=50, y=93, width=700, height=550)
    tree.heading(1, text="Rang");tree.column(1, width=60, anchor='center')
    tree.heading(2, text="Nr.");tree.column(2, width=40, anchor='center')
    tree.heading(3, text="Nom");tree.column(3, width=140)
    tree.heading(4, text="Prénom");tree.column(4, width=120)
    tree.heading(5, text="Club");tree.column(5, width=120)
    tree.heading(6, text="Catégorie");tree.column(6, width=60, anchor='center')
    tree.heading(7, text="Note finale");tree.column(7, width=85)

    #trier les résultats par catégorie et afficher dans Treeview - championnat vaudois
    cat = [] # créer liste des catégories
    gymnastes = [] # créer liste des nr. de gymnastes
    # Championnat vaudois résultats
    selection = cursor.execute("select * FROM PP_Gymnastes, PP_Inscriptions WHERE PP_Inscriptions.Invité= 'non' "
                               "AND PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste ORDER BY Catégorie, Note_totale DESC ")

    for ligne in selection:  # afficher résultats dans Treeview
        cat.append(ligne[8])
        gymnastes.append(ligne[7])
        tree.insert('', END, values=(ligne[12], ligne[7], ligne[1], ligne[2], ligne[4], ligne[8], ligne[11]))

    # créer listes des rangs
    l = len(cat)  # compte le nbre de gymnastes
    rang_championnat = [1]
    c = 1
    r = 1
    for x in cat: # boucle pour créer liste des rangs dans l'odre des notes et des catégories
        if c==l:
            break
        if cat[c] == cat[c-1]:
            rang_championnat.append(r+1) # même catégorie
            r+= 1
        else:
            r=1
            rang_championnat.append(r) # autre catégorie
        c+=1

    # noter dans BD les rangs
    ra=0
    for item in gymnastes:
        cursor.execute ("UPDATE PP_Inscriptions set Rang_championnat = ? WHERE Nr_gymnaste= ? ", [(rang_championnat[ra]), item]) # noter le rang ds BD
        ra+=1

    def aff_résultats_finaux_coupe():
        # refaire liste résultat par catégorie pour affichage dans Google sheet
        connection = sqlite3.connect("BD GRS.db")
        cursor = connection.cursor()

        """
        worksheet = spreadsheet.worksheet("Résultats coupe")  # ouverture de la feuille de calcul résultats coupe
        worksheet.format("A2:A100", {"textFormat": {"fontSize": 11, "bold": True}})
        worksheet.format("E2:E100", {"textFormat": {"fontSize": 11, "bold": True}})
        # worksheet.format ("A3", {"backgroundColor": {"red": 0.0,"green": 0.0,"blue": 0.0 },"horizontalAlignment": "CENTER","textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 1.0},"fontSize": 12,"bold": True}})
        # boucle avec les  catégories, dans l'ordre des rangs des gymnastes
        """

        catégorie = list(set(cat))  # enlever les doublons
        catégorie = sorted(catégorie)  # mettre dans l'ordre

        wb = xw.Book('Résultats.xlsx')
        sht = wb.sheets['Résultats coupe']

        for item in catégorie:
            rang = [item]
            rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
            rownum = rownum + 1;
            rownum = str(rownum)
            sht.range('A' + rownum).value = rang
            #worksheet.append_row(rang)
            gymn = []
            selection = cursor.execute("select * FROM PP_Gymnastes, PP_Inscriptions  WHERE "
                                       "PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste "
                                       "AND PP_Inscriptions.Catégorie = ? "
                                       "AND PP_Inscriptions.Type_cat = 'G' "
                                       "AND PP_Inscriptions.ID_gymnaste ORDER BY Note_totale DESC ", (item,))
            for ligne in selection:
                gymn.append((ligne[13], ligne [7], ligne[1], ligne[4], ligne[11]))
            for x in gymn:
                rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
                rownum = rownum + 1 ; rownum = str(rownum)
                sht.range('A' + rownum).value = x
                #worksheet.append_row(x)

            espace = [' ']
            rownum = sht.range('A1').current_region.last_cell.row  # me donne la dernière ligne utilisée
            rownum = rownum + 1 ; rownum = str(rownum)
            sht.range('A' + rownum).value = espace
            #worksheet.append_row(espace)

        connection.close()

    fenetrecoupe = tkinter.Toplevel()
    fenetrecoupe.title("Résultats concours")
    fenetrecoupe.geometry("800x800")
    fenetrecoupe.config(background='blue')
    fenetrecoupe.iconbitmap("Logo.ico")

    def fermer_fenetre_coupe():
        fenetrecoupe.destroy()

    # titre fenêtre
    label_titre = tkinter.Label(fenetrechampionnat, text="Résultats Open Broyard", width=50, height=3, fg='yellow', bg='blue',font=("Courrier", 16));label_titre.place(x=120, y=2)
    # bouton retour
    retour = tkinter.Button(fenetrechampionnat, text="RETOUR", width=12, height=2, font=("Courrier", 10), bg='cornflowerblue',command=fermer_fenetre_champ);retour.place(x=400, y=720)
    # bouton envoi sur internet
    envoi = tkinter.Button(fenetrechampionnat, text="ENVOI SUR INTERNET", width=20, height=2, font=("Courrier", 10), bg='cornflowerblue', command=aff_résultats_finaux_champ);envoi.place(x=210, y=720)

    # titre fenêtre
    label_titre = tkinter.Label(fenetrecoupe, text="Résultats coupe Vaudoise", width=50, height=3, fg='yellow', bg='blue',font=("Courrier", 16));label_titre.place(x=120, y=2)
    # bouton retour
    retour = tkinter.Button(fenetrecoupe, text="RETOUR", width=12, height=2, font=("Courrier", 10), bg='cornflowerblue',command=fermer_fenetre_coupe);retour.place(x=400, y=720)
    # bouton envoi résultats
    envoi = tkinter.Button(fenetrecoupe, text="ENVOI SUR INTERNET", width=20, height=2, font=("Courrier", 10), bg='cornflowerblue',command=aff_résultats_finaux_coupe);envoi.place(x=210, y=720)

    # création fenêtre résultat avec Treeview
    scrollbar = Scrollbar(fenetrecoupe)
    scrollbar.pack(side=RIGHT, fill=Y)
    tree = ttk.Treeview(fenetrecoupe, columns=(1, 2, 3, 4, 5, 6,7), height=5, show="headings",yscrollcommand=scrollbar.set)
    scrollbar.config(command=tree.yview);tree.place(x=50, y=93, width=700, height=550)
    tree.heading(1, text="Rang");tree.column(1, width=60, anchor='center')
    tree.heading(2, text="Nr.");tree.column(2, width=40, anchor='center')
    tree.heading(3, text="Nom");tree.column(3, width=140)
    tree.heading(4, text="Prénom");tree.column(4, width=120)
    tree.heading(5, text="Club");tree.column(5, width=120)
    tree.heading(6, text="Catégorie");tree.column(6, width=60, anchor='center')
    tree.heading(7, text="Note finale");tree.column(7, width=85)

    #trier les résultats par catégorie et afficher dans Treeview - coupe vaudoise
    cat = [] # créer liste des catégories
    gymnastes = [] # créer liste des nr. de gymnastes
    # Coupe vaudoise résultats
    selection = cursor.execute("select * FROM PP_Gymnastes, PP_Inscriptions WHERE PP_Inscriptions.Type_cat = 'G' "
                               "AND PP_Gymnastes.ID_gymnaste = PP_Inscriptions.ID_gymnaste ORDER BY Catégorie, Note_totale DESC ")

    for ligne in selection:  # afficher résultats dans Treeview
        cat.append(ligne[8])
        gymnastes.append(ligne[7])
        tree.insert('', END, values=(ligne[13], ligne[7], ligne[1], ligne[2], ligne[4], ligne[8], ligne[11]))

    # créer listes des rangs
    l = len(cat)  # compte le nbre de gymnastes
    rang_coupe = [1] # 1er rang de la catégorie sélectionnée
    c = 1
    r = 1
    for x in cat: # boucle pour créer liste des rangs dans l'odre des notes et des catégories
        if c==l:
            break
        if cat[c] == cat[c-1]:
            rang_coupe.append(r+1) # même catégorie
            r+= 1
        else:
            r=1
            rang_coupe.append(r) # autre catégorie
        c+=1

    # noter dans BD les rangs
    ra=0
    for item in gymnastes:
        cursor.execute ("UPDATE PP_Inscriptions set Rang_coupe = ? WHERE Nr_gymnaste= ? ", [(rang_coupe[ra]), item]) # noter le rang ds BD
        ra+=1

    connection.commit()
    connection.close()

def preparation():
    # création de la fenêtre préparation
    fenetreprep = tkinter.Toplevel()
    fenetreprep.title("Préparation")
    fenetreprep.geometry("900x420")
    fenetreprep.config(background='gray')
    fenetreprep.iconbitmap("Logo.ico")

    # création fenêtre juges et catégories
    def cat_juges():
        fenetrecatjuges = tkinter.Toplevel()
        fenetrecatjuges.title("Enregistrement juges et catégories")
        fenetrecatjuges.geometry("850x700")
        fenetrecatjuges.config(background='olive')

        nom_cat = ["P1", "P2", "P3", "P4", "P5", "P6", "R2", "R3", "R4", "R5", "R6", "RJ", "G1", "G2", "G3","G4"]  # liste des noms de toutes les catégories

        # choix notes départ indiv et groupe
        note_exe_indiv = tkinter.Label(fenetrecatjuges, text="note exé indiv. : ", width=28, height=2, fg='blue',bg='olive', font=("Courrier", 13));
        note_exe_indiv.place(x=15, y=400)
        note_exe_group = tkinter.Label(fenetrecatjuges, text="note exé groupe : ", width=28, height=2, fg='blue',bg='olive', font=("Courrier", 13));
        note_exe_group.place(x=21, y=430)

        with open("depart_note exe_indiv.txt", 'r') as fichier:
            v_indiv_read = fichier.readline()

        with open("depart_note exe_group.txt", 'r') as fichier:
            v_group_read = fichier.readline()

        label_indiv = tkinter.Label(fenetrecatjuges, text=v_indiv_read, width=2, height=2, fg='yellow', bg='olive',font=("Courrier", 13));
        label_indiv.place(x=250, y=400)
        label_group = tkinter.Label(fenetrecatjuges, text=v_group_read, width=2, height=2, fg='yellow', bg='olive',font=("Courrier", 13));
        label_group.place(x=250, y=430)

        # observeur note exé
        def update_exe(*args):

            with open("depart_note exe_indiv.txt", 'r') as fichier:
                v_indiv_read = fichier.readline()

            with open("depart_note exe_group.txt", 'r') as fichier:
                v_group_read = fichier.readline()

            v_indiv.set(v_indiv_entry.get())
            if v_indiv.get() != "":
                indiv = v_indiv.get()
            else:
                indiv = v_indiv_read

            v_group.set(v_group_entry.get())
            if v_group.get() != "":
                group = v_group.get()
            else:
                group = v_group_read

            with open("depart_note exe_indiv.txt", 'w') as fichier:
                fichier.write(indiv)

            with open("depart_note exe_group.txt", 'w') as fichier:
                fichier.write(group)

        v_indiv_entry = StringVar()
        v_group_entry = StringVar()
        v_indiv_entry.trace('w', update_exe)
        v_group_entry.trace('w', update_exe)
        entry_indiv = tkinter.Entry(fenetrecatjuges, textvariable=v_indiv_entry, justify='center',font=("Courrier", 13), width=4);
        entry_indiv.place(x=280, y=408)
        entry_group = tkinter.Entry(fenetrecatjuges, textvariable=v_group_entry, justify='center',font=("Courrier", 13), width=4);
        entry_group.place(x=280, y=438)

        v_indiv = StringVar()
        v_group = StringVar()

        def enreg_note_depart():

            with open("depart_note exe_indiv.txt", 'r') as fichier:
                v_indiv_read = fichier.readline()

            with open("depart_note exe_group.txt", 'r') as fichier:
                v_group_read = fichier.readline()

            label_indiv_trace = tkinter.Label(fenetrecatjuges, textvariable=v_indiv, width=2, height=2, fg='blue',bg='olive', font=("Courrier", 13));
            label_group_trace = tkinter.Label(fenetrecatjuges, textvariable=v_group, width=2, height=2, fg='blue',bg='olive', font=("Courrier", 13));

            fichier.close()

        # action enclenchée dès changement état de l'une des catégories
        def update_observer(*args):
            varp1 = vp1.get();varp2 = vp2.get();varp3 = vp3.get();varp4 = vp4.get();varp5 = vp5.get();varp6 = vp6.get();varp7 = vp7.get();varp8 = vp8.get();
            varp9 = vp9.get();varp10 = vp10.get();varp11 = vp11.get();varp12 = vp12.get();varp13 = vp13.get();varp14 = vp14.get();varp15 = vp15.get();varp16 = vp16.get()

            new_nom = []  # liste à créer des catégories sélectionnées
            categories_base = [varp1, varp2, varp3, varp4, varp5, varp6, varp7, varp8, varp9, varp10, varp11, varp12,
                               varp13, varp14, varp15, varp16]
            i = 0
            for x in categories_base:  # enregistrer dans liste les catégories sélectionnées
                if x == 1:
                    new_nom.append(nom_cat[i])
                i += 1

            with open("donnees_cat.txt", 'w') as fichier:  # enregistrer dans fichier les catégories sélectionnées
                for item in new_nom:
                    fichier.write(item + '\n')

            fichier.close()

            # déclaration variable de contrôle catégories
        vp1 = tkinter.IntVar();vp2 = tkinter.IntVar();vp3 = tkinter.IntVar();vp4 = tkinter.IntVar();vp5 = tkinter.IntVar();vp6 = tkinter.IntVar();vp7 = tkinter.IntVar();vp8 = tkinter.IntVar();
        vp9 = tkinter.IntVar();vp10 = tkinter.IntVar();vp11 = tkinter.IntVar();vp12 = tkinter.IntVar();vp13 = tkinter.IntVar();vp14 = tkinter.IntVar();vp15 = tkinter.IntVar();vp16 = tkinter.IntVar()

        # afficher les catégories
        cat_label = tkinter.Label(fenetrecatjuges, text="Sélectionner les catégories : ", width=25, height=2, fg='blue',bg='olive', font=("Courrier", 13));
        cat_label.place(x=30, y=10)
        p1 = tkinter.Checkbutton(fenetrecatjuges, text="P1", font=("Courrier", 13), variable=vp1, fg='blue', bg='olive',width=5);
        p1.place(x=5, y=50)
        p2 = tkinter.Checkbutton(fenetrecatjuges, text="P2", font=("Courrier", 13), variable=vp2, fg='blue', bg='olive',width=5);
        p2.place(x=5, y=80)
        p3 = tkinter.Checkbutton(fenetrecatjuges, text="P3", font=("Courrier", 13), variable=vp3, fg='blue', bg='olive',width=5);
        p3.place(x=5, y=110)
        p4 = tkinter.Checkbutton(fenetrecatjuges, text="P4", font=("Courrier", 13), variable=vp4, fg='blue', bg='olive',width=5);
        p4.place(x=5, y=140)
        p5 = tkinter.Checkbutton(fenetrecatjuges, text="P5", font=("Courrier", 13), variable=vp5, fg='blue', bg='olive',width=5);
        p5.place(x=5, y=170)
        p6 = tkinter.Checkbutton(fenetrecatjuges, text="P6", font=("Courrier", 13), variable=vp6, fg='blue', bg='olive',width=5);
        p6.place(x=5, y=200)
        p7 = tkinter.Checkbutton(fenetrecatjuges, text="R2", font=("Courrier", 13), variable=vp7, fg='blue', bg='olive',width=5);
        p7.place(x=100, y=50)
        p8 = tkinter.Checkbutton(fenetrecatjuges, text="R3", font=("Courrier", 13), variable=vp8, fg='blue', bg='olive',width=5);
        p8.place(x=100, y=80)
        p9 = tkinter.Checkbutton(fenetrecatjuges, text="R4", font=("Courrier", 13), variable=vp9, fg='blue', bg='olive',width=5);
        p9.place(x=100, y=110)
        p10 = tkinter.Checkbutton(fenetrecatjuges, text="R5", font=("Courrier", 13), variable=vp10, fg='blue',bg='olive', width=5);
        p10.place(x=100, y=140)
        p11 = tkinter.Checkbutton(fenetrecatjuges, text="R6", font=("Courrier", 13), variable=vp11, fg='blue',bg='olive', width=5);
        p11.place(x=100, y=170)
        p12 = tkinter.Checkbutton(fenetrecatjuges, text="RJ", font=("Courrier", 13), variable=vp12, fg='blue',bg='olive', width=5);
        p12.place(x=191, y=50)
        p13 = tkinter.Checkbutton(fenetrecatjuges, text="G1", font=("Courrier", 13), variable=vp13, fg='blue',bg='olive', width=5);
        p13.place(x=191, y=80)
        p14 = tkinter.Checkbutton(fenetrecatjuges, text="G2", font=("Courrier", 13), variable=vp14, fg='blue',bg='olive', width=5);
        p14.place(x=191, y=110)
        p15 = tkinter.Checkbutton(fenetrecatjuges, text="G3", font=("Courrier", 13), variable=vp15, fg='blue',bg='olive', width=5);
        p15.place(x=191, y=140)
        p16 = tkinter.Checkbutton(fenetrecatjuges, text="G4", font=("Courrier", 13), variable=vp16, fg='blue',bg='olive', width=5);
        p16.place(x=191, y=170)

        nom_cat_base = [p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15,p16]  # liste des noms des checkbutton

        # traceur
        vp1.trace("w", update_observer);vp2.trace("w", update_observer);vp3.trace("w", update_observer);vp4.trace("w", update_observer);vp5.trace("w", update_observer);vp6.trace("w", update_observer);vp7.trace("w", update_observer);vp8.trace("w", update_observer);
        vp9.trace("w", update_observer);vp10.trace("w", update_observer);vp11.trace("w", update_observer);vp12.trace("w", update_observer);vp13.trace("w", update_observer);vp14.trace("w", update_observer);vp15.trace("w", update_observer);vp16.trace("w", update_observer)

        # activer les catégories choisies
        with open("donnees_cat.txt", 'r') as fichier:
            new_nom_bis = fichier.readlines()
            long_new_nom_bis = len(new_nom_bis)

        i = 0
        j = 0
        for item in nom_cat:
            if long_new_nom_bis == 0:
                break
            if new_nom_bis[j] == nom_cat[i] + '\n':
                nom_cat_base[i].select()
                j += 1
                if j >= long_new_nom_bis:
                    break
            i += 1

        fichier.close()

        def fermer_fenetre():
            fenetrecatjuges.destroy()
        #retour page fentrecatjuges
        retour = tkinter.Button(fenetrecatjuges, text="RETOUR ", width=15, height=2, fg='blue', font=("Courrier", 13),bg='white', command = fermer_fenetre)
        retour.place(x=650, y=600)

    # Affichage des titres des informations
    labelprep = tkinter.Label(fenetreprep, text="MODIFICATIONS / AJOUTS :", width=25, height=2, fg='blue', bg='gray',font=("Courrier", 12));labelprep.place(x=460, y=10)
    labelprep = tkinter.Label(fenetreprep, text="Compétition :", width=12, height=3, fg='yellow', bg='gray',font=("Courrier", 12));labelprep.place(x=15, y=40)
    labelprep = tkinter.Label(fenetreprep, text="Organisateur :", width=10, height=3, fg='yellow', bg='gray',font=("Courrier", 12));labelprep.place(x=24, y=90)
    labelprep = tkinter.Label(fenetreprep, text="Lieu :", width=6, height=3, fg='yellow', bg='gray',font=("Courrier", 12));labelprep.place(x=15, y=140)
    labelprep = tkinter.Label(fenetreprep, text="Date :", width=5, height=3, fg='yellow', bg='gray',font=("Courrier", 12));labelprep.place(x=20, y=190)

    # bouton ouverture fenêtre juges et catégories
    labelprep = tkinter.Button(fenetreprep, text="Enregistrement juges et catégories", width=30, height=3, fg='yellow', bg='tomato',font=("Courrier", 12), command = cat_juges);labelprep.place(x=18, y=330)

    # lire dans le fichier txt
    with open("donnees_concours.txt", "r") as fichier:
        competition = fichier.readline()
        organisation = fichier.readline()
        lieuconcours = fichier.readline()
        dateconcours = fichier.readline()

    # affichages des données du fichier txt
    label_competition = tkinter.Label(fenetreprep, text=competition, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_competition.place(x=150, y=50)
    label_organisation = tkinter.Label(fenetreprep, text=organisation, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_organisation.place(x=150, y=100)
    label_lieuconcours = tkinter.Label(fenetreprep, text=lieuconcours, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_lieuconcours.place(x=150, y=150)
    label_dateconcours = tkinter.Label(fenetreprep, text=dateconcours, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_dateconcours.place(x=150, y=200)

    fichier.close()

    def fermer_fenetre():
        fenetreprep.destroy()

    # observateur données principale du concours
    def update_label(*args):
        # lire dans le fichier txt
        with open("donnees_concours.txt", "r") as fichier:
            competition = fichier.readline()
            organisation = fichier.readline()
            lieuconcours = fichier.readline()
            dateconcours = fichier.readline()

        var_label_compet.set(var_entry_compet.get())
        if var_label_compet.get()!="":
            compet = var_label_compet.get()
        else:
            compet=competition

        var_label_orga.set(var_entry_orga.get())
        if var_label_orga.get()!="":
            orga = var_label_orga.get()
        else:
            orga=organisation

        var_label_lieu.set(var_entry_lieu.get())
        if var_label_lieu.get()!="":
            lieu = var_label_lieu.get()
        else:
            lieu=lieuconcours

        var_label_date.set(var_entry_date.get())
        if var_label_date.get()!="":
            date = var_label_date.get()
        else:
            date=dateconcours

        # écrire dans le fichier texte
        with open("donnees_concours.txt", "w") as fichier:
            fichier.write(compet+'\n')
            fichier.write(orga+'\n')
            fichier.write(lieu+'\n')
            fichier.write(date+'\n')

        # effacer les lignes vides
        for line_number, line in enumerate(fileinput.input('donnees_concours.txt', inplace=1)):
            if len(line) > 1:
                sys.stdout.write(line)

        fichier.close()
        # champs d'entrée pour enregistrer nouvelles données compétition
    var_entry_compet = tkinter.StringVar()
    var_entry_compet.trace("w", update_label)
    entry_compet = tkinter.Entry(fenetreprep, textvariable=var_entry_compet, font=("Courrier", 12), width=30);entry_compet.place(x=460, y=60)
    var_entry_orga = tkinter.StringVar()
    var_entry_orga.trace("w", update_label)
    entry_orga = tkinter.Entry(fenetreprep, textvariable=var_entry_orga, font=("Courrier", 12), width=30);entry_orga.place(x=460, y=110)
    var_entry_lieu = tkinter.StringVar()
    var_entry_lieu.trace("w", update_label)
    entry_lieu = tkinter.Entry(fenetreprep, textvariable=var_entry_lieu, font=("Courrier", 12), width=30);entry_lieu.place(x=460, y=160)
    var_entry_date = tkinter.StringVar()
    var_entry_date.trace("w", update_label)
    entry_date = tkinter.Entry(fenetreprep, textvariable=var_entry_date, font=("Courrier", 12), width=30);entry_date.place(x=460, y=210)

    # label afficher nouvelles données (sous les champs d'entrée)
    var_label_compet = tkinter.StringVar()
    label_compet = tkinter.Label(fenetreprep, textvariable=var_label_compet, width=30, height=1, fg='blue', bg='gray',font=("Courrier", 12));label_compet.place(x=450, y=82)
    var_label_orga = tkinter.StringVar()
    label_orga = tkinter.Label(fenetreprep, textvariable=var_label_orga, width=30, height=1, fg='blue', bg='gray',font=("Courrier", 12));label_orga.place(x=450, y=134)
    var_label_lieu = tkinter.StringVar()
    label_lieu = tkinter.Label(fenetreprep, textvariable=var_label_lieu, width=30, height=1, fg='blue', bg='gray',font=("Courrier", 12));label_lieu.place(x=450, y=186)
    var_label_date = tkinter.StringVar()
    label_date = tkinter.Label(fenetreprep, textvariable=var_label_date, width=30, height=1, fg='blue', bg='gray',font=("Courrier", 12));label_date.place(x=450, y=238)

    # début fonction du bouton enregistrer
    def preparation_concours():

        # lire dans le fichier txt
        with open("donnees_concours.txt", "r") as fichier:
            competition = fichier.readline()
            organisation = fichier.readline()
            lieuconcours = fichier.readline()
            dateconcours = fichier.readline()

        # Affiches les données enregistrées dans le fichier texte
        label_competition = tkinter.Label(fenetreprep, text=competition, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_competition.place(x=150, y=50)
        label_organisation = tkinter.Label(fenetreprep, text=organisation, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_organisation.place(x=150, y=100)
        label_lieuconcours = tkinter.Label(fenetreprep, text=lieuconcours, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_lieuconcours.place(x=150, y=150)
        label_dateconcours = tkinter.Label(fenetreprep, text=dateconcours, width=30, height=3, fg='yellow', bg='gray',font=("Courrier", 12));label_dateconcours.place(x=150, y=200)

        fichier.close()
    # fin fonction du bouton enregistrer

    # bouton retour
    retour = tkinter.Button(fenetreprep, text="RETOUR", width=12, height=2, font=("Courrier", 10), bg='cornflowerblue',command=fermer_fenetre);retour.place(x=750, y=350)
    # bouton enregistrer
    enr = tkinter.Button(fenetreprep, text="ENREGISTRER", width=15, height=3, font=("Courrier", 12),bg='cornflowerblue', command=preparation_concours);enr.place(x=520, y=330)

    #def question():
        #messagebox.showerror("Pas prêt", parent = maincadre)

# titre fenêtre principale
titre = tkinter.Label(mainapp, text="Bienvenue sur l'application", font=("Courrier", 30), bg='sandybrown')
titre.pack(side="top")
maincadre = tkinter.LabelFrame(mainapp, text="MENU PRINCIPAL", bd=5)
maincadre.place(x=50, y=70)

# boutons principaux fenêtre principale
long = 20
haut = 5
compet = tkinter.Button(maincadre, text="Compétition", width=long, height=haut, font=("Courrier", 15), bg='#339900',command=liste_passage);compet.grid(row=2, column=0)
#arch = tkinter.Button(maincadre, text="Archives", width=long, height=haut, font=("Courrier", 15), bg='lightseagreen',command=question);arch.grid(row=4, column=0)
#rech = tkinter.Button(maincadre, text="Recherches", width=long, height=haut, font=("Courrier", 15), bg='#66FFFF',command=question);rech.grid(row=3, column=0)
prep = tkinter.Button(maincadre, text="Préparation", width=long, height=haut, font=("Courrier", 15), bg='lime',command=preparation);prep.grid(row=1, column=0)

# bouton quitter
def fermer_fenetre():
    mainapp.destroy()
quitter = tkinter.Button(mainapp, text="QUITTER", width=12, height=3, font=("Courrier", 12), bg='#FFCC99',command=fermer_fenetre)
quitter.place(x=940, y=630)

"""
Lignes 195/558/680 et 787 - affichages résultats passages, championnat et coupe

"""

mainapp.mainloop()