import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import random

mainapp = tkinter.Tk()
mainapp.title("mot magique")
mainapp.geometry("1500x1200")
mainapp.minsize(480, 360)
mainapp.config(background='sandybrown')

min = 0
max = 50
nbre_magique = random.randint(min, max)

def calcul_nbre():
    while not nbre_magique=="":
        try:
            reponse=input ("Trouves le nombre magique (entre 0 et 50) ")
            if int(reponse)==nbre_magique:
                print("BRAVO, tu as trouvé le nombre magique")
                return
            elif int(reponse)<nbre_magique:
                print("Ton nombre est trop petit")
            else :
                print("Ton nombre est trop grand")
        except :
            print("Il faut noter un nombre (pas de lettres")


def start_jeu():
    bouton_start_jeu.destroy()
    fenetre_calcul=tkinter.Toplevel()
    fenetre_calcul.title("mot magique")
    fenetre_calcul.geometry("1500x200")
    fenetre_calcul.config(background='blue')

    label_texte = tkinter.Label (fenetre_calcul, text="Trouves le nombre magique (entre 0 et 50) : ",width=50, height=5,bg='deeppink', font=("Courrier", 12));
    label_texte.place(x=1, y=10)

    label_entree=tkinter.Entry(fenetre_calcul,font=("Courrier", 12), width=10);
    label_entree.place(x=240, y=160)
    
    #calcul_nbre()


# bouton start
bouton_start_jeu = tkinter.Button(mainapp, text="Démarrer le jeu", width=20, height=5, font=("Courrier", 40), bg='#339900',command=start_jeu);bouton_start_jeu.place(x=450, y=100)

# bouton quitter
def fermer_fenetre():
    mainapp.destroy()
quitter = tkinter.Button(mainapp, text="QUITTER", width=12, height=3, font=("Courrier", 12), bg='#FFCC99',command=fermer_fenetre)
quitter.place(x=1350, y=850)












mainapp.mainloop()