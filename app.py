import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import shutil
import smtplib
from email.message import EmailMessage
from tkinter import simpledialog
from fpdf import FPDF
# --- CHARTE GRAPHIQUE (Style Crédit Agricole) ---
CA_VERT = "#007A5E"
CA_FOND = "#F4F6F6"
CA_TEXTE = "#333333"
CA_BLANC = "#FFFFFF"

def quitter_proprement():
    if messagebox.askyesno("Déconnexion", "Êtes-vous sûr de vouloir fermer le terminal ?\nToute offre non exportée sera perdue."):
        fenetre.destroy()

# --- FONCTIONS DU BACKEND ---
def lancer_simulation():
    nom = entree_nom.get()
    cap = entree_capital.get()
    taux = entree_taux.get()
    duree = entree_duree.get()

    if not all([nom, cap, taux, duree]):
        messagebox.showwarning("Saisie Incomplète", "Veuillez remplir tous les champs du prêt.")
        return

    if not nom.endswith(".csv"):
        nom += ".csv"
    
    entree_nom.delete(0, tk.END)
    entree_nom.insert(0, nom)

    donnees = f"{nom}\n{cap}\n{taux}\n{duree}\nN\n"
    chemin_exe = os.path.join(os.getcwd(), "calc_pret.exe")

    try:
        processus = subprocess.run([chemin_exe], input=donnees, text=True, capture_output=True)
        if processus.returncode == 0:
            generer_graphique(nom)
            bouton_pdf.config(state=tk.NORMAL)
            bouton_excel.config(state=tk.NORMAL)
            bouton_mail.config(state=tk.NORMAL)
        else:
            messagebox.showerror("Erreur", "Le moteur COBOL a échoué.")
    except Exception as e:
        messagebox.showerror("Erreur Système", str(e))

def generer_graphique(fichier_csv):
    # Nettoyage de l'ancien graphique s'il existe
    for widget in frame_graphique.winfo_children():
        widget.destroy()

    try:
        # Lecture du CSV généré par COBOL
        df = pd.read_csv(fichier_csv, sep=";", header=None, 
                         names=["Mois", "Mensualité", "Intérêts", "Amortissement", "Capital Restant", "Vide"])
        
        # Création de la figure
        fig, ax1 = plt.subplots(figsize=(6, 4), dpi=100)
        fig.patch.set_facecolor(CA_FOND)
        
        ax1.plot(df["Mois"], df["Capital Restant"], color=CA_VERT, linewidth=2, label="Capital Restant")
        ax1.set_xlabel("Mois")
        ax1.set_ylabel("Euros (€)")
        ax1.set_title(f"Évolution de l'amortissement - {fichier_csv}")
        ax1.grid(True, linestyle='--', alpha=0.6)
        
        # Intégration dans Tkinter
        canvas = FigureCanvasTkAgg(fig, master=frame_graphique)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
    except Exception as e:
        messagebox.showerror("Erreur Graphique", f"Impossible de lire les données : {e}")

# --- FONCTIONS D'EXPORTATION ---
def exporter_excel():
    csv_actuel = entree_nom.get()
    if not os.path.exists(csv_actuel):
        return
    fichier_dest = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Fichiers Excel", "*.xlsx")])
    if fichier_dest:
        df = pd.read_csv(csv_actuel, sep=";", header=None, names=["Mois", "Mensualité", "Intérêts", "Amortissement", "Capital Restant", "Vide"])
        df.drop(columns=["Vide"], inplace=True)
        df.to_excel(fichier_dest, index=False, engine='openpyxl')
        messagebox.showinfo("Export Excel", "Tableau d'amortissement converti et sauvegardé avec succès.")

def exporter_pdf():
    # 1. Capture des données commerciales de l'interface
    csv_actuel = entree_nom.get()
    capital_saisi = entree_capital.get()
    taux_saisi = entree_taux.get()
    duree_saisie = entree_duree.get()
    
    if not csv_actuel:
        messagebox.showwarning("Attention", "Aucun nom de fichier n'est renseigné dans l'interface.")
        return
        
    if not os.path.exists(csv_actuel):
        messagebox.showwarning("Erreur", f"Le fichier {csv_actuel} est introuvable. Veuillez relancer le calcul.")
        return
        
    dossier_app = os.path.dirname(os.path.abspath(__file__))
    fichier_dest = filedialog.asksaveasfilename(
        initialdir=dossier_app,
        defaultextension=".pdf", 
        filetypes=[("Fichiers PDF", "*.pdf")]
    )
    
    if not fichier_dest:
        return 
        
    if not fichier_dest.lower().endswith('.pdf'):
        fichier_dest += '.pdf'
        
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # 2. En-tête institutionnel
        pdf.set_font("Arial", style="B", size=16)
        pdf.set_text_color(0, 122, 94) # Vert Crédit Agricole
        pdf.cell(0, 10, text="PROPOSITION COMMERCIALE DE FINANCEMENT", align='C')
        pdf.ln(15)
        
        # Lecture des données pour extraire la mensualité exacte du premier mois
        df = pd.read_csv(csv_actuel, sep=";", header=None, names=["Mois", "Mensualite", "Interets", "Amortissement", "Restant", "Vide"])
        mensualite_estimee = df['Mensualite'].iloc[0]
        
        # 3. Bloc Récapitulatif du Prêt
        pdf.set_font("Arial", size=11)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(0, 8, text=f"Montant du capital emprunte : {capital_saisi} EUR", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 8, text=f"Taux d'intérêt débiteur annuel : {taux_saisi} %", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 8, text=f"Durée totale du crédit : {duree_saisie} mois", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 8, text=f"Échéance mensuelle (hors assurance) : {mensualite_estimee} EUR", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(10)
        
        # 4. Tableau d'Amortissement
        pdf.set_font("Arial", style="B", size=10)
        largeurs = [20, 40, 40, 40, 40]
        colonnes = ["Mois", "Mensualite", "Interets", "Amorti", "Restant"]
        
        for i in range(len(colonnes)):
            pdf.cell(largeurs[i], 8, text=colonnes[i], border=1, align='C')
        pdf.ln()
        
        pdf.set_font("Arial", size=9)
        for index, row in df.iterrows():
            pdf.cell(largeurs[0], 6, text=str(row['Mois']), border=1, align='C')
            pdf.cell(largeurs[1], 6, text=str(row['Mensualite']), border=1, align='R')
            pdf.cell(largeurs[2], 6, text=str(row['Interets']), border=1, align='R')
            pdf.cell(largeurs[3], 6, text=str(row['Amortissement']), border=1, align='R')
            pdf.cell(largeurs[4], 6, text=str(row['Restant']), border=1, align='R')
            pdf.ln()
            
        # 5. Bloc de Signatures (fin du document)
        pdf.ln(15)
        pdf.set_font("Arial", style="B", size=11)
        pdf.cell(95, 10, text="Signature de l'Emprunteur", align='L')
        pdf.cell(95, 10, text="Cachet et Signature de la Banque", new_x="LMARGIN", new_y="NEXT", align='R')
        pdf.set_font("Arial", size=9)
        pdf.cell(95, 5, text="(Précédée de la mention 'Lu et approuvé')",align='L')
            
        pdf.output(fichier_dest)
        messagebox.showinfo("Export PDF", f"Proposition commerciale générée avec succès :\n{fichier_dest}")
        
    except Exception as e:
        messagebox.showerror("Erreur PDF", f"La génération a échoué : {e}")

def envoyer_mail():
    csv_actuel = entree_nom.get()
    if not os.path.exists(csv_actuel):
        messagebox.showerror("Erreur", "Aucune simulation à envoyer.")
        return

    # Fenêtre pop-up pour demander l'email du client
    email_client = simpledialog.askstring("Envoi Email", "Saisissez l'adresse e-mail du client :")
    if not email_client:
        return

    try:
        # --- À CONFIGURER ---
        mon_email = "TON_ADRESSE@gmail.com"
        mot_de_passe_app = "TON_MOT_DE_PASSE_D_APPLICATION" 

        # Structuration de l'e-mail
        msg = EmailMessage()
        msg['Subject'] = f"Simulation de prêt bancaire - {csv_actuel}"
        msg['From'] = mon_email
        msg['To'] = email_client
        msg.set_content("Bonjour,\n\nVeuillez trouver ci-joint votre tableau d'amortissement détaillé.\n\nCordialement,\nVotre Conseiller.")

        # Injection de la pièce jointe (CSV)
        with open(csv_actuel, 'rb') as f:
            msg.add_attachment(f.read(), maintype='text', subtype='csv', filename=csv_actuel)

        # Connexion SSL au serveur Gmail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as serveur:
            serveur.login(mon_email, mot_de_passe_app)
            serveur.send_message(msg)

        messagebox.showinfo("Succès", f"Simulation envoyée à {email_client}")

    except Exception as e:
        messagebox.showerror("Erreur Réseau", f"Échec de la connexion SMTP : {e}")
# --- INTERFACE UTILISATEUR (Frontend) ---
fenetre = tk.Tk()
fenetre.title("Terminal Conseiller - Simulation de Prêt")
fenetre.geometry("900x550")
fenetre.configure(bg=CA_FOND)

# Cadre principal divisé en deux (Saisie à gauche, Graphique à droite)
frame_gauche = tk.Frame(fenetre, bg=CA_BLANC, width=300, relief=tk.FLAT, bd=2)
frame_gauche.pack(side=tk.LEFT, fill=tk.Y, padx=20, pady=20)
frame_gauche.pack_propagate(False)

frame_droite = tk.Frame(fenetre, bg=CA_FOND)
frame_droite.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=20, pady=20)

# Titre Panneau Saisie
tk.Label(frame_gauche, text="NOUVELLE OFFRE", font=("Arial", 14, "bold"), bg=CA_BLANC, fg=CA_VERT).pack(pady=(20, 20))

# Champs de saisie
champs = [("Nom du dossier :", entree_nom := tk.Entry(frame_gauche, font=("Arial", 10))),
          ("Capital emprunté (€) :", entree_capital := tk.Entry(frame_gauche, font=("Arial", 10))),
          ("Taux annuel (%) :", entree_taux := tk.Entry(frame_gauche, font=("Arial", 10))),
          ("Durée (mois) :", entree_duree := tk.Entry(frame_gauche, font=("Arial", 10)))]

for label_text, entry_widget in champs:
    tk.Label(frame_gauche, text=label_text, font=("Arial", 10), bg=CA_BLANC, fg=CA_TEXTE).pack(anchor="w", padx=20, pady=(10, 2))
    entry_widget.pack(fill=tk.X, padx=20, pady=2)

# Bouton d'exécution
tk.Button(frame_gauche, text="CALCULER L'AMORTISSEMENT", font=("Arial", 10, "bold"), bg=CA_VERT, fg=CA_BLANC, 
          relief=tk.FLAT, command=lancer_simulation, pady=8).pack(fill=tk.X, padx=20, pady=30)

# Cadre pour les boutons d'export (désactivés par défaut)
frame_actions = tk.Frame(frame_gauche, bg=CA_BLANC)
frame_actions.pack(fill=tk.X, padx=20, side=tk.BOTTOM, pady=20)

bouton_pdf = tk.Button(frame_actions, text="Exporter PDF", state=tk.DISABLED, command=exporter_pdf)
bouton_pdf.pack(fill=tk.X, pady=2)
bouton_excel = tk.Button(frame_actions, text="Exporter Excel", state=tk.DISABLED, command=exporter_excel)
bouton_excel.pack(fill=tk.X, pady=2)
bouton_mail = tk.Button(frame_actions, text="Envoyer par Mail", state=tk.DISABLED, command=envoyer_mail)
bouton_mail.pack(fill=tk.X, pady=2)
tk.Button(frame_actions, text="Quitter", bg="#D32F2F", fg=CA_BLANC, command=quitter_proprement).pack(fill=tk.X, pady=(15,0))

# Espace d'affichage graphique
tk.Label(frame_droite, text="VISUALISATION DU CRÉDIT", font=("Arial", 14, "bold"), bg=CA_FOND, fg=CA_VERT).pack(anchor="w")
frame_graphique = tk.Frame(frame_droite, bg=CA_BLANC, relief=tk.SOLID, bd=1)
frame_graphique.pack(fill=tk.BOTH, expand=True, pady=10)

fenetre.mainloop()