# 🏦 Terminal Bancaire - Simulateur de Crédit Hybride (Python & COBOL)

Application de bureau professionnelle conçue pour les conseillers bancaires, combinant la puissance de calcul d'un moteur **COBOL** bas-niveau avec l'ergonomie d'une interface graphique moderne en **Python**.

---

## 🚀 Fonctionnalités Clés

* **Moteur de Calcul COBOL :** Traitement mathématique brut et performant des amortissements de prêts via un exécutable compilé (`calc_pret.exe`).
* **Interface Graphique Moderne (Python / Tkinter) :** Saisie ergonomique des paramètres de l'offre (Capital, Taux, Durée).
* **Visualisation Dynamique :** Génération instantanée de courbes d'évolution du capital restant dû grâce à **Matplotlib**.
* **Génération de Contrats PDF :** Exportation d'une proposition commerciale formelle et paginée (`fpdf2`) incluant le tableau d'amortissement complet et les blocs de signatures.
* **Architecture de Distribution Propre :** Livrable packagé en mode répertoire pour un lancement instantané sans latence de décompression.

---

## 🛠️ Stack Technique

* **Backend :** GnuCOBOL (`cobc`)
* **Frontend & Logique :** Python (Tkinter, Pandas, Matplotlib, FPDF2)
* **Packaging :** PyInstaller & PowerShell

---

## 📂 Structure du Projet

```text
Projet_bancaire_cobol/
│
├── app.py                # Interface graphique et orchestration Python
├── calc_pret.cbl         # Code source du moteur de calcul COBOL
├── calc_pret.exe         # Exécutable du moteur COBOL compilé
├── deploy.ps1            # Script d'automatisation et de packaging
└── dist/
    └── Terminal_Conseiller/
        ├── app.exe       # Exécutable principal de l'application
        └── calc_pret.exe # Moteur COBOL embarqué