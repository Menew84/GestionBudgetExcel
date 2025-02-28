# Tableau de Gestion de Budget Personnel

## Table of Contents
1. [Prérequis](#prérequis)  
2. [Génération du Fichier Excel](#génération-du-fichier-excel)  
3. [Structure du Fichier Excel](#structure-du-fichier-excel)  
   1. [Page de Garde](#page-de-garde)  
   2. [Feuilles Mensuelles](#feuilles-mensuelles)  
4. [Conversion en .xlsm et Ajout des Macros VBA](#conversion-en-xlsm-et-ajout-des-macros-vba)  
5. [Activation des Macros](#activation-des-macros)  
6. [Résumé](#résumé)

---

Ce projet vous permet de générer automatiquement un tableau Excel pour gérer votre budget personnel. Le système se compose de deux parties :

1. **Script Python (`generer_budget.py`)**  
   Ce script génère un fichier Excel structuré (au format `.xlsx`) qui contient :  
   - Une **Page de Garde** avec un récapitulatif annuel et un graphique comparatif.  
   - 12 **Feuilles Mensuelles** (une par mois) pour saisir vos revenus et vos dépenses.  
     - **En-tête** (ligne 1) affichant le mois et l’année.  
     - **Solde Initial** (cellule B2) qui est 0 pour le premier mois et est reporté (via une référence à la cellule C4 du mois précédent) pour les mois suivants.  
     - **Zone Revenus** (colonnes A à D) comprenant :  
       - **A** : Date (format `jj/mm/aaaa`)  
       - **B** : Catégorie/Desc.  
       - **C** : Montant (formaté en euros)  
       - **D** : "Payé ?" – (destinée à recevoir des cases à cocher via macros VBA)  
     - **Zone Dépenses** (colonnes F à I) comprenant :  
       - **F** : Date (format `jj/mm/aaaa`)  
       - **G** : Catégorie/Desc.  
       - **H** : Montant  
         *Les montants sont saisis en positif, mais affichés avec un signe “–” devant et en rouge grâce à un format personnalisé.*  
       - **I** : "Payé ?" – (destinée à recevoir des cases à cocher via macros VBA)  
     - **Totaux et Solde Mensuel** calculés en haut :  
       - **Total Revenus** en cellule **C3** (somme des montants de la zone Revenus, lignes 6 à 55).  
       - **Total Dépenses** en cellule **H3** (somme des montants de la zone Dépenses, affiché avec un “–” grâce au format personnalisé).  
       - **Solde Mensuel** en cellule **C4** = *Solde Initial + Total Revenus – Total Dépenses*.  
     - Des **cellules de synthèse** (ligne 60 : C60, D60, E60) reprennent respectivement le Total Revenus, le Total Dépenses et le Solde Mensuel. Cette ligne est masquée et sert à alimenter la Page de Garde.

   La **Page de Garde** comprend :  
   - Un **titre** en A1:D1 indiquant "BUDGET PERSONNEL 2025 - Récapitulatif Annuel".  
   - Un **tableau récapitulatif** (lignes 4 à 15) listant pour chaque mois (ex. "Janvier 2025", "Février 2025", etc.) :  
     - **Total Revenus** (colonne B) – récupéré depuis la cellule de synthèse C60 de la feuille du mois.  
     - **Total Dépenses** (colonne C) – récupéré depuis la cellule de synthèse D60 de la feuille du mois.  
     - **Solde** (colonne D) – récupéré depuis la cellule de synthèse E60 de la feuille du mois.  
   - Une **ligne "TOTAL Annuel"** (ligne 16) qui affiche la somme des Totaux Revenus (colonne B), la somme des Totaux Dépenses (colonne C) et le solde final du dernier mois (colonne D, par exemple `=D15`).  
   - Un **graphique comparatif** (barres en mode "clustered") qui affiche pour chaque mois deux barres côte à côte : l’une pour le Total Revenus et l’autre pour le Total Dépenses.

2. **Code VBA (`macro_budget.vba`)**  
   Un module VBA (fourni séparément dans ce repository) permet d’insérer des cases à cocher interactives dans les colonnes "Payé ?" de chaque feuille mensuelle :  
   - Pour la zone **Revenus**, les cases à cocher seront insérées dans la colonne **D**.  
   - Pour la zone **Dépenses**, les cases à cocher seront insérées dans la colonne **I**.  
   Ces macros facilitent la gestion en vous permettant de cocher rapidement si un montant a été payé.

---

## Prérequis
[Prérequis](#)

- **Python 3** – [Téléchargez-le ici](https://www.python.org/downloads/).  
- La bibliothèque **openpyxl** – Installez-la avec :
  ```bash
  pip install openpyxl
  ```
- Microsoft Excel (version 2010 ou ultérieure).

---

## Génération du Fichier Excel
[Génération du Fichier Excel](#)

1. Clonez ou téléchargez ce repository.  
2. Ouvrez le fichier `generer_budget.py` dans votre éditeur préféré.  
3. Dans un terminal ou invite de commandes, placez-vous dans le dossier contenant le script et exécutez :
   ```bash
   python generer_budget.py
   ```
4. Le script génère un fichier **Budget_Personnel.xlsx** dans le même dossier.

---

## Structure du Fichier Excel
[Structure du Fichier Excel](#)

### Page de Garde
[Page de Garde](#)

- **Titre** (A1:D1) : "BUDGET PERSONNEL 2025 - Récapitulatif Annuel"  
- **Tableau récapitulatif** (lignes 4 à 15) pour chaque mois :  
  - **Total Revenus** (colonne B) – référencé via la cellule de synthèse C60 de la feuille du mois.  
  - **Total Dépenses** (colonne C) – référencé via la cellule de synthèse D60 de la feuille du mois.  
  - **Solde** (colonne D) – référencé via la cellule de synthèse E60 de la feuille du mois.  
- **Ligne "TOTAL Annuel"** (ligne 16) :  
  - Colonne B : Somme des Totaux Revenus.  
  - Colonne C : Somme des Totaux Dépenses.  
  - Colonne D : Le solde final du dernier mois (par exemple, `=D15`).  
- **Graphique Comparatif** : Affiche des barres "Revenus" et "Dépenses" côte à côte pour chaque mois.

### Feuilles Mensuelles
[Feuilles Mensuelles](#)

Chaque feuille mensuelle comporte :  
- **En-tête** (ligne 1, fusionnée de A1 à E1) : Affiche "JANVIER 2025", etc.  
- **Solde Initial** (cellule B2) :
  - Pour le premier mois, c’est 0.
  - Pour les mois suivants, il est reporté depuis la cellule C4 du mois précédent.
- **Zone Revenus** (colonnes A à D) :
  - **A** : Date (format `jj/mm/aaaa`)
  - **B** : Catégorie/Desc.
  - **C** : Montant (en euros)
  - **D** : "Payé ?" (destinée à recevoir des cases à cocher via macros VBA)
- **Zone Dépenses** (colonnes F à I) :
  - **F** : Date (format `jj/mm/aaaa`)
  - **G** : Catégorie/Desc.
  - **H** : Montant (saisi en positif mais affiché avec un “–” et en rouge via un format personnalisé)
  - **I** : "Payé ?" (pour les cases à cocher via macros VBA)
- **Totaux** :
  - **Total Revenus** en cellule C3 (somme des montants de la zone Revenus).
  - **Total Dépenses** en cellule H3 (somme des montants de la zone Dépenses).
  - **Solde Mensuel** en cellule C4 = *Solde Initial + C3 – H3*.
- **Cellules de synthèse** (ligne 60) :
  - **C60** = Total Revenus (C3)
  - **D60** = Total Dépenses (H3)
  - **E60** = Solde Mensuel (C4)
  Ces cellules sont masquées et servent à alimenter la Page de Garde.

---

## Conversion en .xlsm et Ajout des Macros VBA
[Conversion en .xlsm et Ajout des Macros VBA](#)

1. **Ouvrez** le fichier **Budget_Personnel.xlsx** dans Excel.  
2. Allez dans **Fichier → Enregistrer sous** et choisissez le format **Classeur Excel (prise en charge des macros) (.xlsm)**.  
3. Appuyez sur `Alt+F11` pour ouvrir l’éditeur VBA.  
4. Dans l’éditeur, insérez un nouveau module via **Insertion → Module**.  
5. **Collez** le contenu du fichier `macro_budget.vba` (fourni dans ce repository) dans le module.  
6. Enregistrez le classeur.  
7. Pour ajouter des cases à cocher, sélectionnez une feuille mensuelle (par exemple, "Janvier 2025"), puis dans l’onglet **Développeur → Macros**, exécutez :
   - `AjouterCheckboxRevenus` pour la colonne "Payé ?" des Revenus (colonne D).
   - `AjouterCheckboxDepenses` pour la colonne "Payé ?" des Dépenses (colonne I).

---

## Activation des Macros
[Activation des Macros](#)

1. Dans Excel, allez dans **Fichier → Options → Centre de gestion de la confidentialité → Paramètres des macros**.  
2. Sélectionnez **Activer toutes les macros** (ou choisissez un niveau de sécurité adapté) et activez l'accès au **modèle d'objet du projet VBA**.  
3. Redémarrez Excel si nécessaire.

---

## Résumé
[Résumé](#)

- Exécutez le script Python pour générer un fichier Excel structuré.  
- La **Page de Garde** affiche un récapitulatif annuel et un graphique comparatif avec des barres "Revenus" et "Dépenses" côte à côte.  
- Chaque **Feuille Mensuelle** permet de saisir vos revenus et dépenses, calcule automatiquement les totaux et le solde mensuel, et reporte le solde initial du mois précédent.  
- Les **cellules de synthèse** (ligne 60) sont masquées et servent à alimenter la Page de Garde.  
- Pour ajouter des cases à cocher interactives, convertissez le fichier en **.xlsm** et insérez le module VBA fourni.  

Ce projet vous offre un outil complet, esthétique et personnalisable pour gérer votre budget personnel.

**Bonne gestion de votre budget !**
