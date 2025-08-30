# Script Feuille d'Accueil – Google Sheets (Apps Script)

## 📌 Objectif
Ce script ajoute une **feuille d'accueil moderne** dans ton Google Sheets, avec :
- Un menu personnalisé **BDS Hassan** dans la barre d’outils
- Des tuiles de statistiques (élèves, CA, relances)
- Des boutons de navigation vers les autres feuilles
- Accès rapide aux mois
- Analyse des forfaits
- Boutons "Retour Accueil" sur chaque feuille

---

## ⚙️ Installation
1. Ouvre ton classeur Google Sheets (BdS).  
2. Menu : **Extensions → Apps Script**.  
3. Crée un nouveau projet, puis copie-colle le fichier [`Code.gs`](./Code.gs) dans l’éditeur.  
4. Enregistre (`Ctrl+S`) et recharge ton Google Sheets.  
5. Un nouveau menu **BDS Hassan** apparaît dans la barre.  

---

## 🚀 Utilisation
- **Créer Feuille Accueil** → génère ou régénère la page d’accueil avec les tuiles et raccourcis.  
- **Ajouter Boutons Retour** → ajoute un bouton “Retour Accueil” sur chaque feuille.  
- **Actualiser Interface** → met à jour la date en bas de la feuille d’accueil.  

---

## 📂 Prérequis (noms d’onglets attendus)
- `base données clients` (colonne B = élèves)  
- `Requete`  
- `Suiviannuel`  
- `Relances`  
- `paramètres`  
- Mois : `Aout`, `Septembre`, `Octobre`, …, `Juillet`

---

## 🔮 Améliorations possibles
- Calcul automatique du **CA mensuel** à partir des feuilles mensuelles.  
- Boutons de mois créés uniquement si l’onglet existe.  
- Gestion d’erreurs plus robuste si une feuille est absente.  

---

✍️ Auteur : Hassan (BdS) – sauvegardé le 30/08/2025
