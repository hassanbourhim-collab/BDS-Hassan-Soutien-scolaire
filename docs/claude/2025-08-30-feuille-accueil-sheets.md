# Session Claude – Feuille d'accueil Google Sheets (30/08/2025)
revoir la taille des boutons, les boutons affichent #erreur, tous
**Objectif :** créer une feuille d’accueil avec tuiles de stats, navigation interne, raccourcis mois, et mini-analyse des forfaits.
créer une synthese et graphique
**Décisions prises :**
- Menu personnalisé “BDS Hassan” avec 3 actions (créer / boutons retour / refresh).
- Tuiles KPI : élèves (base données clients), CA du mois (valeur provisoire), relances (Relances!D:D).
- Navigation : boutons vers “Requete”, “Suiviannuel”, “base données clients”, etc.
- Accès rapide Mois : 12 boutons + liste déroulante.

**À améliorer plus tard :**
- Rendre les mois dynamiques selon les onglets existants.
- Sécuriser les formules si une feuille manque.
- Remplacer “CA Août 2025” par une formule réelle.

**Fichiers liés :**
- `apps-script/feuille-accueil/Code.gs`
