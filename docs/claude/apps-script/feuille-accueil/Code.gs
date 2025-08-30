// SCRIPT FEUILLE ACCUEIL - DESIGN NEO MODERNE
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('BDS Hassan')
      .addItem('Créer Feuille Accueil', 'creerFeuilleAccueilNeo')
      .addItem('Ajouter Boutons Retour', 'ajouterBoutonsRetour')
      .addItem('Actualiser Interface', 'actualiserInterface')
      .addToUi();
}

function creerFeuilleAccueilNeo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Créer ou vider la feuille Accueil
  var accueil;
  try {
    accueil = ss.getSheetByName('Accueil');
    accueil.clear();
  } catch (e) {
    accueil = ss.insertSheet('Accueil', 0);
  }
  
  // Configuration avancée
  accueil.setRowHeights(1, 60, 55);
  accueil.setColumnWidths(1, 15, 160);
  
  // EN-TÊTE STYLE NEO
  accueil.getRange('A1:O5').merge();
  accueil.getRange('A1').setValue('BdS Hassan - Interface Moderne\nGestion des Cours Particuliers 2025/2026');
  accueil.getRange('A1')
    .setFontSize(22)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#1565c0')
    .setFontColor('white')
    .setWrap(true);
  
  // STATISTIQUES TEMPS RÉEL
  accueil.getRange('A7').setValue('STATISTIQUES TEMPS RÉEL')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1565c0');
  
  creerCarteNeo(accueil, 'A9:D12', 'Élèves Inscrits', '=COUNTA(INDIRECT("base données clients!B:B"))-1', '#4caf50');
  creerCarteNeo(accueil, 'F9:I12', 'CA Août 2025', '75,60€', '#1976d2');
  creerCarteNeo(accueil, 'K9:N12', 'Relances Actives', '=COUNTIF(Relances!D:D,"<0")', '#f44336');
  
  // NAVIGATION PRINCIPALE
  accueil.getRange('A15').setValue('NAVIGATION PRINCIPALE')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1565c0');
  
  // Ligne 1
  creerBoutonNeo(accueil, 'A17:D20', 'Tableau de Bord', 'Vue d\'ensemble et graphiques', '#4caf50', 'Tableau de Bord');
  creerBoutonNeo(accueil, 'F17:I20', 'Consulter un Client', 'Recherche dans Requete', '#2196f3', 'Requete');
  creerBoutonNeo(accueil, 'K17:N20', 'Suivi Annuel', 'Évolution annuelle', '#ff9800', 'Suiviannuel');
  
  // Ligne 2  
  creerBoutonNeo(accueil, 'A22:D25', 'Gestion Élèves', 'Base de données complète', '#9c27b0', 'base données clients');
  creerBoutonNeo(accueil, 'F22:I25', 'Relances', 'Gestion des impayés', '#f44336', 'Relances');
  creerBoutonNeo(accueil, 'K22:N25', 'Paramètres', 'Configuration système', '#607d8b', 'paramètres');
  
  // OUTILS AVANCÉS
  accueil.getRange('A28').setValue('OUTILS AVANCÉS')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1565c0');
  
  creerBoutonNeo(accueil, 'A30:D33', 'Éditer Facture BdS', 'Génération factures (à venir)', '#795548', null);
  creerBoutonNeo(accueil, 'F30:I33', 'Saisie des Cours', 'Voir menu déroulant →', '#3f51b5', null);
  
  // Menu déroulant des mois
  accueil.getRange('K30').setValue('Choisir le mois :')
    .setFontWeight('bold')
    .setFontColor('#1565c0');
  
  var mois = ['Aout', 'Septembre', 'Octobre', 'Novembre', 'Décembre', 'Janvier', 'Fevrier', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet'];
  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(mois)
    .setAllowInvalid(false)
    .build();
  
  accueil.getRange('K31:N32').merge();
  accueil.getRange('K31').setDataValidation(validation);
  accueil.getRange('K31').setValue('Aout');
  accueil.getRange('K31')
    .setBackground('#e3f2fd')
    .setBorder(true, true, true, true, true, true, '#1976d2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold')
    .setFontColor('#1976d2')
    .setFontSize(12);
  
  // Boutons directs pour chaque mois
  accueil.getRange('A35').setValue('ACCÈS RAPIDE MOIS')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1565c0');
  
  // Créer 12 boutons mois en 3 lignes de 4
  var moisBoutons = [
    ['Aout', 'Septembre', 'Octobre', 'Novembre'],
    ['Décembre', 'Janvier', 'Fevrier', 'Mars'],  
    ['Avril', 'Mai', 'Juin', 'Juillet']
  ];
  
  for (var i = 0; i < moisBoutons.length; i++) {
    for (var j = 0; j < moisBoutons[i].length; j++) {
      var startRow = 37 + (i * 3);
      var startCol = 1 + (j * 3.5);
      var endCol = startCol + 2;
      creerBoutonMois(accueil, startRow, startCol, endCol, moisBoutons[i][j]);
    }
  }
  
  // ANALYSE FORFAITS
  accueil.getRange('A47').setValue('ANALYSE FORFAITS BDS')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1565c0');
  
  var forfaits = [
    ['Forfait', 'Tarif Brut', 'Net (84%)', 'Commission (16%)', 'Status'],
    ['Groupe', 15, '=B50*0.84', '=B50*0.16', 'Actif'],
    ['Triple', 20, '=B51*0.84', '=B51*0.16', 'Actif'], 
    ['Duo', 25, '=B52*0.84', '=B52*0.16', 'Actif'],
    ['Individuel', 35, '=B53*0.84', '=B53*0.16', 'Actif']
  ];
  
  accueil.getRange(49, 1, forfaits.length, forfaits[0].length).setValues(forfaits);
  
  // Style tableau moderne
  accueil.getRange('A49:E53').setBorder(true, true, true, true, true, true);
  accueil.getRange('A49:E49')
    .setBackground('#1976d2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  accueil.getRange('E50:E53')
    .setBackground('#4caf50')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Format monétaire
  accueil.getRange('B50:D53').setNumberFormat('#,##0.00 "€"');
  
  // PIED DE PAGE
  accueil.getRange('A56:N56').merge();
  accueil.getRange('A56').setValue('Interface NEO BDS • Dernière mise à jour : ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'));
  accueil.getRange('A56')
    .setHorizontalAlignment('center')
    .setBackground('#e3f2fd')
    .setFontColor('#1976d2')
    .setFontStyle('italic');
  
  SpreadsheetApp.getUi().alert('Feuille Accueil NEO créée !\n\nNavigation par liens hypertexte activée.\nUtilisez "Ajouter Boutons Retour" pour finaliser.');
}

function creerCarteNeo(feuille, plage, titre, valeur, couleur) {
  feuille.getRange(plage).merge();
  var cellule = feuille.getRange(plage.split(':')[0]);
  
  if (typeof valeur === 'string' && valeur.startsWith('=')) {
    cellule.setFormula(valeur);
  } else {
    cellule.setValue(valeur);
  }
  
  cellule
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(couleur)
    .setFontColor('white');
  
  // Bordures épaisses pour effet 3D
  feuille.getRange(plage).setBorder(true, true, true, true, false, false, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // Label au-dessus
  var labelRange = plage.split(':')[0];
  var labelRow = parseInt(labelRange.match(/\d+/)[0]) - 1;
  var labelCol = labelRange.match(/[A-Z]+/)[0];
  
  feuille.getRange(labelCol + labelRow).setValue(titre)
    .setFontSize(11)
    .setFontWeight('bold')
    .setFontColor('#718096')
    .setHorizontalAlignment('center');
}

function creerBoutonNeo(feuille, plage, titre, description, couleur, feuilleCible) {
  feuille.getRange(plage).merge();
  var cellule = feuille.getRange(plage.split(':')[0]);
  
  // Ajouter lien hypertexte si feuille cible existe
  if (feuilleCible) {
    var gid = obtenirGidFeuille(feuilleCible);
    if (gid !== 0) {
      var texteEscape = titre.replace(/"/g, '""') + '\\n' + description.replace(/"/g, '""');
      cellule.setFormula('=HYPERLINK("#gid=' + gid + '","' + texteEscape + '")');
    } else {
      cellule.setValue(titre + '\n' + description + '\n(Feuille introuvable)');
    }
  } else {
    cellule.setValue(titre + '\n' + description);
  }
  
  cellule
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(couleur)
    .setFontColor('white')
    .setWrap(true);
  
  // Bordures épaisses pour effet 3D
  feuille.getRange(plage).setBorder(true, true, true, true, false, false, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function creerBoutonMois(feuille, row, startCol, endCol, mois) {
  var plage = feuille.getRange(row, startCol, 2, endCol - startCol + 1);
  plage.merge();
  
  var gid = obtenirGidFeuille(mois);
  if (gid !== 0) {
    plage.setFormula('=HYPERLINK("#gid=' + gid + '","' + mois + '")');
  } else {
    plage.setValue(mois + '\n(Non trouvé)');
  }
  
  plage
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#3f51b5')
    .setFontColor('white');
  
  plage.setBorder(true, true, true, true, false, false, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function obtenirGidFeuille(nomFeuille) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var feuille = ss.getSheetByName(nomFeuille);
    return feuille.getSheetId();
  } catch (e) {
    return 0;
  }
}

function ajouterBoutonsRetour() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuilles = ss.getSheets();
  var accueilGid = obtenirGidFeuille('Accueil');
  var count = 0;
  
  feuilles.forEach(function(feuille) {
    var nom = feuille.getName();
    
    if (nom !== 'Accueil') {
      feuille.getRange('A1')
        .setFormula('=HYPERLINK("#gid=' + accueilGid + '","Retour Accueil")')
        .setFontSize(12)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBackground('#1565c0')
        .setFontColor('white')
        .setBorder(true, true, true, true, true, true, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK);
      
      feuille.setRowHeight(1, 35);
      count++;
    }
  });
  
  SpreadsheetApp.getUi().alert('Boutons retour ajoutés sur ' + count + ' feuilles !');
}

function actualiserInterface() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accueil = ss.getSheetByName('Accueil');
  
  if (accueil) {
    accueil.getRange('A56').setValue('Interface NEO BDS • Dernière mise à jour : ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'));
    SpreadsheetApp.getUi().alert('Interface actualisée !');
  }
}
