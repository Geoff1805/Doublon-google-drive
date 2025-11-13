/*
 * SCRIPT AVANCÉ DE RECHERCHE DE DOUBLONS PAR CONTENU (v4.10.1 - Incrémentiel)
 *
 * ====================================================================
 * HISTORIQUE DES VERSIONS (A.B.C)
 * ====================================================================
 * v4.0.0 - Refonte majeure : analyse "Incrémentielle" avec [DB].
 * v4.1.0 - Ajout de la colonne "Dossier (Parent)".
 * v4.2.0 - Génération d'un Google Sheet "[ACTION] Suppression Doublons".
 * v4.3.0 - (FIX) "Argument too large": "liste de tâches" déplacée vers [TEMP].
 * v4.4.0 - (FIX) "runtime exited": Baisse de MAX_FILE_SIZE_BYTES (150Mo).
 * v4.5.0 - Ajout d'un fichier journal Google Sheet "[LOG]".
 * v4.6.0 - Ajout de la gestion des dossiers "/script/doublondrive/".
 * v4.9.2 - (FIX) Le problème est la TAILLE. MAX_FILE_SIZE_BYTES baissé à 45Mo.
 * v4.9.6 - (FIX) Corrections de syntaxe (apostrophes).
 * v4.10.0 - (FEATURE) Implémentation de toutes les demandes en attente :
 * 1. (FIX) La pause est TOUJOURS de 1 minute.
 * 2. (FEATURE) Ajout du "Réveil Intelligent" (lanceurNocturneIntelligent).
 * 3. (FEATURE) Ajout du nouveau Sheet [STATS] (Tableau de Bord, Historique, Formats, Top 100).
 * 4. (FEATURE) Ajout des compteurs de lots et logs améliorés.
 * v4.10.1 - (FIX) Suppression de la fonction 'INSTALLER_DECLENCHEUR_NOCTURNE' (devenue inutile).
 * ====================================================================
 */

// --- VARIABLES DE CONFIGURATION ---
const EMAIL_POUR_RAPPORT = "toto@hotmail.com";
const CHEMIN_DOSSIER = "script/doublondrive";
const NOM_SHEET_DB = `[DB] Hashes Fichiers Drive`;
const NOM_SHEET_ACTION = `[ACTION] Suppression Doublons`;
const NOM_SHEET_TEMP_TODO = `[TEMP] Fichiers à Traiter`; 
const NOM_SHEET_LOG = `[LOG] Journal Analyse Doublons`;
const NOM_SHEET_STATS = `[STATS] Tableau de Bord`; 

const MINUTES_ENTRE_LOTS = 1; // v4.10.0 : Pause unique de 1 minute
const TEMPS_MAX_EXECUTION_SECONDES = 270; // 4.5 minutes

// v4.9.2 : Limite à 45 Mo
const MAX_FILE_SIZE_BYTES = 47185920; // 45 Mo
const BATCH_SIZE_TRAITEMENT = 50; 

// Noms des propriétés de script (pour gérer l'état)
const PROP_ETAT_SCRIPT = 'ETAT_SCRIPT';
const PROP_CONTINUATION_TOKEN = 'CONTINUATION_TOKEN';
const PROP_SHEET_ID_DB = 'SHEET_ID_DB'; 
const PROP_SHEET_ID_TODO = 'SHEET_ID_TODO'; 
const PROP_SHEET_ID_LOG = 'SHEET_ID_LOG'; 
const PROP_SHEET_ID_STATS = 'SHEET_ID_STATS';
const PROP_DB_DATA = 'DB_DATA';
const PROP_FICHIERS_VUS = 'FICHIERS_VUS';
// v4.9.8 : Compteurs de boucle
const PROP_COMPTEUR_DECOUVERTE = 'COMPTEUR_DECOUVERTE';
const PROP_COMPTEUR_TRAITEMENT = 'COMPTEUR_TRAITEMENT';
const PROP_FICHIERS_TRAITES_CYCLE = 'FICHIERS_TRAITES_CYCLE';
const PROP_DATE_DEBUT_CYCLE = 'DATE_DEBUT_CYCLE';
// ----------------------------------


/**
 * ====================================================================
 * FONCTION UTILITAIRE : Gestion des dossiers (v4.6)
 * ====================================================================
 */
function getOrCreateFolderByPath(path) {
  let parts = path.split('/');
  let currentFolder = DriveApp.getRootFolder();
  for (let part of parts) {
    if (part) { 
      let folders = currentFolder.getFoldersByName(part);
      if (folders.hasNext()) {
        currentFolder = folders.next();
      } else {
        currentFolder = currentFolder.createFolder(part);
      }
    }
  }
  Logger.log(`Dossier de travail assuré : ${currentFolder.getName()}`);
  return currentFolder;
}

/**
 * ====================================================================
 * FONCTION 1 : À EXÉCUTER UNE SEULE FOIS MANUELLEMENT
 * ====================================================================
 */
function LANCER_PREMIERE_ANALYSE_UNIQUEMENT() {
  Logger.log("Lancement de la PREMIÈRE ANALYSE (v4.10.1)...");
  
  // 1. Nettoyer tout état précédent
  supprimerDeclencheursScript();
  PropertiesService.getScriptProperties().deleteAllProperties();
  
  // 2. Obtenir le dossier de travail
  const folder = getOrCreateFolderByPath(CHEMIN_DOSSIER);
  
  // 3. Nettoyer les anciens Sheets dans ce dossier
  function supprimerSheetParNom(nom, folder) {
    try {
      const fichiers = folder.getFilesByName(nom);
      while (fichiers.hasNext()) {
        fichiers.next().setTrashed(true);
      }
    } catch (e) {}
  }
  supprimerSheetParNom(NOM_SHEET_DB, folder);
  supprimerSheetParNom(NOM_SHEET_ACTION, folder);
  supprimerSheetParNom(NOM_SHEET_TEMP_TODO, folder);
  supprimerSheetParNom(NOM_SHEET_LOG, folder);
  supprimerSheetParNom(NOM_SHEET_STATS, folder); // NOUVEAU

  // 4. Créer les Google Sheets et les déplacer dans le dossier
  const sheetDB = SpreadsheetApp.create(NOM_SHEET_DB);
  const sheetId_DB = sheetDB.getId();
  DriveApp.getFileById(sheetId_DB).moveTo(folder);
  sheetDB.appendRow(["ID Fichier", "Nom", "URL", "Taille", "ModifiéLe (ISO)", "Dossier (Parent)", "Hash / Statut"]);
  sheetDB.getRange("A:A").setNumberFormat('@'); 
  sheetDB.getRange("G:G").setNumberFormat('@'); 
  Logger.log(`Base de données permanente créée : ${sheetId_DB}`);
  
  const sheetAction = SpreadsheetApp.create(NOM_SHEET_ACTION);
  DriveApp.getFileById(sheetAction.getId()).moveTo(folder);
  sheetAction.appendRow(["Analyse en cours..."]);
  Logger.log(`Sheet d'action créé : ${sheetAction.getUrl()}`);
  
  const sheetTempTodo = SpreadsheetApp.create(NOM_SHEET_TEMP_TODO);
  const sheetId_Todo = sheetTempTodo.getId();
  DriveApp.getFileById(sheetId_Todo).moveTo(folder);
  sheetTempTodo.appendRow(["Action", "ID", "Nom", "URL", "Taille", "ModifiedISO", "Dossier", "RowToUpdate"]);
  Logger.log(`Sheet temporaire de tâches créé : ${sheetId_Todo}`);

  const sheetLog = SpreadsheetApp.create(NOM_SHEET_LOG);
  const sheetId_Log = sheetLog.getId();
  DriveApp.getFileById(sheetId_Log).moveTo(folder);
  sheetLog.appendRow(["Horodatage", "État", "Message"]);
  Logger.log(`Fichier journal créé : ${sheetId_Log}`);
  
  // NOUVEAU : Créer le Sheet de Stats
  const sheetStats = SpreadsheetApp.create(NOM_SHEET_STATS);
  const sheetId_Stats = sheetStats.getId();
  DriveApp.getFileById(sheetId_Stats).moveTo(folder);
  sheetStats.insertSheet("Tableau de Bord");
  sheetStats.insertSheet("Historique").appendRow(["Date", "Fichiers Totaux (DB)", "Espace Perdu (Go)", "Fichiers Traités (Cycle)", "Durée Cycle (min)", "Lots Découverte", "Lots Traitement"]);
  sheetStats.insertSheet("Analyse des Formats").appendRow(["Extension", "Nombre de Fichiers", "Taille Totale (Octets)"]);
  sheetStats.insertSheet("Top 100 Fichiers (Taille)").appendRow(["Nom", "Dossier", "Taille (Mo)"]);
  sheetStats.deleteSheet(sheetStats.getSheetByName("Feuille 1"));
  Logger.log(`Fichier de Stats créé : ${sheetId_Stats}`);
  
  // 5. Configurer l'état initial
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(PROP_SHEET_ID_DB, sheetId_DB);
  properties.setProperty(PROP_SHEET_ID_TODO, sheetId_Todo); 
  properties.setProperty(PROP_SHEET_ID_LOG, sheetId_Log);
  properties.setProperty(PROP_SHEET_ID_STATS, sheetId_Stats); // NOUVEAU
  properties.setProperty(PROP_ETAT_SCRIPT, 'DECOUVERTE_INITIALE');
  properties.setProperty(PROP_CONTINUATION_TOKEN, DriveApp.getFiles().getContinuationToken());
  properties.setProperty(PROP_DB_DATA, JSON.stringify({}));
  properties.setProperty(PROP_COMPTEUR_DECOUVERTE, "0"); 
  properties.setProperty(PROP_COMPTEUR_TRAITEMENT, "0");
  properties.setProperty(PROP_FICHIERS_TRAITES_CYCLE, "0");
  properties.setProperty(PROP_DATE_DEBUT_CYCLE, new Date().toISOString());

  // 6. Créer le premier déclencheur
  creerProchainDeclencheur('traiterLotFichiers', 1, 'DECOUVERTE_INITIALE'); 
  
  logToFile('DECOUVERTE_INITIALE', `Lancement de la PREMIÈRE ANALYSE (v4.10.1). Fichiers dans ${CHEMIN_DOSSIER}`);
  
  MailApp.sendEmail(EMAIL_POUR_RAPPORT, 
                    "[Drive v4.10.1] Lancement de l'analyse initiale", 
                    "L'analyse initiale (v4.10.1) de tous vos fichiers a commencé.\n" +
                    `Ajout des stats et du réveil intelligent.`);
}

/**
 * ====================================================================
 * FONCTION UTILITAIRE : Écriture dans le fichier journal (v4.9.7)
 * ====================================================================
 */
function logToFile(etat, message) {
  // v4.9.7 : Écrit dans le Logger.log standard (visible dans les Exécutions)
  const logMessage = `[${etat}] ${message}`;
  Logger.log(logMessage);
  
  // Garde l'écriture dans le Sheet [LOG] comme backup en cas de crash
  try {
    const sheetId_Log = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID_LOG);
    if (sheetId_Log) {
      const sheetLog = SpreadsheetApp.openById(sheetId_Log).getSheets()[0];
      const horodatage = new Date().toISOString();
      sheetLog.appendRow([horodatage, etat, message]);
    }
  } catch (e) {
    Logger.log(`ERREUR CRITIQUE : Impossible d'écrire dans le fichier journal. ${e.message}`);
  }
}

/**
 * ====================================================================
 * FONCTION 2 : DÉCLENCHEUR NOCTURNE QUOTIDIEN
 * ====================================================================
 */
function lancerAnalyseQuotidienne() {
  logToFile("QUOTIDIEN", "Lancement de l'analyse quotidienne (incrémentielle v4.10.1)...");
  const properties = PropertiesService.getScriptProperties();
  const sheetId_DB = properties.getProperty(PROP_SHEET_ID_DB);
  const sheetId_Todo = properties.getProperty(PROP_SHEET_ID_TODO);

  if (!sheetId_DB || !sheetId_Todo) {
    logToFile("ERREUR", "Base de données non initialisée. Exécutez 'LANCER_PREMIERE_ANALYSE_UNIQUEMENT' d'abord.");
    return;
  }
  
  // v4.9.9 : Vider le TEMP (logique de "table rase" pour un nouveau jour)
  try {
    SpreadsheetApp.openById(sheetId_Todo).getSheets()[0].clearContents().appendRow(["Action", "ID", "Nom", "URL", "Taille", "ModifiedISO", "Dossier", "RowToUpdate"]);
  } catch (e) {
     logToFile("ERREUR", `Erreur lors du vidage du Sheet TEMP: ${e.message}`);
  }
  
  let dbData = {}; 
  try {
    const sheet = SpreadsheetApp.openById(sheetId_DB).getSheets()[0];
    const data = sheet.getDataRange().getValues();
    data.shift(); 
    
    data.forEach((row, index) => {
      const fileId = row[0];
      if (fileId) {
        dbData[fileId] = {
          row: index + 2, 
          modifiedISO: row[4],
          hash: row[6] // Colonne G
        };
      }
    });
    logToFile("QUOTIDIEN", `Base de données chargée. ${Object.keys(dbData).length} fichiers en référence.`);
  } catch (e) {
    logToFile("ERREUR", `ERREUR critique lors du chargement de la DB: ${e.message}`);
    return;
  }

  supprimerDeclencheursScript();
  properties.setProperty(PROP_ETAT_SCRIPT, 'DECOUVERTE');
  properties.setProperty(PROP_CONTINUATION_TOKEN, DriveApp.getFiles().getContinuationToken());
  properties.setProperty(PROP_DB_DATA, JSON.stringify(dbData));
  properties.setProperty(PROP_FICHIERS_VUS, JSON.stringify({})); 
  properties.setProperty(PROP_COMPTEUR_DECOUVERTE, "0"); // Reset compteurs
  properties.setProperty(PROP_COMPTEUR_TRAITEMENT, "0");
  properties.setProperty(PROP_FICHIERS_TRAITES_CYCLE, "0");
  properties.setProperty(PROP_DATE_DEBUT_CYCLE, new Date().toISOString());

  creerProchainDeclencheur('traiterLotFichiers', 1, 'DECOUVERTE');
  
  MailApp.sendEmail(EMAIL_POUR_RAPPORT, 
                    "[Drive v4.10.1] Lancement de l'analyse quotidienne", 
                    "L'analyse incrémentielle des fichiers nouveaux/modifiés a commencé.");
}


/**
 * ====================================================================
 * FONCTION 3 : LE CŒUR DU SCRIPT (Machine à états)
 * ====================================================================
 */
function traiterLotFichiers() {
  const properties = PropertiesService.getScriptProperties();
  const etat = properties.getProperty(PROP_ETAT_SCRIPT);

  try {
    switch (etat) {
      case 'DECOUVERTE_INITIALE':
        logiqueDeDecouverte(true); 
        break;
      case 'DECOUVERTE':
        logiqueDeDecouverte(false);
        break;
      case 'TRAITEMENT':
        logiqueDeTraitement();
        break;
      case 'NETTOYAGE':
        logiqueDeNettoyage();
        break;
      case 'RAPPORT':
        genererRapportFinal();
        break;
      default:
        logToFile("ERREUR", `État inconnu: ${etat}. Arrêt.`);
        supprimerDeclencheursScript();
    }
  } catch (err) {
    // Erreur "normale"
    logToFile("ERREUR FATALE", `ERREUR MAJEURE : ${err.message}. Arrêt du script. ${err.stack}`);
    MailApp.sendEmail(EMAIL_POUR_RAPPORT, "[Drive v4.10.1] ERREUR FATALE", `Une erreur a stoppé le script : ${err.message}\n${err.stack}`);
    supprimerDeclencheursScript();
  }
}

/**
 * ====================================================================
 * LOGIQUE ÉTAT 1 & 2 : DÉCOUVERTE (MODIFIÉ v4.10.0)
 * ====================================================================
 */
function logiqueDeDecouverte(estInitial) {
  const properties = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  
  // v4.9.8 : Incrémenter le compteur
  let compteur = parseInt(properties.getProperty(PROP_COMPTEUR_DECOUVERTE) || "0", 10) + 1;
  properties.setProperty(PROP_COMPTEUR_DECOUVERTE, compteur.toString());
  const etatActuel = estInitial ? 'DECOUVERTE_INITIALE' : 'DECOUVERTE';
  logToFile(etatActuel, `Exécution du lot de découverte #${compteur}...`);
  
  const token = properties.getProperty(PROP_CONTINUATION_TOKEN);
  const sheetId_Todo = properties.getProperty(PROP_SHEET_ID_TODO); 
  let dbData = JSON.parse(properties.getProperty(PROP_DB_DATA) || '{}');
  let fichiersVus = estInitial ? {} : JSON.parse(properties.getProperty(PROP_FICHIERS_VUS) || '{}');
  
  if (!token) {
    logToFile("ERREUR", "Token de découverte manquant. Arrêt.");
    return;
  }
  
  let iterator = DriveApp.continueFileIterator(token);
  const sheetTodo = SpreadsheetApp.openById(sheetId_Todo).getSheets()[0]; 
  let nouvellesTaches = []; // Batch pour l'écriture

  while (iterator.hasNext()) {
    const tempsEcoule = (new Date().getTime() - startTime) / 1000;
    if (tempsEcoule > TEMPS_MAX_EXECUTION_SECONDES) {
      logToFile(etatActuel, "Temps limite du lot de DÉCOUVERTE atteint. Sauvegarde...");
      
      if (nouvellesTaches.length > 0) {
        sheetTodo.getRange(sheetTodo.getLastRow() + 1, 1, nouvellesTaches.length, nouvellesTaches[0].length).setValues(nouvellesTaches);
      }
      
      properties.setProperty(PROP_CONTINUATION_TOKEN, iterator.getContinuationToken());
      if (!estInitial) properties.setProperty(PROP_FICHIERS_VUS, JSON.stringify(fichiersVus));
      
      // v4.10.0 : Utilise la PAUSE de 1 min (votre demande)
      creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, etatActuel);
      return; // Sortir
    }
    
    try {
      const fichier = iterator.next();
      const fileId = fichier.getId();
      const currentModifiedISO = fichier.getLastUpdated().toISOString();

      if (!estInitial) {
        fichiersVus[fileId] = 1; 
      }
      
      let dossier = "Racine"; 
      try {
        if (fichier.getParents().hasNext()) {
          dossier = fichier.getParents().next().getName();
        }
      } catch (e) {}

      let infoFichier = {
        id: fileId,
        nom: fichier.getName(),
        url: fichier.getUrl(),
        taille: fichier.getSize(),
        modifiedISO: currentModifiedISO,
        mime: fichier.getMimeType(),
        dossier: dossier 
      };
      
      let action = null; 
      let rowToUpdate = null;
      
      if (estInitial) {
        action = 'NEW';
      } else {
        const storedData = dbData[fileId];
        if (!storedData) {
          action = 'NEW';
        } else if (storedData.modifiedISO !== currentModifiedISO) {
          action = 'MODIFIED';
          rowToUpdate = storedData.row;
        }
      }
      
      if (action) {
        nouvellesTaches.push([
          action, infoFichier.id, infoFichier.nom, infoFichier.url, 
          infoFichier.taille, infoFichier.modifiedISO, infoFichier.dossier, rowToUpdate
        ]);
      }
      
      if (nouvellesTaches.length >= 100) {
        sheetTodo.getRange(sheetTodo.getLastRow() + 1, 1, nouvellesTaches.length, nouvellesTaches[0].length).setValues(nouvellesTaches);
        nouvellesTaches = []; // Vider le batch
      }
      
    } catch (e) {
      logToFile(etatActuel, `Erreur sur 1 fichier pendant découverte: ${e.message}. Ignoré.`);
    }
  } 
  
  if (nouvellesTaches.length > 0) {
    sheetTodo.getRange(sheetTodo.getLastRow() + 1, 1, nouvellesTaches.length, nouvellesTaches[0].length).setValues(nouvellesTaches);
  }
  
  logToFile(etatActuel, "Phase de DÉCOUVERTE terminée.");
  
  if (estInitial) {
    properties.setProperty(PROP_ETAT_SCRIPT, 'TRAITEMENT');
  } else {
    properties.setProperty(PROP_ETAT_SCRIPT, 'NETTOYAGE');
    properties.setProperty(PROP_FICHIERS_VUS, JSON.stringify(fichiersVus));
  }
  
  creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'TRAITEMENT');
}


/**
 * ====================================================================
 * LOGIQUE ÉTAT 3 : TRAITEMENT (MODIFIÉ v4.10.0)
 * ====================================================================
 */
function logiqueDeTraitement() {
  const properties = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  let tempsEcoule = 0;
  
  // v4.9.8 : Incrémenter le compteur
  let compteur = parseInt(properties.getProperty(PROP_COMPTEUR_TRAITEMENT) || "0", 10) + 1;
  properties.setProperty(PROP_COMPTEUR_TRAITEMENT, compteur.toString());
  let fichiersTraitesCeLot = 0;

  const sheetId_DB = properties.getProperty(PROP_SHEET_ID_DB);
  const sheetId_Todo = properties.getProperty(PROP_SHEET_ID_TODO);
  
  const sheetDB = SpreadsheetApp.openById(sheetId_DB).getSheets()[0];
  const sheetTodo = SpreadsheetApp.openById(sheetId_Todo).getSheets()[0];
  
  let taches = [];
  try {
    const data = sheetTodo.getDataRange().getValues();
    data.shift(); // Enlever l'en-tête
    if (data.length === 0) {
      logToFile("TRAITEMENT", "Phase de TRAITEMENT terminée (rien à faire). Passage au RAPPORT.");
      properties.setProperty(PROP_ETAT_SCRIPT, 'RAPPORT');
      creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'RAPPORT');
      return;
    }
    
    const nbTaches = Math.min(data.length, BATCH_SIZE_TRAITEMENT);
    taches = data.slice(0, nbTaches);
    
  } catch (e) {
    logToFile("ERREUR", `Erreur lecture Sheet TEMP: ${e.message}`);
    return;
  }
  
  logToFile("TRAITEMENT", `Début du lot de TRAITEMENT #${compteur}. Traitement de ${taches.length} fichiers.`);
  
  let lignesNouvellesDB = [];
  let i = 0;
  
  while (i < taches.length) {
    tempsEcoule = (new Date().getTime() - startTime) / 1000;
    if (tempsEcoule > TEMPS_MAX_EXECUTION_SECONDES) {
      logToFile("TRAITEMENT", "Temps limite du lot de TRAITEMENT atteint. Sauvegarde...");
      break; 
    }
    
    const item = taches[i];
    const action = item[0];
    const info = {
      id: item[1],
      nom: item[2],
      url: item[3],
      taille: parseFloat(item[4]),
      modifiedISO: item[5],
      dossier: item[6],
    };
    const rowToUpdate = item[7];
    
    let hashOuStatut = '';
    const logPrefix = `Traitement: ID ${info.id} | Nom: ${info.nom} | Taille: ${info.taille}`;

    try {
      // VÉRIFICATION v4.9.2 : TAILLE UNIQUEMENT
      if (info.taille > MAX_FILE_SIZE_BYTES) { 
          hashOuStatut = 'IGNORÉ - Fichier trop volumineux';
          logToFile("TRAITEMENT", `${logPrefix} | Statut: ${hashOuStatut} (Limite 45Mo)`);
      }
      // Si non exclu, on continue
      else {
        logToFile("TRAITEMENT", `${logPrefix} | Étape 1: getFileById...`);
        const fichier = DriveApp.getFileById(info.id);
        
        logToFile("TRAITEMENT", `${logPrefix} | Étape 2: getMimeType...`);
        const mimeType = fichier.getMimeType();

        if (info.taille === 0 || mimeType === MimeType.SHORTCUT || mimeType.includes('google-apps')) {
          hashOuStatut = 'IGNORÉ - Type Google/Vide';
          logToFile("TRAITEMENT", `${logPrefix} | Statut: ${hashOuStatut}`);
        } else {
          logToFile("TRAITEMENT", `${logPrefix} | Étape 3: getBlob (opération risquée)...`);
          const blob = fichier.getBlob(); 
          
          logToFile("TRAITEMENT", `${logPrefix} | Étape 4: computeDigest (calcul)...`);
          const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, blob.getBytes());
          
          hashOuStatut = hashBytes.map(function(byte) {
            var hex = ((byte + 256) % 256).toString(16);
            return hex.length === 1 ? '0' + hex : hex;
          }).join('');
          logToFile("TRAITEMENT", `${logPrefix} | Étape 5: Hash calculé.`);
        }
      }
    } catch (e) {
      logToFile("ERREUR", `${logPrefix} | ERREUR NORMALE: ${e.message}`);
      hashOuStatut = `ERREUR - ${e.message}`;
    }
    
    const rowData = [info.id, info.nom, info.url, info.taille, info.modifiedISO, info.dossier, hashOuStatut];
    
    if (action === 'NEW') {
      lignesNouvellesDB.push(rowData);
    } else if (action === 'MODIFIED') {
      const range = sheetDB.getRange(rowToUpdate, 1, 1, rowData.length);
      range.setValues([rowData]);
    }
    i++;
    fichiersTraitesCeLot++;
  } // Fin du While

  // Mise à jour du compteur global
  let fichiersTraitesTotal = parseInt(properties.getProperty(PROP_FICHIERS_TRAITES_CYCLE) || "0", 10) + fichiersTraitesCeLot;
  properties.setProperty(PROP_FICHIERS_TRAITES_CYCLE, fichiersTraitesTotal.toString());

  if (lignesNouvellesDB.length > 0) {
    logToFile("TRAITEMENT", `Ajout de ${lignesNouvellesDB.length} nouvelles lignes à la DB...`);
    sheetDB.getRange(sheetDB.getLastRow() + 1, 1, lignesNouvellesDB.length, lignesNouvellesDB[0].length).setValues(lignesNouvellesDB);
  }
  
  if (i > 0) {
    logToFile("TRAITEMENT", `Suppression de ${i} tâches traitées du Sheet [TEMP]...`);
    sheetTodo.deleteRows(2, i); // Supprime les lignes 2 à i+1 (car 1-indexé + en-tête)
  }

  // v4.10.0 : Logique de pause (toujours 1 minute)
  if (tempsEcoule > TEMPS_MAX_EXECUTION_SECONDES) {
    logToFile("TRAITEMENT", "Lot terminé (timeout). Relance avec pause de 1 min...");
  } else {
    logToFile("TRAITEMENT", "Lot terminé (rapide). Relance avec relais de 1 min...");
  }
  creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'TRAITEMENT');
}


/**
 * ====================================================================
 * LOGIQUE ÉTAT 4 : NETTOYAGE
 * ====================================================================
 */
function logiqueDeNettoyage() {
  logToFile("NETTOYAGE", "Début de la phase de NETTOYAGE...");
  const properties = PropertiesService.getScriptProperties();
  const sheetId_DB = properties.getProperty(PROP_SHEET_ID_DB);
  
  const dbData = JSON.parse(properties.getProperty(PROP_DB_DATA) || '{}');
  const fichiersVus = JSON.parse(properties.getProperty(PROP_FICHIERS_VUS) || '{}');
  
  const sheetDB = SpreadsheetApp.openById(sheetId_DB).getSheets()[0];
  let modifications = []; 
  
  for (const fileId in dbData) {
    if (!fichiersVus[fileId]) {
      const info = dbData[fileId];
      if (info.hash !== 'SUPPRIMÉ') { 
        modifications.push({
          range: sheetDB.getRange(info.row, 7), // Colonne G (Hash/Statut)
          value: 'SUPPRIMÉ'
        });
      }
    }
  }
  
  logToFile("NETTOYAGE", `Marquage de ${modifications.length} fichiers comme 'SUPPRIMÉ'.`);
  
  for (const modif of modifications) {
    try {
      modif.range.setValue(modif.value);
    } catch(e) {
      logToFile("ERREUR", `Erreur lors du marquage 'SUPPRIMÉ': ${e.message}`);
    }
  }

  logToFile("NETTOYAGE", "Phase de NETTOYAGE terminée.");
  
  properties.setProperty(PROP_ETAT_SCRIPT, 'TRAITEMENT');
  creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'TRAITEMENT');
}


/**
 * ====================================================================
 * LOGIQUE ÉTAT 5 : RAPPORT (MODIFIÉ v4.10.0)
 * ====================================================================
 */
function genererRapportFinal() {
  logToFile("RAPPORT", "Génération du rapport final (v4.10.1)...");
  const properties = PropertiesService.getScriptProperties();
  const sheetId_DB = properties.getProperty(PROP_SHEET_ID_DB);
  const sheetId_Todo = properties.getProperty(PROP_SHEET_ID_TODO); 
  const sheetId_Stats = properties.getProperty(PROP_SHEET_ID_STATS);
  const folder = getOrCreateFolderByPath(CHEMIN_DOSSIER);

  let dataDB;
  try {
    const sheet = SpreadsheetApp.openById(sheetId_DB).getSheets()[0];
    dataDB = sheet.getDataRange().getValues();
    dataDB.shift(); // Enlève l'en-tête
  } catch (e) {
    logToFile("ERREUR", `Impossible de lire la DB ${sheetId_DB}. ${e.message}`);
    MailApp.sendEmail(EMAIL_POUR_RAPPORT, "[Drive v4.10.1] ERREUR Rapport", `Impossible de lire la base de donnees pour generer le rapport final.`);
    return;
  }

  // --- Initialisation des stats ---
  const mapHashes = {};
  const fichiersIgnores = [];
  const mapFormats = {}; // v4.9.8
  const top100Fichiers = []; // v4.9.8
  let totalFichiersDB = 0;
  let espaceTotal = 0;
  
  for (const row of dataDB) {
    totalFichiersDB++;
    // row = [ID, Nom, URL, Taille, ModifiéLe, Dossier, Hash/Statut]
    const hashOrStatus = row[6];
    const dossier = row[5];
    const nom = row[1];
    const taille = parseFloat(row[3]) || 0;
    espaceTotal += taille;
    
    // v4.9.8 : Analyse des formats
    const extension = (nom.includes('.')) ? nom.substring(nom.lastIndexOf('.')).toLowerCase() : "[Aucune extension]";
    if (!mapFormats[extension]) {
      mapFormats[extension] = { count: 0, size: 0 };
    }
    mapFormats[extension].count++;
    mapFormats[extension].size += taille;

    // v4.9.8 : Top 100
    top100Fichiers.push({ nom: nom, dossier: dossier, taille: taille });
    
    // Logique de doublons
    if (hashOrStatus.startsWith("IGNORÉ") || hashOrStatus.startsWith("ERREUR") || hashOrStatus.startsWith("SUPPRIMÉ")) {
      fichiersIgnores.push({
        nom: nom,
        taille: taille,
        raison: hashOrStatus,
        dossier: dossier
      });
    } else if (hashOrStatus) {
      const infoFichier = { id: row[0], nom: row[1], url: row[2], taille: taille, hash: hashOrStatus, dossier: dossier };
      if (!mapHashes[infoFichier.hash]) {
        mapHashes[infoFichier.hash] = [];
      }
      mapHashes[infoFichier.hash].push(infoFichier);
    }
  }

  // --- Préparation du Sheet d'ACTION ---
  let sheetAction;
  let sheetActionUrl;
  try {
    const files = folder.getFilesByName(NOM_SHEET_ACTION); // Chercher dans le dossier
    if (files.hasNext()) {
      sheetAction = SpreadsheetApp.open(files.next());
      sheetActionUrl = sheetAction.getUrl();
    } else {
      sheetAction = SpreadsheetApp.create(NOM_SHEET_ACTION);
      sheetActionUrl = sheetAction.getUrl();
      DriveApp.getFileById(sheetAction.getId()).moveTo(folder);
    }
    
    const sheet = sheetAction.getSheets()[0];
    sheet.clear(); 
    sheet.appendRow(["Nom du Fichier", "Dossier (Parent)", "Taille (octets)", "URL (Lien)", "ID Fichier", "Hash (Groupe)", "ACTION"]);
    sheet.setFrozenRows(1);
    
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['GARDER'], true).build();
    sheet.getRange('G2:G').setDataValidation(rule);
    
  } catch (e) {
    logToFile("ERREUR", `Impossible de créer/modifier le Sheet d'ACTION: ${e.message}`);
    MailApp.sendEmail(EMAIL_POUR_RAPPORT, "[Drive v4.10.1] ERREUR Rapport", `Impossible de générer le Sheet d'action : ${e.message}`);
    return;
  }

  // --- Remplissage du Sheet d'ACTION et calcul des stats ---
  let totalDoublons = 0;
  let espacePerdu = 0;
  let lignesAction = [];

  for (const hash in mapHashes) {
    const fichiersIdentiques = mapHashes[hash];
    if (fichiersIdentiques.length > 1) { 
      totalDoublons += fichiersIdentiques.length;
      espacePerdu += fichiersIdentiques[0].taille * (fichiersIdentiques.length - 1);
      
      for (const info of fichiersIdentiques) {
        lignesAction.push([
          info.nom, info.dossier, info.taille,
          info.url, info.id, info.hash, "" 
        ]);
      }
      lignesAction.push(["", "", "", "", "", "", ""]); 
    }
  }
  
  if (lignesAction.length > 0) {
    sheetAction.getSheets()[0].getRange(2, 1, lignesAction.length, lignesAction[0].length).setValues(lignesAction);
    logToFile("RAPPORT", `Sheet d'action mis à jour avec ${totalDoublons} fichiers.`);
  } else {
    sheetAction.getSheets()[0].getRange(2, 1).setValue("Aucun doublon trouvé.");
  }


  // --- CSV des IGNORÉS/SUPPRIMÉS ---
  let csvContenuIgnore = '"Nom du Fichier","Dossier (Parent)","Taille (octets)","Raison"\n';
  for (const info of fichiersIgnores) {
    csvContenuIgnore += `"${info.nom.replace(/"/g, '""')}","${info.dossier.replace(/"/g, '""')}","${info.taille}","${info.raison}"\n`;
  }
  const DATE_AUJOURDHUI = new Date().toISOString().slice(0, 10);
  const nomFichierIgnore = `Rapport_Fichiers_Ignorés_${DATE_AUJOURDHUI}.csv`;
  const fichierCsvIgnore = folder.createFile(nomFichierIgnore, csvContenuIgnore, MimeType.CSV); // Créer dans le dossier
  logToFile("RAPPORT", `Rapport des ignorés créé : ${fichierCsvIgnore.getName()}`);

  // --- v4.9.8 : Mise à jour du [STATS] Tableau de Bord ---
  try {
    const statsSheet = SpreadsheetApp.openById(sheetId_Stats);
    
    // 1. Tableau de Bord
    const tdbSheet = statsSheet.getSheetByName("Tableau de Bord");
    tdbSheet.clear();
    const dateFinCycle = new Date();
    const dateDebutCycle = new Date(properties.getProperty(PROP_DATE_DEBUT_CYCLE));
    const dureeCycleMs = dateFinCycle.getTime() - dateDebutCycle.getTime();
    const dureeCycleMin = (dureeCycleMs / 1000 / 60).toFixed(2);
    const espacePerduGo = (espacePerdu / 1024 / 1024 / 1024).toFixed(2);
    const espaceTotalGo = (espaceTotal / 1024 / 1024 / 1024).toFixed(2);
    
    tdbSheet.appendRow(["Statistique", "Valeur"]);
    tdbSheet.appendRow(["Dernière Analyse", dateFinCycle.toLocaleString("fr-BE")]);
    tdbSheet.appendRow(["Durée Totale du Cycle (min)", dureeCycleMin]);
    tdbSheet.appendRow(["Fichiers Traités (ce cycle)", properties.getProperty(PROP_FICHIERS_TRAITES_CYCLE) || "0"]);
    tdbSheet.appendRow(["Lots de Découverte (ce cycle)", properties.getProperty(PROP_COMPTEUR_DECOUVERTE) || "0"]);
    tdbSheet.appendRow(["Lots de Traitement (ce cycle)", properties.getProperty(PROP_COMPTEUR_TRAITEMENT) || "0"]);
    tdbSheet.appendRow(["---", "---"]);
    tdbSheet.appendRow(["Fichiers Totaux (dans la DB)", totalFichiersDB]);
    tdbSheet.appendRow(["Espace Disque Total (Go)", espaceTotalGo]);
    tdbSheet.appendRow(["Fichiers en Doublon", totalDoublons]);
    tdbSheet.appendRow(["Espace Perdu (Go)", espacePerduGo]);
    tdbSheet.appendRow(["Fichiers Ignorés/Erreurs", fichiersIgnores.length]);
    tdbSheet.getRange("A1:B1").setFontWeight("bold");
    tdbSheet.getRange("A:A").setFontWeight("bold");
    
    // 2. Historique
    const histSheet = statsSheet.getSheetByName("Historique");
    histSheet.appendRow([
      dateFinCycle.toISOString().slice(0, 10),
      totalFichiersDB,
      espacePerduGo,
      properties.getProperty(PROP_FICHIERS_TRAITES_CYCLE) || "0",
      dureeCycleMin,
      properties.getProperty(PROP_COMPTEUR_DECOUVERTE) || "0",
      properties.getProperty(PROP_COMPTEUR_TRAITEMENT) || "0"
    ]);

    // 3. Analyse des Formats
    const formatSheet = statsSheet.getSheetByName("Analyse des Formats");
    formatSheet.clear();
    formatSheet.appendRow(["Extension", "Nombre de Fichiers", "Taille Totale (Octets)"]);
    let formatData = [];
    for (const ext in mapFormats) {
      formatData.push([ext, mapFormats[ext].count, mapFormats[ext].size]);
    }
    formatData.sort((a, b) => b[1] - a[1]); // Trier par nombre
    formatSheet.getRange(2, 1, formatData.length, 3).setValues(formatData);
    formatSheet.getRange("A1:C1").setFontWeight("bold");
    
    // 4. Top 100 Fichiers
    const topSheet = statsSheet.getSheetByName("Top 100 Fichiers (Taille)");
    topSheet.clear();
    topSheet.appendRow(["Nom", "Dossier", "Taille (Mo)"]);
    top100Fichiers.sort((a, b) => b.taille - a.taille); // Trier par taille
    let top100Data = [];
    for (let k = 0; k < Math.min(top100Fichiers.length, 100); k++) {
      const f = top100Fichiers[k];
      top100Data.push([f.nom, f.dossier, (f.taille / 1024 / 1024).toFixed(2)]);
    }
    topSheet.getRange(2, 1, top100Data.length, 3).setValues(top100Data);
    topSheet.getRange("A1:C1").setFontWeight("bold");

    logToFile("RAPPORT", `Tableau de bord [STATS] mis à jour.`);
    
  } catch (e) {
    logToFile("ERREUR", `Impossible de mettre à jour le Sheet [STATS]: ${e.message}`);
  }

  // --- Email final ---
  let corpsEmail = `Analyse des doublons terminée (v4.10.1 - Incrémentiel).
    
Fichiers dans la base de données: ${totalFichiersDB}
Fichiers en doublon trouvés: ${totalDoublons}
Espace disque potentiellement perdu: ${espacePerdu} octets
Fichiers ignorés/supprimés: ${fichiersIgnores.length}

Un tableau de bord complet des statistiques a été mis à jour ici :
${SpreadsheetApp.openById(sheetId_Stats).getUrl()}
`;
  
  if(totalDoublons > 0) {
    corpsEmail += `\n\nPOUR AGIR :
Ouvrez votre panneau de contrôle pour sélectionner les fichiers à garder :
${sheetActionUrl}`;
  } else {
    corpsEmail += "\n\nBonne nouvelle ! Aucun doublon de contenu n'a été trouvé.";
  }
  
  corpsEmail += `\n\nRapport CSV des **fichiers ignorés/supprimés**: ${fichierCsvIgnore.getUrl()}`;
  corpsEmail += `\n\nTous les fichiers de travail se trouvent dans le dossier : ${folder.getUrl()}`;
  
  MailApp.sendEmail(EMAIL_POUR_RAPPORT, "[Drive v4.10.1] Rapport d'analyse des doublons terminé", corpsEmail);
  logToFile("RAPPORT", "Rapport d'action créé et email envoyé.");
  
  // Nettoyer le Sheet [TEMP] à la fin
  try {
    DriveApp.getFileById(sheetId_Todo).setTrashed(true);
    logToFile("RAPPORT", "Sheet temporaire de tâches supprimé.");
  } catch (e) {
    logToFile("RAPPORT", `Avertissement: Impossible de supprimer le Sheet TEMP: ${e.message}`);
  }
  
  supprimerDeclencheursScript();
  properties.setProperty(PROP_ETAT_SCRIPT, 'IDLE'); 
  logToFile("RAPPORT", "Script terminé. En attente de la prochaine exécution nocturne.");
}


// --- FONCTIONS UTILITAIRES ---

/**
 * Crée le prochain déclencheur pour la fonction donnée. (v4.9.8)
 */
function creerProchainDeclencheur(fonction, minutes, etat) {
  supprimerDeclencheursScript();
  ScriptApp.newTrigger(fonction)
      .timeBased()
      .after(minutes * 60 * 1000)
      .create();
  // v4.9.8 : Log amélioré
  Logger.log(`Prochain lot (${fonction} - État: ${etat || 'INCONNU'}) programmé dans ${minutes} minutes.`);
}

/**
 * Supprime tous les déclencheurs de ce script.
 */
function supprimerDeclencheursScript() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'traiterLotFichiers' || 
        funcName === 'lancerAnalyseQuotidienne' ||
        funcName === 'lanceurNocturneIntelligent') { // v4.9.9 : Ajout du nouveau
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * ====================================================================
 * FONCTION 5 : DÉCLENCHEUR NOCTURNE "INTELLIGENT" (v4.9.9)
 * ====================================================================
 * C'est le "concierge" (janitor) qui s'exécute à 2h du matin.
 * Il vérifie si le travail de la veille a été interrompu par le quota.
 */
function lanceurNocturneIntelligent() {
  logToFile("SYSTEM", "Réveil intelligent (02:00) activé.");
  const etat = PropertiesService.getScriptProperties().getProperty(PROP_ETAT_SCRIPT);
  
  if (etat === 'IDLE' || !etat) {
    // Le travail de la veille était terminé, on lance un nouveau scan
    logToFile("SYSTEM", "État = IDLE. Lancement d'une nouvelle analyse quotidienne.");
    lancerAnalyseQuotidienne();
  } else {
    // Le travail était en cours (ex: gelé par quota de 6h)
    logToFile("SYSTEM", `État = ${etat}. Reprise du travail interrompu.`);
    // On relance simplement le "coeur" du script pour qu'il continue
    creerProchainDeclencheur('traiterLotFichiers', 1, etat);
  }
}

/**
 * ====================================================================
 * FONCTION 6 : (Supprimée v4.10.1)
 * ====================================================================
 * L'ancienne fonction 'INSTALLER_DECLENCHEUR_NOCTURNE' a été supprimée.
 * Veuillez créer le déclencheur manuellement via l'interface ⏰ :
 * 1. Déclencheurs > + Ajouter un déclencheur
 * 2. Fonction à exécuter : lanceurNocturneIntelligent
 * 3. Source de l'événement : Temporel
 * 4. Type : Déclencheur basé sur les jours
 * 5. Heure : 2h - 3h du matin
 * ====================================================================
 */
