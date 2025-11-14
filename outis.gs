/*
 * ====================================================================
 * BOITE À OUTILS & ARCHIVES (v4.14.1)
 * ====================================================================
 *
 * Ce fichier contient les fonctions d'initialisation, de réparation
 * et de maintenance.
 *
 * POUR UTILISER UNE FONCTION :
 * 1. Sélectionnez la fonction voulue dans l'éditeur Apps Script.
 * 2. Cliquez sur "Exécuter".
 *
 * v4.14.1:
 * - Nettoyage des espaces insécables.
 * - Activation des fonctions (retrait des commentaires bloquants).
 * - Formatage des commentaires en-têtes de section
 * - Archivage de REINITIALISER_ETAT_MANUELLEMENT.
 * - CORRECTION : Ajout de l'onglet "Dossiers Vides" manquant
 * dans RESTAURER_PROPRIETES_RAPPORT et LANCER_PREMIERE_ANALYSE.
 * ====================================================================
 */

/* --- Variables de configuration (Dupliquées pour autonomie) --- */
const EMAIL_POUR_RAPPORT_ARCHIVE = "g@alsteens.net";
const CHEMIN_DOSSIER_ARCHIVE = "script/doublondrive";
const NOM_SHEET_DB_ARCHIVE = `[DB] Hashes Fichiers Drive`;
const NOM_SHEET_ACTION_ARCHIVE = `[ACTION] Suppression Doublons`;
const NOM_SHEET_TEMP_TODO_ARCHIVE = `[TEMP] Fichiers à Traiter`; 
const NOM_SHEET_LOG_ARCHIVE = `[LOG] Journal Analyse Doublons`;
const NOM_SHEET_STATS_ARCHIVE = `[STATS] Tableau de Bord`; 
const NOM_DOSSIER_ORPHELINS_ARCHIVE = "FICHIER ORPHELIN";

const MINUTES_ENTRE_LOTS_ARCHIVE = 1; 
const TEMPS_MAX_EXECUTION_SECONDES_ARCHIVE = 270; 

const MAX_FILE_SIZE_BYTES_ARCHIVE = 47185920; /* 45 Mo */
const BATCH_SIZE_TRAITEMENT_ARCHIVE = 50; 


/* --- Outil 1 : Créer le Log --- */
function CREER_FICHIER_LOG_MANUELLEMENT() {
  const folder = _archive_getOrCreateFolderByPath(CHEMIN_DOSSIER_ARCHIVE);
  let sheetId_Log = null;

  try {
    const files = folder.getFilesByName(NOM_SHEET_LOG_ARCHIVE);
    if (files.hasNext()) {
        const file = files.next();
        sheetId_Log = file.getId();
        Logger.log(`Fichier log existant trouvé: ${NOM_SHEET_LOG_ARCHIVE}. ID: ${sheetId_Log}`);
    } else {
        const sheetLog = SpreadsheetApp.create(NOM_SHEET_LOG_ARCHIVE);
        sheetId_Log = sheetLog.getId();
        DriveApp.getFileById(sheetId_Log).moveTo(folder);
        /* En-tête structuré v4.12 */
        sheetLog.appendRow(["Horodatage", "État", "ID Fichier", "Nom Fichier", "Message"]); 
        Logger.log(`Fichier log '${NOM_SHEET_LOG_ARCHIVE}' créé avec succès. ID: ${sheetId_Log}`);
    }
    
    PropertiesService.getScriptProperties().setProperty('SHEET_ID_LOG', sheetId_Log);
    
    MailApp.sendEmail(
      Session.getActiveUser().getEmail(), 
      `[Script Action Manuelle] Log prêt`, 
      `Le fichier log '${NOM_SHEET_LOG_ARCHIVE}' est prêt à être utilisé par le script.`
    );
  } catch (e) {
    Logger.log(`ERREUR CRITIQUE: Impossible de créer/trouver le fichier log : ${e.message}`);
  }
}


/* --- Outil 2 : Restaurer les propriétés --- */
function RESTAURER_PROPRIETES_RAPPORT() {
  Logger.log("Restauration des propriétés de rapport...");
  const properties = PropertiesService.getScriptProperties();
  const folder = _archive_getOrCreateFolderByPath(CHEMIN_DOSSIER_ARCHIVE);
  
  /* Fonction interne pour trouver ou créer le fichier */
  function findOrCreateSheet(nomFichier, setupFunction) {
    const files = folder.getFilesByName(nomFichier);
    let sheetId = null;
    let sheet;
    
    if (files.hasNext()) {
      const file = files.next();
      sheetId = file.getId();
      Logger.log(`Fichier ${nomFichier} trouvé. ID: ${sheetId}`);
      sheet = SpreadsheetApp.openById(sheetId);
    } else {
      sheet = SpreadsheetApp.create(nomFichier);
      sheetId = sheet.getId();
      DriveApp.getFileById(sheetId).moveTo(folder);
      if (setupFunction) {
        setupFunction(sheet); /* Applique les en-têtes, etc. */
      }
      Logger.log(`Fichier ${nomFichier} non trouvé. Création...`);
    }
    return sheetId;
  }

  /* Configuration pour le [STATS] */
  const setupStats = (sheet) => {
    sheet.insertSheet("Tableau de Bord");
    sheet.insertSheet("Historique").appendRow(["Date", "Fichiers Totaux (DB)", "Espace Perdu (Go)", "Fichiers Traités (Cycle)", "Durée Cycle (min)", "Lots Découverte", "Lots Traitement"]);
    sheet.insertSheet("Analyse des Formats").appendRow(["Extension", "Description", "Nombre", "Taille (Octets)"]);
    sheet.insertSheet("Top 100 Fichiers (Taille)").appendRow(["Nom", "Chemin Complet", "Taille (Mo)", "URL", "Date", "Heure"]);
    sheet.insertSheet("Fichiers Ignorés").appendRow(["ID", "Nom", "Chemin/Dossier", "Taille", "Raison"]);
    try { sheet.deleteSheet(sheet.getSheetByName("Feuille 1")); } catch(e) {}
  };
  
  /* Configuration pour le [ACTION] (Corrigé V4.14.1) */
  const setupAction = (sheet) => {
      sheet.insertSheet("Doublons").appendRow(["EFFACER", "Nom", "Chemin Complet", "Taille", "Date", "Heure", "URL", "ID", "Hash"]);
      sheet.insertSheet("Fichiers 0 ko").appendRow(["EFFACER", "Nom", "Chemin Complet", "Taille", "URL", "ID", "Date", "Heure"]);
      /* CORRECTION V4.14.1 : Ajout de l'onglet manquant "Dossiers Vides" */
      sheet.insertSheet("Dossiers Vides").appendRow(["Nom", "Chemin Complet", "URL", "ID"]);
      try { sheet.deleteSheet(sheet.getSheetByName("Feuille 1")); } catch(e) {}
  };
  
  /* Retrouver ou créer les fichiers critiques */
  const sheetId_Stats = findOrCreateSheet(NOM_SHEET_STATS_ARCHIVE, setupStats);
  const sheetId_Action = findOrCreateSheet(NOM_SHEET_ACTION_ARCHIVE, setupAction);
  const sheetId_DB = findOrCreateSheet(NOM_SHEET_DB_ARCHIVE, (sheet) => {
    sheet.appendRow(["ID", "Nom", "URL", "Taille", "ISO", "Dossier", "Hash", "Chemin", "Date", "Heure"]);
  });
  const sheetId_Log = findOrCreateSheet(NOM_SHEET_LOG_ARCHIVE, (sheet) => {
    sheet.appendRow(["Horodatage", "État", "ID Fichier", "Nom Fichier", "Message"]);
  });

  /* Sauvegarder les IDs (utilise les clés du script principal) */
  properties.setProperty('SHEET_ID_STATS', sheetId_Stats);
  properties.setProperty('SHEET_ID_ACTION', sheetId_Action);
  properties.setProperty('SHEET_ID_DB', sheetId_DB);
  properties.setProperty('SHEET_ID_LOG', sheetId_Log);

  Logger.log("Restauration des propriétés terminée.");
  try { 
    SpreadsheetApp.getUi().alert(`Restauration terminée. Les IDs des fichiers [STATS], [LOG], [ACTION] et [DB] ont été vérifiés et sauvegardés.`);
  } catch(e) {}
}


/* --- Outil 3 : Reset Usine (DANGEREUX) --- */
function LANCER_PREMIERE_ANALYSE_UNIQUEMENT() {
  Logger.log("Lancement de la PREMIÈRE ANALYSE (Reset usine)...");
  
  /* 1. Nettoyer tout état précédent */
  _archive_supprimerDeclencheursScript();
  PropertiesService.getScriptProperties().deleteAllProperties();
  
  /* 2. Obtenir le dossier de travail */
  const folder = _archive_getOrCreateFolderByPath(CHEMIN_DOSSIER_ARCHIVE);
  
  /* 3. Nettoyer les anciens Sheets */
  _archive_supprimerSheetParNom(NOM_SHEET_DB_ARCHIVE, folder);
  _archive_supprimerSheetParNom(NOM_SHEET_ACTION_ARCHIVE, folder);
  _archive_supprimerSheetParNom(NOM_SHEET_TEMP_TODO_ARCHIVE, folder);
  _archive_supprimerSheetParNom(NOM_SHEET_LOG_ARCHIVE, folder);
  _archive_supprimerSheetParNom(NOM_SHEET_STATS_ARCHIVE, folder);

  /* 4. Créer les Google Sheets */
  const sheetDB = SpreadsheetApp.create(NOM_SHEET_DB_ARCHIVE);
  const sheetId_DB = sheetDB.getId();
  DriveApp.getFileById(sheetId_DB).moveTo(folder);
  sheetDB.appendRow(["ID", "Nom", "URL", "Taille", "ISO", "Dossier", "Hash", "Chemin", "Date", "Heure"]);
  sheetDB.getRange("A:A").setNumberFormat('@'); 
  sheetDB.getRange("G:G").setNumberFormat('@'); 
  Logger.log(`Base de données permanente créée : ${sheetId_DB}`);
  
  const sheetAction = SpreadsheetApp.create(NOM_SHEET_ACTION_ARCHIVE);
  DriveApp.getFileById(sheetAction.getId()).moveTo(folder);
  sheetAction.insertSheet("Doublons").appendRow(["EFFACER", "Nom", "Chemin Complet", "Taille", "Date", "Heure", "URL", "ID", "Hash"]);
  sheetAction.insertSheet("Fichiers 0 ko").appendRow(["EFFACER", "Nom", "Chemin Complet", "Taille", "URL", "ID", "Date", "Heure"]);
  /* CORRECTION V4.14.1 : Ajout de l'onglet manquant "Dossiers Vides" */
  sheetAction.insertSheet("Dossiers Vides").appendRow(["Nom", "Chemin Complet", "URL", "ID"]);
  try { sheetAction.deleteSheet(sheetAction.getSheetByName("Feuille 1")); } catch(e) {}
  Logger.log(`Sheet d'action créé : ${sheetAction.getUrl()}`);
  
  const sheetTempTodo = SpreadsheetApp.create(NOM_SHEET_TEMP_TODO_ARCHIVE);
  const sheetId_Todo = sheetTempTodo.getId();
  DriveApp.getFileById(sheetId_Todo).moveTo(folder);
  sheetTempTodo.appendRow(["Action", "ID", "Nom", "URL", "Taille", "ModifiedISO", "Dossier", "RowToUpdate"]);
  Logger.log(`Sheet temporaire de tâches créé : ${sheetId_Todo}`);

  const sheetLog = SpreadsheetApp.create(NOM_SHEET_LOG_ARCHIVE);
  const sheetId_Log = sheetLog.getId();
  DriveApp.getFileById(sheetId_Log).moveTo(folder);
  sheetLog.appendRow(["Horodatage", "État", "ID Fichier", "Nom Fichier", "Message"]);
  Logger.log(`Fichier journal créé : ${sheetId_Log}`);
  
  const sheetStats = SpreadsheetApp.create(NOM_SHEET_STATS_ARCHIVE);
  const sheetId_Stats = sheetStats.getId();
  DriveApp.getFileById(sheetId_Stats).moveTo(folder);
  sheetStats.insertSheet("Tableau de Bord");
  sheetStats.insertSheet("Historique").appendRow(["Date", "Fichiers Totaux (DB)", "Espace Perdu (Go)", "Fichiers Traités (Cycle)", "Durée Cycle (min)", "Lots Découverte", "Lots Traitement"]);
  sheetStats.insertSheet("Analyse des Formats").appendRow(["Extension", "Description", "Nombre", "Taille (Octets)"]);
  sheetStats.insertSheet("Top 100 Fichiers (Taille)").appendRow(["Nom", "Chemin Complet", "Taille (Mo)", "URL", "Date", "Heure"]);
  sheetStats.insertSheet("Fichiers Ignorés").appendRow(["ID", "Nom", "Chemin/Dossier", "Taille", "Raison"]);
  try { sheetStats.deleteSheet(sheetStats.getSheetByName("Feuille 1")); } catch(e) {}
  Logger.log(`Fichier de Stats créé : ${sheetId_Stats}`);
  
  /* 5. Configurer l'état initial (utilise les clés du script principal) */
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('SHEET_ID_DB', sheetId_DB);
  properties.setProperty('SHEET_ID_TODO', sheetId_Todo); 
  properties.setProperty('SHEET_ID_LOG', sheetId_Log);
  properties.setProperty('SHEET_ID_STATS', sheetId_Stats); 
  properties.setProperty('SHEET_ID_ACTION', sheetAction.getId()); /* Ajouté pour la complétude */
  properties.setProperty('ETAT_SCRIPT', 'DECOUVERTE_INITIALE');
  properties.setProperty('CONTINUATION_TOKEN', DriveApp.getFiles().getContinuationToken());
  properties.setProperty('DB_DATA', JSON.stringify({}));
  properties.setProperty('COMPTEUR_DECOUVERTE', "0"); 
  properties.setProperty('COMPTEUR_TRAITEMENT', "0");
  properties.setProperty('FICHIERS_TRAITES_CYCLE', "0");
  properties.setProperty('DATE_DEBUT_CYCLE', new Date().toISOString());

  /* 6. Créer le premier déclencheur */
  _archive_creerProchainDeclencheur('traiterLotFichiers', 1, 'DECOUVERTE_INITIALE'); 
  
  _archive_logToFile('DECOUVERTE_INITIALE', `Lancement de la PREMIÈRE ANALYSE (Reset Usine). Fichiers dans ${CHEMIN_DOSSIER_ARCHIVE}`);
  
  MailApp.sendEmail(EMAIL_POUR_RAPPORT_ARCHIVE, 
                    "[Drive] Lancement de l'analyse initiale (Reset Usine)", 
                    "L'analyse initiale de tous vos fichiers a recommencé.\n");
}

/* --- Outil 4 : Fonction de Maintenance (Archivée V4.14.1) --- */
function REINITIALISER_ETAT_MANUELLEMENT() {
  const props = PropertiesService.getScriptProperties();
  /* Utilise les constantes du script principal (si elles sont accessibles) ou des chaînes directes */
  const PROP_ETAT_SCRIPT = 'ETAT_SCRIPT';
  const PROP_CACHE_DB_FILE_ID = 'cacheDBFileId';
  const PROP_CONTINUATION_TOKEN = 'CONTINUATION_TOKEN';

  props.setProperty(PROP_ETAT_SCRIPT, 'IDLE');
  props.deleteProperty(PROP_CACHE_DB_FILE_ID); // Nettoie l'ID du cache
  props.deleteProperty(PROP_CONTINUATION_TOKEN); // Nettoie le token

  /* Tente d'appeler la fonction de suppression du script principal si elle existe */
  try {
    supprimerFichierCacheDB(); // Force la suppression de tout vieux fichier cache
    Logger.log("Appel à supprimerFichierCacheDB() réussi.");
  } catch(e) {
    Logger.log("supprimerFichierCacheDB() n'est pas accessible, suppression manuelle tentée (peut échouer si le nom de la constante n'est pas ici).");
    /* Logique de secours si la fonction n'est pas dans le même fichier */
    try {
      const folder = _archive_getOrCreateFolderByPath(CHEMIN_DOSSIER_ARCHIVE);
      const files = folder.getFilesByName("[CACHE] db_lookup.json"); // Nom en dur
      if (files.hasNext()) {
        files.next().setTrashed(true);
        Logger.log("Fichier cache de secours supprimé.");
      }
    } catch (e2) {}
  }

  Logger.log("État du script forcé à IDLE. Cache et déclencheurs nettoyés.");
}


/* --- Fonctions de support (isolées) pour l'archive --- */

/* Fonction logToFile isolée pour éviter les conflits */
function _archive_logToFile(etat, message) {
  Logger.log(`[ARCHIVE LOG] [${etat}] ${message}`);
  try {
    const props = PropertiesService.getScriptProperties();
    const logId = props.getProperty('SHEET_ID_LOG');
    if (logId) {
       SpreadsheetApp.openById(logId).getSheets()[0].appendRow([new Date().toISOString(), etat, "", "", message]);
    }
  } catch (e) {}
}

function _archive_getOrCreateFolderByPath(path) {
  let parts = path.split('/');
  let currentFolder = DriveApp.getRootFolder();
  for (let part of parts) {
    if (part) { 
      let folders = currentFolder.getFoldersByName(part);
      if (folders.hasNext()) currentFolder = folders.next();
      else currentFolder = currentFolder.createFolder(part);
    }
  }
  return currentFolder;
}

function _archive_supprimerSheetParNom(nom, folder) {
  try {
    const fichiers = folder.getFilesByName(nom);
    while (fichiers.hasNext()) {
      fichiers.next().setTrashed(true);
    }
  } catch (e) {}
}

function _archive_creerProchainDeclencheur(fonction, minutes, etat) {
  _archive_supprimerDeclencheursScript();
  ScriptApp.newTrigger(fonction)
      .timeBased()
      .after(minutes * 60 * 1000)
      .create();
  Logger.log(`Prochain lot (${fonction} - État: ${etat || 'INCONNU'}) programmé dans ${minutes} minutes.`);
}

function _archive_supprimerDeclencheursScript() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'traiterLotFichiers' || 
        funcName === 'lancerAnalyseQuotidienne' ||
        funcName === 'lanceurNocturneIntelligent') { 
      ScriptApp.deleteTrigger(trigger);
    }
  }
}
