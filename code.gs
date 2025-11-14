/*
 * =========================================================================================
 * HISTORIQUE DES VERSIONS
 * =========================================================================================
 * SCRIPT AVANCÉ DE RECHERCHE DE DOUBLONS PAR CONTENU
 *
 * ====================================================================
 * HISTORIQUE DES VERSIONS (A.B.C)
 * ====================================================================
 * v4.13.0 - (FEATURE + FIX) Mise à jour majeure (Fix "Argument too large", Rapports enrichis, Sauvegarde DB).
 * v4.13.1 - (FIX) Rétablissement du formatage des commentaires sécurisé.
 * v4.15.0 - (FEATURE) Optimisation "Changes API" (Simplifiée) - Remplace V4.14.
 * v4.15.4 - (PATCH) Correction API V2 (Finalisation des champs V2/V3).
 * -----------------------------------------------------------------------------------------
 * V5.0.0 - Release Majeure : Optimisation I/O Sheets API.
 * -----------------------------------------------------------------------------------------
 * V5.0.1 - Patch de Stabilité (Correction de la faute de frappe DOSSIER_ORPHELINS_ID).
 * -----------------------------------------------------------------------------------------
 * - PRÉREQUIS CRITIQUE: Le service "Drive API" et "Sheets API" DOIVENT être activés.
 * - FIX: Correction d'une faute de frappe critique dans la déclaration de variable (DOSSIER_ORPHELINS_ID).
 * - MAJ: Mise à jour du SCRIPT_VERSION.
 * =========================================================================================
 */

/* --- VARIABLES DE CONFIGURATION --- */
const SCRIPT_VERSION = "V5.0.1"; /* (V5.0.1) */
const EMAIL_POUR_RAPPORT = "toto@hotmail.com";
const CHEMIN_DOSSIER = "script/doublondrive";
const NOM_DOSSIER_ORPHELINS = "FICHIER ORPHELIN";
const NOM_SHEET_DB = `[DB] Hashes Fichiers Drive`;
const NOM_SHEET_ACTION = `[ACTION] Suppression Doublons`;
const NOM_SHEET_TEMP_TODO = `[TEMP] Fichiers à Traiter`; 
const NOM_SHEET_LOG = `[LOG] Journal Analyse Doublons`;
const NOM_SHEET_STATS = `[STATS] Tableau de Bord`; 

const MINUTES_ENTRE_LOTS = 1; 
const TEMPS_MAX_EXECUTION_SECONDES = 270; 

const MAX_FILE_SIZE_BYTES = 47185920; /* 45 Mo */
const BATCH_SIZE_TRAITEMENT = 50; 

/* Noms des propriétés */
const PROP_ETAT_SCRIPT = 'ETAT_SCRIPT';
const PROP_SHEET_ID_DB = 'SHEET_ID_DB'; 
const PROP_SHEET_ID_TODO = 'SHEET_ID_TODO'; 
const PROP_SHEET_ID_LOG = 'SHEET_ID_LOG'; 
const PROP_SHEET_ID_STATS = 'SHEET_ID_STATS';
const PROP_SHEET_ID_ACTION = 'SHEET_ID_ACTION';
const PROP_DOSSIER_ORPHELINS_ID = 'DOSSIER_ORPHELINS_ID';
const PROP_SYNC_TOKEN = 'SYNC_TOKEN'; /* (V4.15.0) */

/* Compteurs de boucle */
const PROP_COMPTEUR_DECOUVERTE = 'COMPTEUR_DECOUVERTE';
const PROP_COMPTEUR_TRAITEMENT = 'COMPTEUR_TRAITEMENT';
const PROP_FICHIERS_TRAITES_CYCLE = 'FICHIERS_TRAITES_CYCLE';
const PROP_DATE_DEBUT_CYCLE = 'DATE_DEBUT_CYCLE';

/* Base de connaissance des formats */
const FORMAT_DESCRIPTIONS = {
  '.pdf': 'Document PDF',
  '.jpg': 'Image JPEG', '.jpeg': 'Image JPEG',
  '.png': 'Image PNG', '.gif': 'Image GIF',
  '.doc': 'Word (Ancien)', '.docx': 'Word',
  '.xls': 'Excel (Ancien)', '.xlsx': 'Excel',
  '.ppt': 'PowerPoint (Ancien)', '.pptx': 'PowerPoint',
  '.txt': 'Fichier Texte', '.csv': 'Fichier CSV',
  '.zip': 'Archive ZIP', '.rar': 'Archive RAR',
  '.mp3': 'Audio MP3', '.wav': 'Audio WAV', '.m4a': 'Audio M4A',
  '.mp4': 'Vidéo MP4', '.mov': 'Vidéo MOV', '.avi': 'Vidéo AVI',
  '.gdoc': 'Google Doc', '.gsheet': 'Google Sheet'
};

/* --- FLUX DE TRAVAIL (DÉPLACÉ V4.15.3) --- */

/**
 * Fonction principale du déclencheur nocturne (ex: 02:00).
 * Vérifie l'état et relance un cycle complet (IDLE) ou le lot suivant.
 */
function lanceurNocturneIntelligent() {
  logToFile("SYSTEM", "Réveil intelligent (02:00).");
  const etat = PropertiesService.getScriptProperties().getProperty(PROP_ETAT_SCRIPT);
  if (etat === 'IDLE' || !etat) {
    lancerAnalyseQuotidienne();
  } else {
    creerProchainDeclencheur('traiterLotFichiers', 1, etat);
  }
}

/* ---------------------------------- */
/* --- FONCTIONS PRINCIPALES --- */
/* ---------------------------------- */

/**
 * Utilitaire pour trouver/créer un dossier par chemin.
 */
function getOrCreateFolderByPath(path) {
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

/**
 * Écrit un log dans le [LOG] Sheet (MODIFIÉ V5.0.0 pour API Sheets).
 */
function logToFile(etat, message) {
  const pattern = /ID (\w+).*\| Nom: (.*?)(?: \| Taille:.*)? \| /;
  const match = message.match(pattern);
  let fichierId = '', fichierNom = '', messageStruct = message;
  if (match) {
    fichierId = match[1]; fichierNom = match[2]; messageStruct = message.replace(match[0], '');
  }
  Logger.log(`[${etat}] ${message}`);
  
  try {
    const sheetId_Log = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID_LOG);
    if (sheetId_Log) {
      /* V5.0.0 : Utilisation de l'API Sheets pour APPEND ultra-rapide */
      Sheets.Spreadsheets.Values.append({
        values: [[new Date().toISOString(), etat, fichierId, fichierNom, messageStruct]]
      }, sheetId_Log, 'Feuille 1', { valueInputOption: 'USER_ENTERED' });
    }
  } catch (e) {
    /* Si l'API Sheets ou le Log Sheet n'est pas prêt, on revient au simple log dans Apps Script */
    Logger.log(`ERREUR LOG V5.0.0: Échec d'écriture dans le Sheet. ${e.message}`);
    PropertiesService.getScriptProperties().deleteProperty(PROP_SHEET_ID_LOG);
  }
}

/* --- FONCTIONS D'OUTILS (SUPPRIMÉES V4.15.1) --- */


/* --- FLUX DE TRAVAIL (MODIFIÉ V5.0.0) --- */

/**
 * Initialise un nouveau cycle d'analyse.
 * Fait une sauvegarde, vide [TEMP] et lance le premier lot V4.15.
 */
function lancerAnalyseQuotidienne() {
  logToFile("QUOTIDIEN", `Lancement analyse ${SCRIPT_VERSION} (API Changes)...`);
  const props = PropertiesService.getScriptProperties();

  /* 1. Nettoyer les anciens déclencheurs temporaires */
  supprimerDeclencheursScript(); 

  /* 2. (v4.13.0) Sauvegarde de la [DB] */
  try {
    const folder = getOrCreateFolderByPath(CHEMIN_DOSSIER);
    const dbFileId = props.getProperty(PROP_SHEET_ID_DB);
    if (dbFileId) {
      const dateStr = new Date().toISOString().slice(0, 10);
      DriveApp.getFileById(dbFileId).makeCopy(`[DB] (Backup ${dateStr})`, folder);
      logToFile("QUOTIDIEN", "Sauvegarde de la [DB] effectuée.");
    }
  } catch(e) { logToFile("ERREUR", `Echec sauvegarde [DB]: ${e.message}`); }

  const sheetId_DB = props.getProperty(PROP_SHEET_ID_DB);
  const sheetId_Todo = props.getProperty(PROP_SHEET_ID_TODO);

  if (!sheetId_DB || !sheetId_Todo) {
    logToFile("ERREUR", "DB non init. Voir Archive.");
    return;
  }
  
  /* 3. Vider le fichier [TEMP] */
  try { 
    /* Utilisation de l'API Sheets pour vider le [TEMP] Sheet */
    Sheets.Spreadsheets.Values.clear({}, sheetId_Todo, 'Feuille 1');
    /* On réécrit l'en-tête */
    Sheets.Spreadsheets.Values.append({
        values: [["Action","ID","Nom","URL","Taille","ISO","Dossier","Row"]]
    }, sheetId_Todo, 'Feuille 1', { valueInputOption: 'USER_ENTERED' });
  } catch (e) {
    logToFile("ERREUR", `Échec de vidage/réécriture de [TEMP]: ${e.message}`);
  }
  
  /* 4. Initialiser les propriétés */
  props.setProperty(PROP_ETAT_SCRIPT, 'DECOUVERTE_SYNCHRO');
  props.setProperty(PROP_COMPTEUR_DECOUVERTE, "0");
  props.setProperty(PROP_COMPTEUR_TRAITEMENT, "0");
  props.setProperty(PROP_FICHIERS_TRAITES_CYCLE, "0");
  props.setProperty(PROP_DATE_DEBUT_CYCLE, new Date().toISOString());

  /* 5. Lancer le premier lot */
  creerProchainDeclencheur('traiterLotFichiers', 1, 'DECOUVERTE_SYNCHRO');
  
  MailApp.sendEmail(EMAIL_POUR_RAPPORT, 
                    `[Drive ${SCRIPT_VERSION}] Lancement de l'analyse quotidienne`, 
                    `L'analyse (API Changes ${SCRIPT_VERSION}) des fichiers nouveaux/modifiés/supprimés a commencé.`);
}

/**
 * Fonction "routeur" appelée par les déclencheurs temporaires.
 */
function traiterLotFichiers() {
  const etat = PropertiesService.getScriptProperties().getProperty(PROP_ETAT_SCRIPT);
  try {
    if (etat === 'DECOUVERTE_SYNCHRO') {
      logiqueDeSynchronisationDesChangements();
    } else if (etat === 'TRAITEMENT') {
      logiqueDeTraitement();
    } else if (etat === 'RAPPORT') {
      genererRapportFinal();
    } else { 
      logToFile("ERREUR", `État inconnu ou obsolète: ${etat}. Réinitialisation à IDLE.`);
      PropertiesService.getScriptProperties().setProperty(PROP_ETAT_SCRIPT, 'IDLE');
      supprimerDeclencheursScript(); 
    }
  } catch (err) {
    logToFile("ERREUR FATALE", `${err.message}`);
    /* AMÉLIORATION V4.15.1 : On ne supprime plus le déclencheur. */
  }
}

/* --- NOUVELLE LOGIQUE DE DÉCOUVERTE (CORRIGÉE V5.0.1) --- */
function logiqueDeSynchronisationDesChangements() {
  const props = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  let compteur = parseInt(props.getProperty(PROP_COMPTEUR_DECOUVERTE) || "0") + 1;
  props.setProperty(PROP_COMPTEUR_DECOUVERTE, compteur.toString());

  const sheetTodo = SpreadsheetApp.openById(props.getProperty(PROP_SHEET_ID_TODO)).getSheets()[0];
  let token = props.getProperty(PROP_SYNC_TOKEN);
  let nouvellesTaches = [];
  let idsASupprimer = [];

  /* Étape 1 : Logique de "Reset Usine" (V4.15.5) */
  if (!token) {
    logToFile("INFO", `Token ${SCRIPT_VERSION} non trouvé. Lancement d'un scan complet 'Reset' (pageToken: null).`);
  }

  /* Étape 2 : Boucler sur les changements */
  let pageToken = token;
  let changements;
  
  try {
    /* V5.0.1 : Correction de la faute de frappe */
    const DOSSIER_ORPHELIN_ID_VALUE = props.getProperty(PROP_DOSSIER_ORPHELINS_ID);
    let dossierOrphelins;
    if (DOSSIER_ORPHELIN_ID_VALUE) dossierOrphelins = DriveApp.getFolderById(DOSSIER_ORPHELIN_ID_VALUE);

    while (pageToken !== undefined) { /* Utiliser undefined si nextPageToken n'existe pas */
      if ((new Date().getTime() - startTime) / 1000 > TEMPS_MAX_EXECUTION_SECONDES) {
        logToFile("DECOUVERTE_SYNCHRO", `Pause (limite temps). ${compteur} lots de changements traités.`);
        break; /* Sortir de la boucle while pour sauvegarder */
      }
      
      /* V4.15.4 : Syntaxe API V2 (title, labels.trashed) */
      changements = Drive.Changes.list({
        pageToken: pageToken,
        fields: "newStartPageToken, nextPageToken, changes(removed, fileId, file(id, title, mimeType, labels(trashed), modifiedDate, webViewLink, fileSize))"
      });

      if (!changements.changes || changements.changes.length === 0) {
         if (changements.nextPageToken) {
           pageToken = changements.nextPageToken;
           continue;
         } else {
           logToFile("INFO", "Aucun changement détecté dans ce lot.");
           pageToken = null; /* Fin de la synchro */
           break;
         }
      }

      for (let chg of changements.changes) {
        /* CAS 1 : Fichier supprimé ou mis à la corbeille (Syntaxe V2) */
        if (chg.removed || (chg.file && chg.file.labels && chg.file.labels.trashed)) {
          const idASupprimer = chg.fileId || (chg.file ? chg.file.id : null);
          if (idASupprimer) {
            idsASupprimer.push(idASupprimer);
          }
          continue;
        }

        /* CAS 2 : Fichier ajouté ou modifié */
        let f = chg.file;
        if (!f || !f.id) continue; /* Ignore les changements sans fichier */
        
        let dossier = "?? API V4.15 ??"; // Sera recalculé par logiqueDeTraitement

        nouvellesTaches.push([
          "MODIFIED",               /* A: Action (MODIFIED par défaut) */
          f.id,                     /* B: ID */
          f.title,                  /* C: Nom (V2 utilise "title") */
          f.webViewLink,            /* D: URL */
          f.fileSize || 0,          /* E: Taille */
          f.modifiedDate,           /* F: ModifiedISO */
          dossier,                  /* G: Dossier (Sera recalculé) */
          ""                        /* H: RowToUpdate (vide par défaut, forçant V4.13 à chercher) */
        ]);
      }
      
      /* Si le temps est écoulé, on sauvegarde le token de cette page */
      if ((new Date().getTime() - startTime) / 1000 > TEMPS_MAX_EXECUTION_SECONDES) {
        props.setProperty(PROP_SYNC_TOKEN, pageToken);
        pageToken = null; /* Arrêter la boucle */
      } else {
        /* Sinon, on passe à la page suivante (le token sera le nextPageToken s'il existe) */
        pageToken = changements.nextPageToken;
      }
    } /* Fin While */

    /* Étape 3 : Écrire les lots de travail (V4.15.5) */
    if (nouvellesTaches.length > 0) {
      sheetTodo.getRange(sheetTodo.getLastRow() + 1, 1, nouvellesTaches.length, nouvellesTaches[0].length).setValues(nouvellesTaches);
      logToFile("DECOUVERTE_SYNCHRO", `Lot de ${nouvellesTaches.length} tâches (NEW/MODIFIED) ajouté à [TEMP].`);
    }
    if (idsASupprimer.length > 0) {
      mettreAJourStatutDB(idsASupprimer, 'SUPPRIMÉ');
      logToFile("DECOUVERTE_SYNCHRO", `${idsASupprimer.length} fichiers marqués comme 'SUPPRIMÉ' dans [DB].`);
    }

    /* Étape 4 : Prochain cycle */
    if (pageToken) {
      /* Le temps était écoulé, on relance */
      creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'DECOUVERTE_SYNCHRO');
    } else {
      /* C'est terminé, on passe au traitement */
      logToFile("DECOUVERTE_SYNCHRO", `Phase DECOUVERTE (${SCRIPT_VERSION}) terminée.`);
      props.setProperty(PROP_SYNC_TOKEN, changements.newStartPageToken); /* Sauvegarde le token final */
      props.setProperty(PROP_ETAT_SCRIPT, 'TRAITEMENT');
      creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'TRAITEMENT');
    }

  } catch (e) {
    logToFile("ERREUR FATALE", `(${SCRIPT_VERSION}) ${e.message}. Token non mis à jour. Reprise au prochain cycle.`);
  }
}

/* --- NOUVEL OUTIL (MODIFIÉ V5.0.0) --- */
/* Met à jour la colonne G (Statut) pour une liste d'IDs (Colonne A) */
function mettreAJourStatutDB(listeIds, statut) {
  if (!listeIds || listeIds.length === 0) return;

  logToFile("INFO", `Mise à jour de ${listeIds.length} statuts vers '${statut}'...`);
  try {
    const props = PropertiesService.getScriptProperties();
    const sheetIdDB = props.getProperty(PROP_SHEET_ID_DB);
    const sheetDB = SpreadsheetApp.openById(sheetIdDB).getSheets()[0];
    const data = sheetDB.getRange("A1:A" + sheetDB.getLastRow()).getValues();
    
    /* 1. Créer un index rapide (Map) des IDs et de leur N° de ligne */
    const idMap = new Map();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) {
        idMap.set(data[i][0], i + 1); /* i+1 = N° de ligne (base 1) */
      }
    }

    /* 2. Préparer les requêtes de mise à jour par lots (API Sheets) */
    const dataToUpdate = [];
    const rangesToUpdate = [];
    
    for (const id of listeIds) {
      if (idMap.has(id)) {
        /* On prépare la valeur à insérer (le statut) et la cellule cible */
        dataToUpdate.push([statut]);
        rangesToUpdate.push(`G${idMap.get(id)}`); /* Colonne G = Statut */
      }
    }

    /* 3. Mettre à jour en un seul appel batch (V5.0.0) */
    if (rangesToUpdate.length > 0) {
      const requests = rangesToUpdate.map((range, index) => ({
          range: range,
          values: [dataToUpdate[index]]
      }));

      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: 'USER_ENTERED',
        data: requests
      }, sheetIdDB);
    }
  } catch (e) {
    logToFile("ERREUR", `(${SCRIPT_VERSION}) Échec de la mise à jour groupée du statut : ${e.message}`);
  }
}


/* --- LOGIQUE DE TRAITEMENT (MODIFIÉE V5.0.0) --- */
/* Cette fonction est appelée APRES la découverte V4.15 */
function logiqueDeTraitement() {
  const props = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();
  let compteur = parseInt(props.getProperty(PROP_COMPTEUR_TRAITEMENT) || "0") + 1;
  props.setProperty(PROP_COMPTEUR_TRAITEMENT, compteur.toString());
  
  const sheetIdDB = props.getProperty(PROP_SHEET_ID_DB);
  const sheetDB = SpreadsheetApp.openById(sheetIdDB).getSheets()[0];
  const sheetTodo = SpreadsheetApp.openById(props.getProperty(PROP_SHEET_ID_TODO)).getSheets()[0];
  
  /* V4.15 : Créer l'index de la DB (pour trouver les RowToUpdate) */
  const dbData = sheetDB.getDataRange().getValues();
  const dbMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    if (dbData[i][0]) {
      dbMap.set(dbData[i][0], i + 1); /* Map<ID, RowIndex> */
    }
  }
  
  let taches = [];
  try {
    const data = sheetTodo.getDataRange().getValues();
    data.shift();
    if (data.length === 0) {
      props.setProperty(PROP_ETAT_SCRIPT, 'RAPPORT');
      creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'RAPPORT');
      return;
    }
    taches = data.slice(0, BATCH_SIZE_TRAITEMENT);
  } catch (e) { return; }
  
  logToFile("TRAITEMENT", `Lot #${compteur} (${taches.length} fichiers).`);
  
  let lignesDB_New = [];
  let updateRequests = []; /* V5.0.0 : Pour les MODIFIED via batchUpdate */
  let i = 0;
  
  while (i < taches.length) {
    if ((new Date().getTime() - startTime) / 1000 > TEMPS_MAX_EXECUTION_SECONDES) break;
    
    const t = taches[i]; /* [Action, ID, Nom, URL, Taille, ISO, Dossier, RowToUpdate] */
    let hash = '', cheminComplet = '', dateStr = '', heureStr = '';
    
    try {
      if (t[4] > MAX_FILE_SIZE_BYTES) {
        hash = 'IGNORÉ - Fichier trop volumineux';
      } else {
        const f = DriveApp.getFileById(t[1]);
        const mime = f.getMimeType();
        
        /* v4.13.0 : Calcul du chemin complet et Timestamp (toujours fait) */
        cheminComplet = getCheminComplet(f);
        const dateObj = new Date(t[5]); /* Utilise l'ISO de l'API (t[5]) */
        dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
        heureStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "HH:mm:ss");
        
        /* V4.15 : Le dossier est recalculé ici pour remplacer "?? API V4.15 ??" */
        const parents = f.getParents();
        const dossier = parents.hasNext() ? parents.next().getName() : NOM_DOSSIER_ORPHELINS;
        t[6] = dossier; /* Met à jour le nom du dossier */

        if (t[4] === 0 || mime === MimeType.SHORTCUT || mime.includes('google-apps')) {
          hash = 'IGNORÉ - Type Google/Vide';
        } else {
          hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, f.getBlob().getBytes())
                 .map(b => ((b+256)%256).toString(16).padStart(2,'0')).join('');
        }
      }
    } catch (e) { hash = `ERREUR - ${e.message}`; }
    
    const row = [t[1], t[2], t[3], t[4], t[5], t[6], hash, cheminComplet, dateStr, heureStr];
    
    /* V4.15 : Décide si c'est NEW ou MODIFIED en cherchant la RowToUpdate */
    let rowToUpdate = dbMap.get(t[1]); /* Cherche l'ID dans la Map V4.15 */
    
    if (rowToUpdate) {
      /* C'est MODIFIED (V5.0.0 : Ajout à la liste de batch pour MODIFIED) */
      updateRequests.push({
          range: `A${rowToUpdate}:J${rowToUpdate}`,
          values: [row]
      });
    } else {
      /* C'est NEW */
      lignesDB_New.push(row);
    }
    
    if(hash.startsWith('IGNORÉ') || hash.startsWith('ERREUR')) logToFile("TRAITEMENT", `ID ${t[1]} | ${hash}`);
    i++;
  }
  
  /* V5.0.0 : Écriture des MODIFIED (Batch Update) et des NEW (Append) */
  try {
    if (updateRequests.length > 0) {
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: 'USER_ENTERED',
        data: updateRequests
      }, sheetIdDB);
    }
    
    if (lignesDB_New.length > 0) {
       Sheets.Spreadsheets.Values.append({
        values: lignesDB_New
       }, sheetIdDB, 'Feuille 1', { valueInputOption: 'USER_ENTERED' });
    }
  } catch (e) {
    logToFile("ERREUR CRITIQUE", `Échec écriture DB (V5.0.0 Batch) : ${e.message}`);
    /* Note: Les lignes TO-DO ne sont pas supprimées, elles seront traitées au prochain lot. */
  }

  if (i > 0) sheetTodo.deleteRows(2, i);
  
  let total = parseInt(props.getProperty(PROP_FICHIERS_TRAITES_CYCLE) || "0") + i;
  props.setProperty(PROP_FICHIERS_TRAITES_CYCLE, total.toString());
  
  creerProchainDeclencheur('traiterLotFichiers', MINUTES_ENTRE_LOTS, 'TRAITEMENT');
}


/* Fonction utilitaire pour le chemin complet (INCHANGÉE - V4.13.0) */
function getCheminComplet(fichier) {
  try {
    let pathParts = [fichier.getName()];
    let folder = fichier.getParents().hasNext() ? fichier.getParents().next() : null;
    let rootId = DriveApp.getRootFolder().getId();
    while (folder && folder.getId() !== rootId) {
        pathParts.unshift(folder.getName());
        folder = folder.getParents().hasNext() ? folder.getParents().next() : null;
    }
    pathParts.unshift("Mon Drive");
    return pathParts.join(' / ');
  } catch(e) { return "[Chemin inaccessible]"; }
}

/* LOGIQUE DE NETTOYAGE (SUPPRIMÉE - V4.15.0) */


/* --- GÉNÉRATION DES RAPPORTS (MODIFIÉ V5.0.0) --- */
function genererRapportFinal() {
  logToFile("RAPPORT", `Génération des rapports (${SCRIPT_VERSION})...`);
  const props = PropertiesService.getScriptProperties();
  const folder = getOrCreateFolderByPath(CHEMIN_DOSSIER);
  
  let totalFichiersDB = 0, espaceTotal = 0, totalDoublons = 0, espacePerdu = 0;
  let doublonsCount = 0, perduSize = 0;
  
  let sheetId_DB = props.getProperty(PROP_SHEET_ID_DB);
  let sheetId_Stats = props.getProperty(PROP_SHEET_ID_STATS);
  let sheetId_Action = props.getProperty(PROP_SHEET_ID_ACTION);

  /* Auto-réparation des IDs manquants (CORRIGÉ V4.15.2) */
  if (!sheetId_DB || !sheetId_Stats || !sheetId_Action) {
    logToFile("ERREUR", "IDs de rapport manquants. Exécutez 'RESTAURER_PROPRIETES_RAPPORT' depuis Outils.gs");
    sheetId_DB = props.getProperty(PROP_SHEET_ID_DB);
    sheetId_Stats = props.getProperty(PROP_SHEET_ID_STATS);
    sheetId_Action = props.getProperty(PROP_SHEET_ID_ACTION);
    
    if (!sheetId_DB || !sheetId_Stats || !sheetId_Action) {
       logToFile("ERREUR CRITIQUE", "Impossible de continuer sans les IDs de rapport.");
       return; /* Arrêt propre */
    }
  }

  let dataDB;
  try {
    const sheetDB = SpreadsheetApp.openById(sheetId_DB).getSheets()[0];
    dataDB = sheetDB.getDataRange().getValues();
    dataDB.shift(); 
  } catch (e) { logToFile("ERREUR", `DB inaccessible: ${e.message}`); return; }

  const mapHashes = {}, fichiersIgnores = [], fichiers0ko = [], mapFormats = {}, top100 = [];

  for (const row of dataDB) {
    totalFichiersDB++;
    /* [ID, Nom, URL, Taille, ISO, Dossier, Hash, Chemin, Date, Heure] */
    const id=row[0], nom=row[1], url=row[2], taille=parseFloat(row[3])||0, dossier=row[5], hash=row[6], chemin=row[7], date=row[8], heure=row[9];
    
    espaceTotal += taille;
    const nomFichierSafe = String(nom || ''); 
    const extension = (nomFichierSafe.includes('.')) ? nomFichierSafe.substring(nomFichierSafe.lastIndexOf('.')).toLowerCase() : "[Aucune extension]";
    if (!mapFormats[extension]) mapFormats[extension] = { count: 0, size: 0, desc: (FORMAT_DESCRIPTIONS[extension] || "Fichier " + extension) };
    mapFormats[extension].count++;
    mapFormats[extension].size += taille;

    top100.push({ nom: nom, dossier: dossier, taille: taille, chemin: chemin, url: url, date: date, heure: heure });
    
    if (taille === 0 && (!hash || !hash.startsWith("IGNORÉ"))) {
      fichiers0ko.push([false, nom, chemin, 0, url, id, date, heure]);
    }
    else if (hash && (hash.startsWith("IGNORÉ") || hash.startsWith("ERREUR") || hash.startsWith("SUPPRIMÉ"))) {
      /* V4.15 : On ne compte plus les "SUPPRIMÉ" comme des "Ignorés" */
      if (hash !== 'SUPPRIMÉ') {
        fichiersIgnores.push([id, nom, chemin, taille, hash]);
      }
    } 
    else if (hash) {
      const infoFichier = { id: id, nom: nom, url: url, taille: taille, hash: hash, dossier: dossier, chemin: chemin, date: date, heure: heure };
      if (!mapHashes[hash]) mapHashes[hash] = [];
      mapHashes[hash].push(infoFichier);
    }
  }

  /* --- [ACTION] --- */
  let sheetActionUrl = "";
  try {
    const actionSpreadsheet = SpreadsheetApp.openById(sheetId_Action);
    sheetActionUrl = actionSpreadsheet.getUrl();
    
    /* Onglet Doublons */
    let sheetDoublons = actionSpreadsheet.getSheetByName("Doublons");
    if(!sheetDoublons) sheetDoublons = actionSpreadsheet.insertSheet("Doublons", 0);
    sheetDoublons.clear(); 
    sheetDoublons.appendRow(["EFFACER", "Nom", "Chemin Complet", "Taille", "Date", "Heure", "URL", "ID", "Hash"]);
    sheetDoublons.setFrozenRows(1);
    
    let actionRows = [];
    for (const h in mapHashes) {
      const grp = mapHashes[h];
      if (grp.length > 1) {
        doublonsCount += grp.length;
        perduSize += grp[0].taille * (grp.length - 1);
        grp.forEach(f => actionRows.push([false, f.nom, f.chemin, f.taille, f.date, f.heure, f.url, f.id, f.hash]));
        actionRows.push(["", "", "", "", "", "", "", "", ""]); 
      }
    }
    if(actionRows.length>0) {
      sheetDoublons.getRange(2,1,actionRows.length,9).setValues(actionRows);
      sheetDoublons.getRange(2,1,actionRows.length,1).insertCheckboxes();
    }

    /* Onglet 0 Ko */
    let sheet0ko = actionSpreadsheet.getSheetByName("Fichiers 0 ko");
    if (!sheet0ko) sheet0ko = actionSpreadsheet.insertSheet("Fichiers 0 ko", 1);
    sheet0ko.clear();
    sheet0ko.appendRow(["EFFACER", "Nom", "Chemin", "Taille", "URL", "ID", "Date", "Heure"]);
    if(fichiers0ko.length > 0) {
      sheet0ko.getRange(2,1,fichiers0ko.length,8).setValues(fichiers0ko);
      sheet0ko.getRange(2,1,fichiers0ko.length,1).insertCheckboxes();
    }

    /* Onglet Dossiers Vides (v4.13.0) */
    let sheetDossiersVides = actionSpreadsheet.getSheetByName("Dossiers Vides");
    if (!sheetDossiersVides) sheetDossiersVides = actionSpreadsheet.insertSheet("Dossiers Vides", 2);
    sheetDossiersVides.clear();
    sheetDossiersVides.appendRow(["Nom", "Chemin Complet", "URL", "ID"]);
    const dossiersVides = chercherDossiersVides(); // Scan lent
    if(dossiersVides.length > 0) {
      sheetDossiersVides.getRange(2,1,dossiersVides.length,4).setValues(dossiersVides);
    }
    
    logToFile("RAPPORT", `Sheet [ACTION] mis à jour. ${doublonsCount} doublons, ${fichiers0ko.length} fichiers 0ko, ${dossiersVides.length} dossiers vides.`);
  } catch(e) { logToFile("ERREUR", `Erreur Action Sheet: ${e.message}`); }

  /* --- [STATS] --- */
  let statsSheetUrl = "";
  try {
    const statsSheet = SpreadsheetApp.openById(sheetId_Stats);
    statsSheetUrl = statsSheet.getUrl();
    
    /* Ignorés */
    let sheetIgn = statsSheet.getSheetByName("Fichiers Ignorés");
    if(!sheetIgn) sheetIgn = statsSheet.insertSheet("Fichiers Ignorés");
    sheetIgn.clear();
    sheetIgn.appendRow(["ID", "Nom", "Chemin/Dossier", "Taille", "Raison"]);
    if(fichiersIgnores.length > 0) sheetIgn.getRange(2,1,fichiersIgnores.length,5).setValues(fichiersIgnores);

    /* Tableau de Bord */
    let tdbSheet = statsSheet.getSheetByName("Tableau de Bord");
    tdbSheet.clear();
    const dateFinCycle = new Date();
    const dateDebutCycle = new Date(props.getProperty(PROP_DATE_DEBUT_CYCLE) || dateFinCycle);
    const dureeCycleMs = dateFinCycle.getTime() - dateDebutCycle.getTime();
    const dureeCycleMin = (dureeCycleMs / 1000 / 60).toFixed(2);
    const espacePerduGo = (perduSize / 1024 / 1024 / 1024).toFixed(2);
    const espaceTotalGo = (espaceTotal / 1024 / 1024 / 1024).toFixed(2);
    
    tdbSheet.appendRow(["Statistique", "Valeur"]);
    tdbSheet.appendRow(["Dernière Analyse", dateFinCycle.toLocaleString("fr-BE")]);
    tdbSheet.appendRow(["Durée Totale du Cycle (min)", dureeCycleMin]);
    tdbSheet.appendRow(["Fichiers Traités (ce cycle)", props.getProperty(PROP_FICHIERS_TRAITES_CYCLE) || "0"]);
    tdbSheet.appendRow(["---", "---"]);
    tdbSheet.appendRow(["Fichiers Totaux (dans la DB)", totalFichiersDB]);
    tdbSheet.appendRow(["Espace Disque Total (Go)", espaceTotalGo]);
    tdbSheet.appendRow(["Fichiers en Doublon", totalDoublons]);
    tdbSheet.appendRow(["Espace Perdu (Go)", espacePerduGo]);
    tdbSheet.appendRow(["Fichiers Ignorés/Erreurs", fichiersIgnores.length]);
    tdbSheet.appendRow(["Fichiers 0 ko", fichiers0ko.length]);
    tdbSheet.getRange("A1:B1").setFontWeight("bold");

    /* Historique */
    let histSheet = statsSheet.getSheetByName("Historique");
    histSheet.appendRow([dateFinCycle.toISOString().slice(0,10), totalFichiersDB, espacePerduGo, props.getProperty(PROP_FICHIERS_TRAITES_CYCLE) || "0", dureeCycleMin, props.getProperty(PROP_COMPTEUR_DECOUVERTE), props.getProperty(PROP_COMPTEUR_TRAITEMENT)]);

    /* Formats */
    let fmtSheet = statsSheet.getSheetByName("Analyse des Formats");
    fmtSheet.clear();
    fmtSheet.appendRow(["Extension", "Description", "Nombre", "Taille (Octets)"]);
    let fmtRows = Object.entries(mapFormats).map(([k,v]) => [k, v.desc, v.count, v.size]);
    fmtRows.sort((a,b)=>b[2]-a[2]);
    if (fmtRows.length > 0) fmtSheet.getRange(2,1,fmtRows.length,4).setValues(fmtRows);

    /* Top 100 */
    let topSheet = statsSheet.getSheetByName("Top 100 Fichiers (Taille)");
    topSheet.clear();
    topSheet.appendRow(["Nom", "Chemin", "Taille (Mo)", "URL", "Date", "Heure"]);
    top100.sort((a,b)=>b.taille-a.taille);
    let topRows = top100.slice(0,100).map(f => [f.nom, f.chemin, (f.taille/1024/1024).toFixed(2), f.url, f.date, f.heure]);
    if (topRows.length > 0) topSheet.getRange(2,1,topRows.length,6).setValues(topRows);

    logToFile("RAPPORT", `Sheet [STATS] mis à jour.`);
  } catch(e) { logToFile("ERREUR", `Erreur Stats Sheet: ${e.message}`); }

  /* --- Email final --- */
  let corps = `Analyse des doublons terminée (${SCRIPT_VERSION}).\n\nFichiers dans la base: ${totalFichiersDB}\nDoublons trouvés: ${doublonsCount}\nEspace perdu: ${(perduSize/1024/1024).toFixed(2)} Mo.`;
  if(doublonsCount>0 || fichiers0ko.length>0) corps += `\n\nOuvrez [ACTION] pour supprimer : ${sheetActionUrl}`;
  corps += `\n\nTableau de bord [STATS] : ${statsSheetUrl}`;
  corps += `\n\nDossier de travail : ${folder.getUrl()}`;
  
  MailApp.sendEmail(EMAIL_POUR_RAPPORT, `[Drive ${SCRIPT_VERSION}] Rapport d'analyse terminé`, corps);
  
  logToFile("RAPPORT", "Terminé. En attente de la prochaine exécution nocturne.");
  
  /* Nettoyage final */
  supprimerDeclencheursScript();
  props.setProperty(PROP_ETAT_SCRIPT, 'IDLE');
}

/* v4.13.0 : Nouvelle fonction pour les dossiers vides (INCHANGÉE) */
function chercherDossiersVides() {
  logToFile("RAPPORT", "Scan des dossiers vides en cours (peut être long)...");
  let dossiersVides = [];
  try {
    const folders = DriveApp.getFolders();
    while (folders.hasNext()) {
      const folder = folders.next();
      if (!folder.getFiles().hasNext() && !folder.getFolders().hasNext()) {
        if (folder.isTrashed()) continue; // Ignorer ceux déjà supprimés
        
        // Calculer le chemin
        let pathParts = [folder.getName()];
        let parent = folder.getParents().hasNext() ? folder.getParents().next() : null;
        let rootId = DriveApp.getRootFolder().getId();
        while (parent && parent.getId() !== rootId) {
            pathParts.unshift(parent.getName());
            parent = parent.getParents().hasNext() ? parent.getParents().next() : null;
        }
        pathParts.unshift("Mon Drive");
        
        dossiersVides.push([folder.getName(), pathParts.join(' / '), folder.getUrl(), folder.getId()]);
      }
    }
  } catch (e) { logToFile("ERREUR", `Echec scan dossiers vides: ${e.message}`); }
  return dossiersVides;
}


/* --- GESTION DES DÉCLENCHEURS (INCHANGÉE - V4.13.0) --- */

function creerProchainDeclencheur(func, min, etat) {
  supprimerDeclencheursScript();
  ScriptApp.newTrigger(func).timeBased().after(min * 60 * 1000).create();
  Logger.log(`[${etat}] Prochain lot dans ${min} min.`);
}

function supprimerDeclencheursScript() {
  /* v4.13.0 : Ne supprime QUE les déclencheurs temporaires. */
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    /* Ne supprime que 'traiterLotFichiers', laisse 'lanceurNocturneIntelligent' intact. */
    if(t.getHandlerFunction() === 'traiterLotFichiers') {
      ScriptApp.deleteTrigger(t);
    }
  }
}
