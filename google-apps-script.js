/**
 * MultiSoumissions Toiture — Google Apps Script
 *
 * Étapes :
 * 1. Ouvre https://script.google.com/
 * 2. Crée un nouveau projet.
 * 3. Colle ce fichier dans Code.gs.
 * 4. Ajuste OWNER_EMAIL au besoin.
 * 5. Déploie comme application Web.
 *
 * Déploiement recommandé :
 * - Exécuter en tant que : Moi
 * - Qui a accès : Tout le monde
 */

const SPREADSHEET_ID = '16R8aGNR0owSprq7gJDZN7ngV3bZsDq_pssCa2ERlsmY';
const CLIENT_SHEET = 'Leads clients';
const CONTRACTOR_SHEET = 'Contracteurs partenaires';
const OWNER_EMAIL = 'multisoumission514@gmail.com';

function doPost(e) {
  try {
    const body = parseBody_(e);
    const type = String(body.type_de_demande || body.form_type || '').trim();

    if (type === 'Contracteur partenaire') {
      return json_(handleContractor_(body));
    }

    return json_(handleLead_(body));
  } catch (error) {
    return json_({
      ok: false,
      message: error && error.message ? error.message : 'Erreur inconnue.'
    });
  }
}

function doGet() {
  return json_({
    ok: true,
    service: 'MultiSoumissions Toiture',
    message: 'Apps Script actif.'
  });
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  const raw = e.postData.contents;

  if (e.postData.type && e.postData.type.indexOf('application/json') !== -1) {
    return JSON.parse(raw);
  }

  return raw.split('&').reduce(function(acc, pair) {
    const parts = pair.split('=');
    const key = decodeURIComponent(parts[0] || '').replace(/\+/g, ' ');
    const value = decodeURIComponent(parts.slice(1).join('=') || '').replace(/\+/g, ' ');
    if (key) acc[key] = value;
    return acc;
  }, {});
}

function handleLead_(data) {
  const name = clean_(data.nom_complet);
  const email = clean_(data.courriel);
  const phone = clean_(data.telephone);
  const address = clean_(data.adresse_complete);
  const city = clean_(data.ville);
  const postal = clean_(data.code_postal).toUpperCase();
  const language = clean_(data.langue_client) === 'Anglais' ? 'Anglais' : 'Français';

  if (!name || !isEmail_(email) || !isPhone_(phone) || !address || !city || !postal) {
    throw new Error('Champs obligatoires invalides.');
  }

  const row = {
    'Date': today_(),
    'Type de demande': 'Client toiture',
    'Nom': name,
    'Courriel': email,
    'Téléphone': phone,
    'Adresse': address,
    'Ville': city,
    'Code postal': postal,
    'Langue du client': language,
    'Besoin': clean_(data.type_de_besoin),
    'Type de toiture': clean_(data.type_de_toiture_actuelle),
    'Matériau souhaité': clean_(data.materiau_souhaite),
    'Superficie ou dimensions': clean_(data.superficie_ou_dimensions),
    'Délai des travaux': clean_(data.delai_des_travaux),
    'Commentaires': clean_(data.commentaires),
    'Photos fournies': 'Non - proposées sur la page merci',
    'Statut': 'Nouveau'
  };

  appendObjectRow_(CLIENT_SHEET, row);
  sendOwnerLeadEmail_(row);
  sendClientConfirmationEmail_(row);

  return { ok: true, type: 'lead' };
}

function handleContractor_(data) {
  const company = clean_(data.entreprise);
  const contact = clean_(data.nom_du_contact);
  const rbq = clean_(data.numero_rbq);
  const phone = clean_(data.telephone_contracteur);
  const email = clean_(data.courriel_contracteur);
  const regions = clean_(data.regions_desservies);

  if (!company || !contact || !rbq || !isPhone_(phone) || !isEmail_(email) || !regions) {
    throw new Error('Champs obligatoires invalides.');
  }

  const row = {
    'Date': today_(),
    'Entreprise': company,
    'Contact': contact,
    'RBQ': rbq,
    'Téléphone': phone,
    'Courriel': email,
    'Régions desservies': regions,
    'Message': clean_(data.message_contracteur),
    'Statut': 'Nouveau'
  };

  appendObjectRow_(CONTRACTOR_SHEET, row);
  sendOwnerContractorEmail_(row);

  return { ok: true, type: 'contractor' };
}

function appendObjectRow_(sheetName, rowObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Onglet introuvable : ' + sheetName);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(function(header) {
    return rowObject[header] || '';
  });
  sheet.appendRow(row);
}

function sendOwnerLeadEmail_(row) {
  const subject = 'Nouveau lead toiture - ' + row['Nom'] + ' - ' + row['Ville'];
  const body = [
    'Nouveau lead toiture reçu.',
    '',
    'Nom : ' + row['Nom'],
    'Courriel : ' + row['Courriel'],
    'Téléphone : ' + row['Téléphone'],
    'Adresse : ' + row['Adresse'] + ', ' + row['Ville'] + ', ' + row['Code postal'],
    'Langue du client : ' + (row['Langue du client'] || 'Français'),
    '',
    'Besoin : ' + row['Besoin'],
    'Type de toiture : ' + row['Type de toiture'],
    'Matériau souhaité : ' + (row['Matériau souhaité'] || 'Non applicable / non précisé'),
    'Superficie ou dimensions : ' + row['Superficie ou dimensions'],
    'Délai des travaux : ' + row['Délai des travaux'],
    'Commentaires : ' + (row['Commentaires'] || 'Aucun commentaire'),
    'Photos : à demander sur la page de confirmation si le client en a.',
    '',
    'Statut Google Sheets : Nouveau'
  ].join('\n');

  MailApp.sendEmail(OWNER_EMAIL, subject, body);
}

function sendClientConfirmationEmail_(row) {
  if (row['Langue du client'] === 'Anglais') {
    const subjectEn = 'Your roofing quote request has been received';
    const bodyEn = [
      'Hello ' + row['Nom'] + ',',
      '',
      'Thank you for submitting your roofing quote request. We have received it.',
      '',
      'Here are the next steps:',
      '',
      '1. Your request is reviewed to understand the roof type, location and preferred timing.',
      '2. Relevant information may be shared with partner entrepreneurs who hold an RBQ licence.',
      '3. Depending on availability in your area, contractors may contact you directly within 24 to 48 hours.',
      '4. You compare prices and choose freely. You are not obligated to accept a quote.',
      '',
      'If you have photos of your roof, you may add them after submitting your request to help contractors better assess the project.',
      '',
      'Keep your phone nearby: contractors will call you directly if your project matches their area and availability.',
      '',
      'MultiSoumissions Roofing'
    ].join('\n');

    MailApp.sendEmail(row['Courriel'], subjectEn, bodyEn);
    return;
  }

  const subject = 'Votre demande de soumission toiture a bien été reçue';
  const body = [
    'Bonjour ' + row['Nom'] + ',',
    '',
    'Merci d’avoir rempli votre demande de soumission pour vos travaux de toiture. Nous l’avons bien reçue.',
    '',
    'Voici les prochaines étapes :',
    '',
    '1. Votre demande est examinée afin de bien comprendre le type de toiture, le secteur et le délai souhaité.',
    '2. Les informations pertinentes peuvent être transférées à des entrepreneurs partenaires détenant une licence RBQ.',
    '3. Selon les disponibilités dans votre région, des contracteurs pourraient vous contacter directement dans les 24 à 48 heures.',
    '4. Vous comparez les prix et choisissez librement. Vous n’êtes pas obligé d’accepter une soumission.',
    '',
    'Si vous avez des photos de votre toiture, vous pourrez les ajouter après l’envoi de votre demande pour aider les contracteurs à mieux évaluer le projet.',
    '',
    'Gardez votre téléphone à portée de main : les entrepreneurs vous appelleront directement si votre projet correspond à leur secteur et à leurs disponibilités.',
    '',
    'MultiSoumissions Toiture'
  ].join('\n');

  MailApp.sendEmail(row['Courriel'], subject, body);
}

function sendOwnerContractorEmail_(row) {
  const subject = 'Nouveau contracteur partenaire - ' + row['Entreprise'];
  const body = [
    'Nouvelle demande de partenariat contracteur.',
    '',
    'Entreprise : ' + row['Entreprise'],
    'Contact : ' + row['Contact'],
    'RBQ : ' + row['RBQ'],
    'Téléphone : ' + row['Téléphone'],
    'Courriel : ' + row['Courriel'],
    'Régions desservies : ' + row['Régions desservies'],
    'Message : ' + (row['Message'] || 'Aucun message')
  ].join('\n');

  MailApp.sendEmail(OWNER_EMAIL, subject, body);
}

function clean_(value) {
  return String(value || '').trim();
}

function isEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean_(email));
}

function isPhone_(phone) {
  return clean_(phone).replace(/\D/g, '').length === 10;
}

function today_() {
  return Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm');
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
