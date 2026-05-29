// Stopka Familijna - event.js
// DIAGNOSTYKA TYLKO DLA: notatki@familijna.pl
// Zgodne z manifest.xml 3.0.0.1:
// LaunchEvent OnNewMessageCompose -> onNewMessageComposeHandler

const GF_DIAG_ALLOWED_MAIL = "notatki@familijna.pl";
const STORAGE_AUTO_KEY = "autoSignatureEnabled";
const STORAGE_PROFILE_KEY = "signatureUserProfile";
const DIAG_MARKER = 'data-familijna-event-diag="1"';

function eventDone(event) {
  try {
    if (event && event.completed) event.completed();
  } catch (e) {}
}

function roamingGet(key) {
  try {
    if (!Office.context || !Office.context.roamingSettings) return null;
    return Office.context.roamingSettings.get(key);
  } catch (e) {
    return null;
  }
}

function normalizeMail(value) {
  return (value || "").toString().trim().toLowerCase();
}

function getMailboxMail() {
  try {
    const p = Office.context.mailbox.userProfile || {};
    return normalizeMail(p.emailAddress || p.userPrincipalName || "");
  } catch (e) {
    return "";
  }
}

function getRoamingProfile() {
  const raw = roamingGet(STORAGE_PROFILE_KEY);
  if (!raw) return null;

  try {
    return typeof raw === "string" ? JSON.parse(raw) : raw;
  } catch (e) {
    return null;
  }
}

function getRoamingProfileMail(profile) {
  if (!profile) return "";
  return normalizeMail(profile.mail || profile.userPrincipalName || "");
}

function getBodyHtml() {
  return new Promise(function (resolve, reject) {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          const msg = result.error && result.error.message ? result.error.message : "Nieznany blad body.getAsync.";
          reject(new Error(msg));
        }
      }
    );
  });
}

function setBodyHtml(html) {
  return new Promise(function (resolve, reject) {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          const msg = result.error && result.error.message ? result.error.message : "Nieznany blad body.setAsync.";
          reject(new Error(msg));
        }
      }
    );
  });
}

function buildDiagHtml(mailboxMail, roamingMail, enabled) {
  const time = new Date().toISOString();
  return [
    '<div ' + DIAG_MARKER + ' style="font-family:Calibri,Arial;font-size:11pt;color:#DF292F;border:1px solid #DF292F;padding:8px;margin:8px 0;">',
    '<b>GF EVENT OK - DIAG NOTATKI 2026-05-29 02</b><br>',
    'allowed=' + GF_DIAG_ALLOWED_MAIL + '<br>',
    'mailboxMail=' + (mailboxMail || '(brak)') + '<br>',
    'roamingMail=' + (roamingMail || '(brak)') + '<br>',
    'autoSignatureEnabled=' + enabled + '<br>',
    'time=' + time,
    '</div>'
  ].join('');
}

async function onNewMessageComposeHandler(event) {
  try {
    const mailboxMail = getMailboxMail();
    const profile = getRoamingProfile();
    const roamingMail = getRoamingProfileMail(profile);

    // Bezpieczenstwo: diagnostyka dziala tylko dla notatki@familijna.pl.
    // Dopuszczamy identyfikacje po mailbox.userProfile albo po profilu zapisanym w roamingSettings.
    if (mailboxMail !== GF_DIAG_ALLOWED_MAIL && roamingMail !== GF_DIAG_ALLOWED_MAIL) {
      eventDone(event);
      return;
    }

    const enabled = roamingGet(STORAGE_AUTO_KEY);
    const bodyHtml = await getBodyHtml();

    if (bodyHtml.indexOf(DIAG_MARKER) !== -1) {
      eventDone(event);
      return;
    }

    await setBodyHtml(bodyHtml + '<br><br>' + buildDiagHtml(mailboxMail, roamingMail, enabled));
    eventDone(event);
  } catch (e) {
    eventDone(event);
  }
}

Office.actions.associate(
  "onNewMessageComposeHandler",
  onNewMessageComposeHandler
);
