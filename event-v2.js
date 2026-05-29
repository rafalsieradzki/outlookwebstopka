// Stopka Familijna - diagnostyka eventu
// WERSJA: BOTH EVENTS DIAG NOTATKI 2026-05-29 03
// Bezpieczenstwo: tekst diagnostyczny jest wstawiany tylko dla mailbox.userProfile.emailAddress == notatki@familijna.pl

const GF_DIAG_ALLOWED_MAIL = "notatki@familijna.pl";
const GF_DIAG_VERSION = "BOTH EVENTS DIAG NOTATKI 2026-05-29 03";
const GF_AUTO_KEY = "autoSignatureEnabled";
const GF_PROFILE_KEY = "signatureUserProfile";
const GF_MARKER = 'data-familijna-event-diag="2026-05-29-03"';

function gfDone(event) {
  try {
    if (event && event.completed) event.completed();
  } catch (e) {}
}

function gfNormalizeMail(value) {
  return (value || "").toString().trim().toLowerCase();
}

function gfGetMailboxMail() {
  try {
    var profile = Office.context.mailbox.userProfile || {};
    return gfNormalizeMail(profile.emailAddress || profile.userPrincipalName || "");
  } catch (e) {
    return "";
  }
}

function gfRoamingGet(key) {
  try {
    if (!Office.context || !Office.context.roamingSettings) return null;
    return Office.context.roamingSettings.get(key);
  } catch (e) {
    return null;
  }
}

function gfGetRoamingMail() {
  var raw = gfRoamingGet(GF_PROFILE_KEY);
  if (!raw) return "";
  try {
    var p = typeof raw === "string" ? JSON.parse(raw) : raw;
    return gfNormalizeMail(p.mail || p.userPrincipalName || "");
  } catch (e) {
    return "";
  }
}

function gfGetBodyHtml() {
  return new Promise(function(resolve, reject) {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          reject(result.error || new Error("body.getAsync failed"));
        }
      }
    );
  });
}

function gfSetBodyHtml(html) {
  return new Promise(function(resolve, reject) {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error || new Error("body.setAsync failed"));
        }
      }
    );
  });
}

function gfHtmlEscape(value) {
  return (value || "").toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;");
}

function gfBuildDiagHtml(handlerName, mailboxMail, roamingMail, enabled) {
  return [
    '<div ' + GF_MARKER + ' style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#DF292F;border:2px solid #DF292F;padding:8px;margin:8px 0;background:#fff7f7;">',
    '<b>GF EVENT OK</b><br>',
    'version=' + gfHtmlEscape(GF_DIAG_VERSION) + '<br>',
    'handler=' + gfHtmlEscape(handlerName) + '<br>',
    'allowed=' + gfHtmlEscape(GF_DIAG_ALLOWED_MAIL) + '<br>',
    'mailboxMail=' + gfHtmlEscape(mailboxMail || '(brak)') + '<br>',
    'roamingMail=' + gfHtmlEscape(roamingMail || '(brak)') + '<br>',
    'autoSignatureEnabled=' + gfHtmlEscape(String(enabled)) + '<br>',
    'time=' + gfHtmlEscape(new Date().toISOString()),
    '</div>'
  ].join('');
}

async function gfDiagnosticHandler(event, handlerName) {
  try {
    var mailboxMail = gfGetMailboxMail();
    var roamingMail = gfGetRoamingMail();

    // Najwazniejszy bezpiecznik: nie wstawiaj nic nikomu poza notatki@familijna.pl.
    // Dla diagnostyki nie ufamy zapisanemu profilowi; wymagana jest zgodnosc biezacej skrzynki Outlooka.
    if (mailboxMail !== GF_DIAG_ALLOWED_MAIL) {
      gfDone(event);
      return;
    }

    var enabled = gfRoamingGet(GF_AUTO_KEY);
    var body = await gfGetBodyHtml();

    if (body.indexOf(GF_MARKER) !== -1) {
      gfDone(event);
      return;
    }

    await gfSetBodyHtml(body + '<br><br>' + gfBuildDiagHtml(handlerName, mailboxMail, roamingMail, enabled));
    gfDone(event);
  } catch (e) {
    gfDone(event);
  }
}

function onNewMessageComposeHandler(event) {
  gfDiagnosticHandler(event, "onNewMessageComposeHandler");
}

function onNewMessageComposeHandlerV2(event) {
  gfDiagnosticHandler(event, "onNewMessageComposeHandlerV2");
}

try {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
} catch (e) {}

try {
  Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
} catch (e) {}
