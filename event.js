// Stopka Familijna - diagnostyka eventu
// WERSJA: EVENT DIAG NOTATKI RELAXED 2026-06-01 01
// Bezpiecznik: diagnostyka dziala tylko, gdy biezace konto Outlooka LUB profil zapisany w roamingSettings to notatki@familijna.pl.

const GF_DIAG_ALLOWED_MAIL = "notatki@familijna.pl";
const GF_DIAG_VERSION = "EVENT DIAG NOTATKI RELAXED 2026-06-01 01";
const GF_AUTO_KEY = "autoSignatureEnabled";
const GF_PROFILE_KEY = "signatureUserProfile";
const GF_MARKER = 'data-familijna-event-diag="2026-06-01-01"';

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

function gfBoolText(value) {
  return value === true ? "true" : value === false ? "false" : String(value);
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
    'autoSignatureEnabled=' + gfHtmlEscape(gfBoolText(enabled)) + '<br>',
    'time=' + gfHtmlEscape(new Date().toISOString()),
    '</div>'
  ].join('');
}

function gfBodyGetHtml() {
  return new Promise(function(resolve, reject) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value || "");
      else reject(result.error || new Error("body.getAsync failed"));
    });
  });
}

function gfBodySetHtml(html) {
  return new Promise(function(resolve, reject) {
    Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error || new Error("body.setAsync failed"));
    });
  });
}

function gfSetSignatureHtml(html) {
  return new Promise(function(resolve, reject) {
    try {
      if (!Office.context.mailbox.item.body.setSignatureAsync) {
        reject(new Error("setSignatureAsync unavailable"));
        return;
      }
      Office.context.mailbox.item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(result.error || new Error("setSignatureAsync failed"));
      });
    } catch (e) {
      reject(e);
    }
  });
}

async function gfInsertDiagHtml(html) {
  // Najpierw metoda dedykowana do stopki. Jesli klient jej nie obsluguje, fallback do body.get/set.
  try {
    await gfSetSignatureHtml(html);
    return;
  } catch (e) {}

  var body = await gfBodyGetHtml();
  if (body.indexOf(GF_MARKER) !== -1) return;
  await gfBodySetHtml(body + '<br><br>' + html);
}

async function gfDiagnosticHandler(event, handlerName) {
  try {
    var mailboxMail = gfGetMailboxMail();
    var roamingMail = gfGetRoamingMail();
    var enabled = gfRoamingGet(GF_AUTO_KEY);

    // Bezpiecznik: dopuszczamy diagnostyke tylko dla notatki@familijna.pl,
    // ale sprawdzamy i biezace konto Outlooka, i profil zapisany w roamingSettings.
    if (mailboxMail !== GF_DIAG_ALLOWED_MAIL && roamingMail !== GF_DIAG_ALLOWED_MAIL) {
      gfDone(event);
      return;
    }

    var diagHtml = gfBuildDiagHtml(handlerName, mailboxMail, roamingMail, enabled);
    await gfInsertDiagHtml(diagHtml);
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
