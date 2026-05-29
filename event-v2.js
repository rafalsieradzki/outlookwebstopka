// event.js - diagnostyka Office.context.roamingSettings
// Uzywac tylko na testowych uzytkownikach.

const DIAG_MARKER = 'data-gf-roaming-diagnostic="1"';

function safeText(value) {
  if (value === null) return "null";
  if (value === undefined) return "undefined";
  try { return String(value); } catch (e) { return "[unreadable]"; }
}

function getRoamingValue(key) {
  try {
    if (Office && Office.context && Office.context.roamingSettings) {
      const value = Office.context.roamingSettings.get(key);
      return value === undefined || value === null ? null : String(value);
    }
    return "roamingSettings unavailable";
  } catch (e) {
    return "roamingSettings ERROR: " + (e && e.message ? e.message : String(e));
  }
}

function getOfficeRuntimeValue(key, callback) {
  try {
    if (window.OfficeRuntime && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      OfficeRuntime.storage.getItem(key)
        .then(function(value) { callback(null, value); })
        .catch(function(err) { callback("OfficeRuntime ERROR: " + safeText(err && err.message ? err.message : err), null); });
      return;
    }
    callback("OfficeRuntime.storage unavailable", null);
  } catch (e) {
    callback("OfficeRuntime ERROR: " + (e && e.message ? e.message : String(e)), null);
  }
}

function getOfficeUserProfileHtml() {
  try {
    const profile = Office.context.mailbox.userProfile || {};
    return [
      "displayName: " + safeText(profile.displayName),
      "emailAddress: " + safeText(profile.emailAddress)
    ].join("<br>");
  } catch (e) {
    return "userProfile ERROR: " + (e && e.message ? e.message : String(e));
  }
}

function insertDiagnostic(event, officeRuntimeAuto, officeRuntimeErr) {
  try {
    const roamingAuto = getRoamingValue("autoSignatureEnabled");
    const roamingProfile = getRoamingValue("signatureUserProfile");

    const html =
      '<div data-gf-roaming-diagnostic="1" style="font-family:Calibri,Arial;font-size:11pt;border:1px solid #555;padding:8px;margin:8px 0;background:#fff;color:#000;">' +
      '<b>DIAGNOSTYKA ROAMING SETTINGS GF</b><br><br>' +
      '<b>Event:</b> DZIALA<br>' +
      '<b>roamingSettings autoSignatureEnabled:</b> ' + safeText(roamingAuto) + '<br>' +
      '<b>OfficeRuntime.storage autoSignatureEnabled:</b> ' + safeText(officeRuntimeErr || officeRuntimeAuto) + '<br><br>' +
      '<b>roamingSettings signatureUserProfile:</b> ' + safeText(roamingProfile).substring(0, 700) + '<br><br>' +
      '<b>Office userProfile:</b><br>' + getOfficeUserProfileHtml() + '<br><br>' +
      '<b>Czas:</b> ' + new Date().toISOString() +
      '</div><br>';

    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: event },
      function(getResult) {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
          getResult.asyncContext.completed();
          return;
        }

        const currentBody = getResult.value || "";

        if (currentBody.indexOf(DIAG_MARKER) !== -1) {
          getResult.asyncContext.completed();
          return;
        }

        Office.context.mailbox.item.body.setAsync(
          html + currentBody,
          {
            coercionType: Office.CoercionType.Html,
            asyncContext: getResult.asyncContext
          },
          function() {
            getResult.asyncContext.completed();
          }
        );
      }
    );
  } catch (e) {
    try { event.completed(); } catch (_) {}
  }
}

function runDiagnostic(event) {
  getOfficeRuntimeValue("autoSignatureEnabled", function(err, value) {
    insertDiagnostic(event, value, err);
  });
}

function onNewMessageComposeHandler(event) {
  runDiagnostic(event);
}

function onNewMessageComposeHandlerV2(event) {
  runDiagnostic(event);
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
