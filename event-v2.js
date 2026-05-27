// event-v2.js / event.js - diagnostyka storage checkboxa
// Uzywaj tylko przy przypisaniu aplikacji do jednego testowego uzytkownika.

const DIAG_MARKER = 'data-gf-storage-diagnostic-v3="1"';
const AUTO_KEY = "autoSignatureEnabled";
const PROFILE_KEY = "signatureUserProfile";

function safeText(value) {
  if (value === null) return "null";
  if (value === undefined) return "undefined";
  try { return String(value); } catch (e) { return "[unreadable]"; }
}

function getLocalStorageValue(key) {
  try {
    if (window.localStorage) {
      return window.localStorage.getItem(key);
    }
    return "localStorage unavailable";
  } catch (e) {
    return "localStorage ERROR: " + (e && e.message ? e.message : String(e));
  }
}

function getOfficeRuntimeStorageValue(key, callback) {
  try {
    if (window.OfficeRuntime && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      OfficeRuntime.storage.getItem(key)
        .then(function(value) {
          callback(null, value);
        })
        .catch(function(err) {
          callback("OfficeRuntime.storage ERROR: " + (err && err.message ? err.message : String(err)), null);
        });
      return;
    }

    callback("OfficeRuntime.storage unavailable", null);
  } catch (e) {
    callback("OfficeRuntime.storage ERROR: " + (e && e.message ? e.message : String(e)), null);
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

function insertDiagnostic(event, officeAuto, officeAutoErr, officeProfile, officeProfileErr) {
  try {
    const localAuto = getLocalStorageValue(AUTO_KEY);
    const localProfile = getLocalStorageValue(PROFILE_KEY);

    const html =
      '<div data-gf-storage-diagnostic-v3="1" style="font-family:Calibri,Arial;font-size:11pt;border:1px solid #666;padding:8px;margin:8px 0;background:#fff;color:#000;">' +
      '<b>DIAGNOSTYKA STORAGE STOPKI GF v3</b><br><br>' +
      '<b>Event:</b> DZIALA<br>' +
      '<b>OfficeRuntime.storage autoSignatureEnabled:</b> ' + safeText(officeAutoErr || officeAuto) + '<br>' +
      '<b>localStorage autoSignatureEnabled:</b> ' + safeText(localAuto) + '<br><br>' +
      '<b>OfficeRuntime.storage signatureUserProfile:</b> ' + safeText(officeProfileErr || officeProfile).substring(0, 500) + '<br>' +
      '<b>localStorage signatureUserProfile:</b> ' + safeText(localProfile).substring(0, 500) + '<br><br>' +
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
  try {
    getOfficeRuntimeStorageValue(AUTO_KEY, function(autoErr, autoValue) {
      getOfficeRuntimeStorageValue(PROFILE_KEY, function(profileErr, profileValue) {
        insertDiagnostic(event, autoValue, autoErr, profileValue, profileErr);
      });
    });
  } catch (e) {
    try { event.completed(); } catch (_) {}
  }
}

function onNewMessageComposeHandler(event) {
  runDiagnostic(event);
}

function onNewMessageComposeHandlerV2(event) {
  runDiagnostic(event);
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
