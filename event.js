/*
 * MINIMALNY TEST EVENT-BASED ACTIVATION - TYLKO NOTATKI
 * WERSJA: EVENT MINIMAL TEST NOTATKI ONLY 2026-06-01 02
 * Cel: sprawdzic, czy Outlook uruchamia onNewMessageComposeHandler.
 * Bez Graph, bez stopki produkcyjnej.
 * Zabezpieczenie: testowy tekst pojawi sie tylko dla notatki@familijna.pl.
 */

function completeEvent(event) {
  try {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  } catch (e) {
  }
}

function normalizeMail(value) {
  try {
    return String(value || "").toLowerCase().trim();
  } catch (e) {
    return "";
  }
}

function getAllowedMailFromRoamingSettings() {
  try {
    var profileRaw = Office.context.roamingSettings.get("signatureUserProfile");
    if (!profileRaw) {
      return "";
    }

    var profile = profileRaw;
    if (typeof profileRaw === "string") {
      profile = JSON.parse(profileRaw);
    }

    return normalizeMail(profile.mail || profile.userPrincipalName || "");
  } catch (e) {
    return "";
  }
}

function isAllowedUser() {
  var allowed = "notatki@familijna.pl";

  var mailboxMail = "";
  try {
    mailboxMail = normalizeMail(
      Office.context &&
      Office.context.mailbox &&
      Office.context.mailbox.userProfile &&
      Office.context.mailbox.userProfile.emailAddress
    );
  } catch (e) {
    mailboxMail = "";
  }

  var roamingMail = getAllowedMailFromRoamingSettings();

  return mailboxMail === allowed || roamingMail === allowed;
}

function onNewMessageComposeHandler(event) {
  try {
    if (!isAllowedUser()) {
      completeEvent(event);
      return;
    }

    var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;

    if (!item || !item.body || typeof item.body.prependAsync !== "function") {
      completeEvent(event);
      return;
    }

    var html =
      '<div style="border:2px solid #d00000;color:#d00000;background:#fff3f3;padding:8px;margin:8px 0;font-family:Arial,sans-serif;font-size:13px;">' +
      '<b>TEST EVENT DZIALA - NOTATKI ONLY</b><br>' +
      'EVENT MINIMAL TEST NOTATKI ONLY 2026-06-01 02<br>' +
      'Handler: onNewMessageComposeHandler' +
      '</div><br>';

    item.body.prependAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function () {
        completeEvent(event);
      }
    );
  } catch (e) {
    try {
      if (!isAllowedUser()) {
        completeEvent(event);
        return;
      }

      Office.context.mailbox.item.body.prependAsync(
        '<div style="border:2px solid #d00000;color:#d00000;padding:8px;margin:8px 0;font-family:Arial,sans-serif;font-size:13px;">' +
        '<b>TEST EVENT ERROR - NOTATKI ONLY</b><br>' +
        String(e && e.message ? e.message : e) +
        '</div><br>',
        { coercionType: Office.CoercionType.Html },
        function () {
          completeEvent(event);
        }
      );
    } catch (ignored) {
      completeEvent(event);
    }
  }
}

try {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
} catch (e) {
}
