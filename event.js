/*
 * MINIMALNY TEST EVENT-BASED ACTIVATION
 * WERSJA: EVENT MINIMAL TEST 2026-06-01 01
 * Cel: sprawdzic, czy Outlook w ogole uruchamia onNewMessageComposeHandler.
 * Bez roamingSettings, bez Graph, bez warunkow, bez stopki produkcyjnej.
 */

function completeEvent(event) {
  try {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  } catch (e) {
    // Brak akcji - event runtime nie pokazuje bledow uzytkownikowi.
  }
}

function onNewMessageComposeHandler(event) {
  try {
    var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;

    if (!item || !item.body || typeof item.body.prependAsync !== "function") {
      completeEvent(event);
      return;
    }

    var html =
      '<div style="border:2px solid #d00000;color:#d00000;background:#fff3f3;padding:8px;margin:8px 0;font-family:Arial,sans-serif;font-size:13px;">' +
      '<b>TEST EVENT DZIALA</b><br>' +
      'EVENT MINIMAL TEST 2026-06-01 01<br>' +
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
      Office.context.mailbox.item.body.prependAsync(
        '<div style="border:2px solid #d00000;color:#d00000;padding:8px;margin:8px 0;font-family:Arial,sans-serif;font-size:13px;">' +
        '<b>TEST EVENT ERROR</b><br>' +
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
  // Jezeli associate sie nie uda, Outlook nie odpali handlera.
}
