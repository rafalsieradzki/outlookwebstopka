// event.js - diagnostyka LaunchEvent
// Ten plik NIE sprawdza checkboxa, NIE używa Graph, NIE używa storage.
// Ma tylko udowodnić, czy Outlook uruchamia OnNewMessageCompose.
// Jeśli event działa, w nowej wiadomości pojawi się czerwony marker diagnostyczny.

const DEBUG_MARKER = 'data-familijna-event-debug="1"';

function onNewMessageComposeHandler(event) {
  try {
    const debugHtml =
      '<div data-familijna-event-debug="1" ' +
      'style="border:2px solid #DF292F;padding:8px;margin:8px 0;color:#DF292F;font-family:Calibri,Arial;font-size:12pt;">' +
      'TEST EVENTU STOPKI: OnNewMessageCompose uruchomił event.js' +
      '</div>';

    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: event },
      function (getResult) {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
          getResult.asyncContext.completed();
          return;
        }

        const currentBody = getResult.value || "";

        if (currentBody.indexOf(DEBUG_MARKER) !== -1) {
          getResult.asyncContext.completed();
          return;
        }

        const newBody = debugHtml + "<br>" + currentBody;

        Office.context.mailbox.item.body.setAsync(
          newBody,
          {
            coercionType: Office.CoercionType.Html,
            asyncContext: getResult.asyncContext
          },
          function (setResult) {
            setResult.asyncContext.completed();
          }
        );
      }
    );
  } catch (e) {
    try {
      event.completed();
    } catch (_) {}
  }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
