// event-v2.js - nowa funkcja i nowy plik, aby ominac cache Outlook runtime
// Test: wstawia Pozdrawiam + __ przy nowej wiadomosci

const SIMPLE_MARKER_V2 = 'data-gf-simple-test-v2="1"';

function onNewMessageComposeHandlerV2(event) {
  try {
    const insertHtml =
      '<div data-gf-simple-test-v2="1">Pozdrawiam,<br>__</div><br><br>';

    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: event },
      function (getResult) {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
          getResult.asyncContext.completed();
          return;
        }

        const currentBody = getResult.value || "";

        if (currentBody.indexOf(SIMPLE_MARKER_V2) !== -1) {
          getResult.asyncContext.completed();
          return;
        }

        const newBody = insertHtml + currentBody;

        Office.context.mailbox.item.body.setAsync(
          newBody,
          {
            coercionType: Office.CoercionType.Html,
            asyncContext: getResult.asyncContext
          },
          function () {
            getResult.asyncContext.completed();
          }
        );
      }
    );
  } catch (e) {
    try { event.completed(); } catch (_) {}
  }
}

Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
