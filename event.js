// event.js - prosty test bez czerwonego tekstu, przez getAsync + setAsync
// Ten wariant używa mechanizmu, który wcześniej zadziałał przy czerwonym teście.

const SIMPLE_MARKER = 'data-gf-simple-test="1"';

function onNewMessageComposeHandler(event) {
  try {
    const insertHtml =
      '<div data-gf-simple-test="1">Pozdrawiam,<br>__</div><br><br>';

    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: event },
      function (getResult) {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
          getResult.asyncContext.completed();
          return;
        }

        const currentBody = getResult.value || "";

        if (currentBody.indexOf(SIMPLE_MARKER) !== -1) {
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

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
