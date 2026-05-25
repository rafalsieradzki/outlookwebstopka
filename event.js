// event.js - prosty test eventu bez czerwonego tekstu
console.log("EVENT JS SIMPLE TEST");

function onNewMessageComposeHandler(event) {
  try {

    const html =
      'Pozdrawiam,<br>__<br><br>';

    Office.context.mailbox.item.body.prependAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function () {
        event.completed();
      }
    );

  } catch (e) {
    try { event.completed(); } catch (_) {}
  }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
