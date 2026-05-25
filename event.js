// event.js - TEST tylko dla Rafała
console.log("EVENT JS RAFAL TEST");

function onNewMessageComposeHandler(event) {
  try {
    const email = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();

    if (email !== "rafal.sieradzki@familijna.pl") {
      event.completed();
      return;
    }

    const html =
      '<div style="border:3px solid red;padding:10px;color:red;font-size:16px;">' +
      'TEST EVENT RAFAL DZIALA' +
      '</div>';

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
