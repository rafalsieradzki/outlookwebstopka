Office.onReady(function () {
  // Office.js gotowy.
});

function replaceAllSafe(text, token, value) {
  return text.split(token).join(value || "");
}

function buildSignatureHtml() {
  var profile = Office.context.mailbox.userProfile || {};

  var displayName = profile.displayName || "";
  var email = profile.emailAddress || "";

  // Na start ustawiamy puste wartości. Później podepniemy Graph/API dla stanowiska i telefonów.
  var title = "";
  var phoneNumber = "";
  var mobileNumber = "";

  var html = "<!-- Wklej tutaj HTML stopki -->";

  html = replaceAllSafe(html, "%%DisplayName%%", displayName);
  html = replaceAllSafe(html, "%%Email%%", email);
  html = replaceAllSafe(html, "%%Title%%", title);
  html = replaceAllSafe(html, "%%PhoneNumber%%", phoneNumber);
  html = replaceAllSafe(html, "%%MobileNumber%%", mobileNumber);

  return "<br><br>" + html;
}

function insertSignature(event) {
  try {
    var html = buildSignatureHtml();

    Office.context.mailbox.item.body.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Nie udało się wstawić stopki:", result.error);
        }
        event.completed();
      }
    );
  } catch (e) {
    console.error(e);
    event.completed();
  }
}

Office.actions.associate("insertSignature", insertSignature);
