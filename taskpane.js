Office.onReady(function () {
  var button = document.getElementById("insertSignature");
  if (button) {
    button.onclick = insertSignature;
  }
});

function replaceAllSafe(text, token, value) {
  return text.split(token).join(value || "");
}

function buildSignatureHtml() {
  var profile = Office.context.mailbox.userProfile || {};

  var displayName = profile.displayName || "";
  var email = profile.emailAddress || "";

  var title = "";
  var phoneNumber = "";
  var mobileNumber = "";

  var html = "<p>%%DisplayName%%<br>%%Email%%</p>";

  html = replaceAllSafe(html, "%%DisplayName%%", displayName);
  html = replaceAllSafe(html, "%%Email%%", email);
  html = replaceAllSafe(html, "%%Title%%", title);
  html = replaceAllSafe(html, "%%PhoneNumber%%", phoneNumber);
  html = replaceAllSafe(html, "%%MobileNumber%%", mobileNumber);

  return "<br><br>" + html;
}

function insertSignature() {
  var html = buildSignatureHtml();

  Office.context.mailbox.item.body.setSelectedDataAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Stopka wstawiona.");
      } else {
        console.error("Nie udało się wstawić stopki:", result.error);
        alert("Nie udało się wstawić stopki: " + result.error.message);
      }
    }
  );
}
