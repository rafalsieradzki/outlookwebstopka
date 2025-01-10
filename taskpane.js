Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("add-signature").onclick = addSignature;
  }
});

function addSignature() {
  try {
    const item = Office.context.mailbox.item;
    const signature = "<p>Best regards,<br>Your Company</p>";
    item.body.appendOnSendAsync(signature, { coercionType: Office.CoercionType.Html }, function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Błąd podczas dodawania podpisu: " + asyncResult.error.message);
      } else {
        console.log("Podpis dodany pomyślnie.");
      }
    });
  } catch (error) {
    console.error("Wystąpił nieoczekiwany błąd: " + error.message);
  }
}
