Office.onReady(() => {
    document.getElementById("addSignature").onclick = addSignature;
});

function addSignature() {
    const signature = "<p>Twoja stopka HTML</p>";
    Office.context.mailbox.item.body.setAsync(
        signature,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Stopka dodana");
            } else {
                console.error("Błąd:", result.error.message);
            }
        }
    );
}
