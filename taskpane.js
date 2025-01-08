Office.onReady(() => {
    document.getElementById("addSignature").onclick = addSignature;
});

async function addSignature() {
    const signatureHTML = `
        <table style="font-family:Calibri, Arial; font-size:10pt;">
            <tr>
                <td>
                    <b>Jan Kowalski</b><br />
                    Stanowisko<br />
                    <b>Email:</b> jan.kowalski@przyklad.pl<br />
                    <b>Telefon:</b> 123-456-789
                </td>
            </tr>
        </table>
    `;

    Office.context.mailbox.item.body.setAsync(
        signatureHTML,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Stopka została dodana.");
            } else {
                console.error("Błąd podczas dodawania stopki:", result.error.message);
            }
        }
    );
}
