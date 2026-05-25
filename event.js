// event.js - wersja normalna v2, bez czerwonego testowego komunikatu.
// Automatyczne wstawianie stopki przy OnNewMessageCompose.
console.log("EVENT JS VERSION 2");

const SIGNATURE_MARKER = 'data-familijna-signature="2"';
const DEBUG_MARKER = 'data-familijna-event-debug="1"';

function replaceAllSafe(text, token, value) {
  return text.split(token).join(value || "");
}

function buildSignatureHtml() {
  const profile = Office.context.mailbox.userProfile || {};
  const displayName = profile.displayName || "";
  const email = profile.emailAddress || "";

  let html = "\n<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"max-width:520px;font-family:Calibri, Arial;\">\n  <tr>\n    <td style=\"margin:auto;width:220px;\" align=\"center\">\n      <img src=\"https://www.familijna.pl/uploads/drive/familijna_logotyp.png\" width=\"60%\" alt=\"GRUPA FAMILIJNA\" />\n    </td>\n    <td style=\"font-size:9pt;line-height:140%;color:#595959;border-left:3px solid #DF292F;padding-left:15px;\">\n      <span style=\"font-size:14pt;color:#DF292F;\">%%DisplayName%%</span><br />\n      <span>%%Title%%</span><br /><br />\n      <a href=\"https://familijna.pl\" style=\"color:#595959;text-decoration:none;\"><span style=\"color:#DF292F;\">www.</span>familijna.pl</a>\n      <span style=\"color:#DF292F;\">email:</span>\n      <a href=\"mailto:%%Email%%\" style=\"color:#595959;text-decoration:none;\">%%Email%%</a>\n      <br />\n      %%PhoneHtml%%\n      <div style=\"padding-top:25px;\">\n        <a href=\"https://www.facebook.com/familijna\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/fb.png\" height=\"25\" width=\"25\" alt=\"facebook\" style=\"margin-right:5px;\" /></a>&nbsp;\n        <a href=\"https://www.instagram.com/familijna/\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/ig.png\" height=\"25\" width=\"25\" alt=\"instagram\" style=\"margin-right:5px;\" /></a>&nbsp;\n        <a href=\"https://m.me/familijna\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/ms.png\" height=\"25\" width=\"25\" alt=\"messenger\" style=\"margin-right:5px;\" /></a>&nbsp;\n        <a href=\"https://goo.gl/maps/kpEMXw6deUcjidot9\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/gm.png\" height=\"25\" width=\"25\" alt=\"google maps\" style=\"margin-right:5px;\" /></a>&nbsp;\n        <a href=\"https://www.youtube.com/@familijna1631/featured\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/yt.png\" height=\"25\" width=\"25\" alt=\"youtube\" style=\"margin-right:5px;\" /></a>&nbsp;\n        <a href=\"https://www.linkedin.com/company/familijna\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/in.png\" height=\"25\" width=\"25\" alt=\"linkedin\" style=\"margin-right:5px;\" /></a>&nbsp;\n      </div>\n    </td>\n  </tr>\n</table>\n<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"900\" style=\"width:900px;max-width:900px;font-family:Calibri, Arial;margin-top:6px;\">\n  <tr>\n    <td style=\"font-size:7pt;line-height:120%;color:#595959;\">\n      <p style=\"margin:0 0 8px 0;\"><span style=\"color:#DF292F;\">GRUPA FAMILIJNA</span> Spółka z ograniczoną odpowiedzialnością, Kuźnica Czeszycka 11, 56-320 Krośnice, tel. 71 384 56 13</p>\n      <p style=\"margin:0 0 20px 0;\">NIP: 9161351695, REGON: 020182505, BDO: 000084673.</p>\n      <p style=\"margin:0 0 8px 0;\">Informacja dla odbiorcy: Informacje zawarte w niniejszym email-u oraz załącznikach do niego mają charakter poufny, są przeznaczone wyłącznie dla wskazanych adresatów. Jeśli nie są Państwo adresatem tego email-a, prosimy niezwłocznie o jego skasowanie oraz poinformowanie nadawcy. Wykonywanie kopii, ujawnienie, dystrybucja lub używanie niniejszego email-a do innych celów jest zabronione. Spółka Grupa Familijna Sp. z o.o. nie ponosi żadnej odpowiedzialności za zmiany email-a dokonane po jego wysłaniu.</p>\n      <p style=\"margin:0;\">Administratorem danych osobowych jest Grupa Familijna sp. z o.o. z siedzibą w Kuźnicy Czeszyckiej. Dane osobowe zawarte w korespondencji mailowej są przetwarzane w celu odpowiadania na pytania, dokonywania ustaleń, zawierania i realizacji umów z kontrahentami, rozpoznawania reklamacji, jak również ustalenia, dochodzenia i obrony roszczeń. Mają Państwo w szczególności prawo dostępu do swoich danych osobowych, żądania ich usunięcia i wniesienia sprzeciwu wobec przetwarzania danych. Szczegóły dotyczące przetwarzania danych osobowych i przysługujących praw znajdują się w <a href=\"https://www.grupafamilijna.pl/pl/polityka-prywatnosci\" style=\"color:#0645AD;text-decoration:underline;\">Polityce prywatności</a>.</p>\n    </td>\n  </tr>\n</table>\n";

  html = replaceAllSafe(html, "%%DisplayName%%", displayName);
  html = replaceAllSafe(html, "%%Email%%", email);
  html = replaceAllSafe(html, "%%Title%%", "");
  html = replaceAllSafe(html, "%%PhoneHtml%%", "");
  html = replaceAllSafe(html, "%%PhoneNumber%%", "");
  html = replaceAllSafe(html, "%%MobileNumber%%", "");
  html = replaceAllSafe(html, "%%Department%%", "");
  html = replaceAllSafe(html, "%%OfficeLocation%%", "");
  html = replaceAllSafe(html, "%%CompanyName%%", "");

  return '<div data-familijna-signature="2">' + html + '</div>';
}

function onNewMessageComposeHandler(event) {
  try {
    const signatureHtml = buildSignatureHtml();

    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: event },
      function (getResult) {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
          getResult.asyncContext.completed();
          return;
        }

        const currentBody = getResult.value || "";

        if (currentBody.indexOf(SIGNATURE_MARKER) !== -1 || currentBody.indexOf(DEBUG_MARKER) !== -1) {
          getResult.asyncContext.completed();
          return;
        }

        const newBody = currentBody + "<br><br>" + signatureHtml;

        Office.context.mailbox.item.body.setAsync(
          newBody,
          { coercionType: Office.CoercionType.Html, asyncContext: getResult.asyncContext },
          function (setResult) {
            setResult.asyncContext.completed();
          }
        );
      }
    );
  } catch (e) {
    try { event.completed(); } catch (_) {}
  }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
