// Stopka Familijna - event.js
// Zgodne z manifest.xml 3.0.0.1:
// LaunchEvent OnNewMessageCompose -> onNewMessageComposeHandler

const STORAGE_AUTO_KEY = "autoSignatureEnabled";
const STORAGE_PROFILE_KEY = "signatureUserProfile";
const SIGNATURE_MARKER = 'data-familijna-signature="1"';

function eventDone(event) {
  try {
    if (event && event.completed) event.completed();
  } catch (e) {}
}

function roamingGet(key) {
  try {
    if (!Office.context || !Office.context.roamingSettings) return null;
    return Office.context.roamingSettings.get(key);
  } catch (e) {
    return null;
  }
}

function replaceAllSafe(text, token, value) { return text.split(token).join(value || ""); }

function firstBusinessPhone(user) {
  return user && user.businessPhones && user.businessPhones.length > 0 ? (user.businessPhones[0] || "") : "";
}

function buildPhoneHtml(phoneNumber, mobileNumber) {
  const parts = [];
  if (phoneNumber) parts.push('<span style="color:#DF292F;">tel.</span> ' + phoneNumber);
  if (mobileNumber) parts.push('<span style="color:#DF292F;">kom.</span> ' + mobileNumber);
  return parts.join(" ");
}

function buildSignatureHtml(user) {
  const officeProfile = Office.context.mailbox.userProfile || {};
  user = user || {};

  const displayName = user.displayName || officeProfile.displayName || "";
  const email = user.mail || user.userPrincipalName || officeProfile.emailAddress || "";
  const title = user.jobTitle || "";
  const phoneNumber = firstBusinessPhone(user);
  const mobileNumber = user.mobilePhone || "";
  const department = user.department || "";
  const officeLocation = user.officeLocation || "";
  const companyName = user.companyName || "";
  const phoneHtml = buildPhoneHtml(phoneNumber, mobileNumber);

  let html = "\n<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"max-width:520px;font-family:Calibri, Arial;\">\n    <tr>\n        <td style=\"margin:auto;width:220px;\" align=\"center\">\n            <img src=\"https://www.familijna.pl/uploads/drive/familijna_logotyp.png\" width=\"80%\" alt=\"GRUPA FAMILIJNA\" />\n        </td>\n        <td style=\"font-size:9pt;line-height:140%;color:#595959;border-left:3px solid #DF292F;padding-left:15px;\">\n            <span style=\"font-size:14pt;color:#DF292F;\">%%DisplayName%%</span>\n            <br />\n            <span>%%Title%%</span>\n            <br /><br />\n            <a href=\"https://familijna.pl\" style=\"color:#595959;text-decoration: none;\"><span style=\"color:#DF292F;\">www.</span>familijna.pl</a>\n            <span style=\"color:#DF292F;\">email:</span>\n            <a href=\"mailto:%%Email%%\" style=\"color:#595959;text-decoration: none;\">%%Email%%</a>\n            <br />\n            %%PhoneHtml%%\n            <div style=\"padding-top:25px;\">\n                <a href=\"https://www.facebook.com/familijna\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/fb.png\" height=\"25\" width=\"25\" alt=\"facebook\" style=\"margin-right:5px;\" /></a>&nbsp;\n                <a href=\"https://www.instagram.com/familijna/\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/ig.png\" height=\"25\" width=\"25\" alt=\"instagram\" style=\"margin-right:5px;\" /></a>&nbsp;\n                <a href=\"https://m.me/familijna\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/ms.png\" height=\"25\" width=\"25\" alt=\"messenger\" style=\"margin-right:5px;\" /></a>&nbsp;\n                <a href=\"https://goo.gl/maps/kpEMXw6deUcjidot9\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/gm.png\" height=\"25\" width=\"25\" alt=\"google maps\" style=\"margin-right:5px;\" /></a>&nbsp;\n                <a href=\"https://www.youtube.com/@familijna1631/featured\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/yt.png\" height=\"25\" width=\"25\" alt=\"youtube\" style=\"margin-right:5px;\" /></a>&nbsp;\n                <a href=\"https://www.linkedin.com/company/familijna\" style=\"display:inline-block;\"><img src=\"https://www.familijna.pl/uploads/drive/in.png\" height=\"25\" width=\"25\" alt=\"linkedin\" style=\"margin-right:5px;\" /></a>&nbsp;\n            </div>\n        </td>\n    </tr>\n</table>\n\n<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"900\" style=\"width:900px;max-width:900px;font-family:Calibri, Arial;margin-top:6px;\">\n    <tr>\n        <td style=\"font-size:7pt;line-height:120%;color:#595959;\">\n            <p style=\"margin:0 0 8px 0;\"><span style=\"color:#DF292F;\">GRUPA FAMILIJNA</span> Spółka z ograniczoną odpowiedzialnością, Kuźnica Czeszycka 11, 56-320 Krośnice, tel. 71 384 56 13</p>\n            <p style=\"margin:0 0 20px 0;\">NIP: 9161351695, REGON: 020182505, BDO: 000084673.</p>\n            <p style=\"margin:0 0 8px 0;\">Informacja dla odbiorcy: Informacje zawarte w niniejszym email-u oraz załącznikach do niego mają charakter poufny, są przeznaczone wyłącznie dla wskazanych adresatów. Jeśli nie są Państwo adresatem tego email-a, prosimy niezwłocznie o jego skasowanie oraz poinformowanie nadawcy. Wykonywanie kopii, ujawnienie, dystrybucja lub używanie niniejszego email-a do innych celów jest zabronione. Spółka Grupa Familijna Sp. z o.o. nie ponosi żadnej odpowiedzialności za zmiany email-a dokonane po jego wysłaniu.</p>\n            <p style=\"margin:0;\">Administratorem danych osobowych jest Grupa Familijna sp. z o.o. z siedzibą w Kuźnicy Czeszyckiej. Dane osobowe zawarte w korespondencji mailowej są przetwarzane w celu odpowiadania na pytania, dokonywania ustaleń, zawierania i realizacji umów z kontrahentami, rozpoznawania reklamacji, jak również ustalenia, dochodzenia i obrony roszczeń. Mają Państwo w szczególności prawo dostępu do swoich danych osobowych, żądania ich usunięcia i wniesienia sprzeciwu wobec przetwarzania danych. Szczegóły dotyczące przetwarzania danych osobowych i przysługujących praw znajdują się w <a href=\"https://www.grupafamilijna.pl/pl/polityka-prywatnosci\" style=\"color:#0645AD;text-decoration:underline;\">Polityce prywatności</a>.</p>\n        </td>\n    </tr>\n</table>\n";

  html = replaceAllSafe(html, "%%DisplayName%%", displayName);
  html = replaceAllSafe(html, "%%Email%%", email);
  html = replaceAllSafe(html, "%%Title%%", title);
  html = replaceAllSafe(html, "%%PhoneNumber%%", phoneNumber);
  html = replaceAllSafe(html, "%%MobileNumber%%", mobileNumber);
  html = replaceAllSafe(html, "%%PhoneHtml%%", phoneHtml);
  html = replaceAllSafe(html, "%%Department%%", department);
  html = replaceAllSafe(html, "%%OfficeLocation%%", officeLocation);
  html = replaceAllSafe(html, "%%CompanyName%%", companyName);

  return '<div data-familijna-signature="1">' + html + '</div>';
}


function getBodyHtml() {
  return new Promise(function (resolve, reject) {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          const msg = result.error && result.error.message ? result.error.message : "Nieznany błąd body.getAsync.";
          reject(new Error(msg));
        }
      }
    );
  });
}

function setBodyHtml(html) {
  return new Promise(function (resolve, reject) {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          const msg = result.error && result.error.message ? result.error.message : "Nieznany błąd body.setAsync.";
          reject(new Error(msg));
        }
      }
    );
  });
}

async function onNewMessageComposeHandler(event) {
  try {
    const enabled = roamingGet(STORAGE_AUTO_KEY);

    if (enabled !== true && enabled !== "true") {
      eventDone(event);
      return;
    }

    const bodyHtml = await getBodyHtml();

    if (bodyHtml.indexOf(SIGNATURE_MARKER) !== -1) {
      eventDone(event);
      return;
    }

    let user = null;
    const profileJson = roamingGet(STORAGE_PROFILE_KEY);
    if (profileJson) {
      try {
        user = typeof profileJson === "string" ? JSON.parse(profileJson) : profileJson;
      } catch (e) {
        user = null;
      }
    }

    const signatureHtml = buildSignatureHtml(user);
    await setBodyHtml(bodyHtml + "<br><br>" + signatureHtml);

    eventDone(event);
  } catch (e) {
    eventDone(event);
  }
}

Office.actions.associate(
  "onNewMessageComposeHandler",
  onNewMessageComposeHandler
);
