// Stopka Familijna - taskpane.js
// Upiększona wersja HTML stopki + logowanie przez Office Dialog API + Microsoft Graph.

const AUTH_URL = "https://rafalsieradzki.github.io/outlookwebstopka/auth.html";
const GRAPH_ME_URL =
  "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,department,officeLocation,companyName";

let authDialog = null;

Office.onReady(function () {
  const button = document.getElementById("insertSignature");
  if (button) {
    button.onclick = insertSignature;
  }

  setStatus("Dodatek gotowy.", false, true);
});

function setStatus(message, isError, isOk) {
  const status = document.getElementById("status");
  if (!status) return;

  status.textContent = message || "";
  status.className = "";

  if (isError) status.className = "error";
  if (isOk) status.className = "ok";
}

function setButtonBusy(isBusy) {
  const button = document.getElementById("insertSignature");
  if (!button) return;

  button.disabled = isBusy;
  button.textContent = isBusy ? "Pobieram dane..." : "Wstaw stopkę";
}

function getAccessTokenWithDialog() {
  return new Promise(function (resolve, reject) {
    setStatus("Otwieram logowanie Microsoft 365...", false, false);

    Office.context.ui.displayDialogAsync(
      AUTH_URL,
      {
        height: 65,
        width: 45,
        displayInIframe: false,
        promptBeforeOpen: false
      },
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Nie udało się otworzyć okna logowania: " + asyncResult.error.message));
          return;
        }

        authDialog = asyncResult.value;

        authDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          function (arg) {
            try {
              const message = JSON.parse(arg.message);

              if (message.status === "success" && message.accessToken) {
                authDialog.close();
                authDialog = null;
                resolve(message.accessToken);
                return;
              }

              if (message.status === "error") {
                authDialog.close();
                authDialog = null;
                reject(new Error(message.message || "Błąd logowania."));
                return;
              }

              reject(new Error("Nieznana odpowiedź z okna logowania."));
            } catch (e) {
              if (authDialog) {
                authDialog.close();
                authDialog = null;
              }
              reject(e);
            }
          }
        );
      }
    );
  });
}

async function getGraphUser(accessToken) {
  setStatus("Pobieram dane użytkownika z Microsoft Graph...", false, false);

  const response = await fetch(GRAPH_ME_URL, {
    headers: {
      Authorization: "Bearer " + accessToken
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error("Graph API: " + response.status + " " + errorText);
  }

  return await response.json();
}

function replaceAllSafe(text, token, value) {
  return text.split(token).join(value || "");
}

function firstBusinessPhone(user) {
  if (user.businessPhones && user.businessPhones.length > 0) {
    return user.businessPhones[0] || "";
  }
  return "";
}

function buildPhoneHtml(phoneNumber, mobileNumber) {
  const parts = [];

  if (phoneNumber) {
    parts.push('<span style="color:#DF292F;font-weight:bold;">tel.</span> ' + phoneNumber);
  }

  if (mobileNumber) {
    parts.push('<span style="color:#DF292F;font-weight:bold;">kom.</span> ' + mobileNumber);
  }

  return parts.join(" ");
}

function buildSignatureHtml(user) {
  const officeProfile = Office.context.mailbox.userProfile || {};

  const displayName = user.displayName || officeProfile.displayName || "";
  const email = user.mail || user.userPrincipalName || officeProfile.emailAddress || "";
  const title = user.jobTitle || "";
  const phoneNumber = firstBusinessPhone(user);
  const mobileNumber = user.mobilePhone || "";
  const department = user.department || "";
  const officeLocation = user.officeLocation || "";
  const companyName = user.companyName || "";
  const phoneHtml = buildPhoneHtml(phoneNumber, mobileNumber);

  let html = "\n<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"760\" style=\"width:760px;max-width:760px;font-family:Calibri, Arial, sans-serif;color:#303847;\">\n  <tr>\n    <td width=\"320\" valign=\"top\" style=\"width:320px;padding:0 28px 0 10px;text-align:left;\">\n      <img src=\"https://www.familijna.pl/uploads/drive/familijna_logotyp.png\" width=\"275\" alt=\"Grupa Familijna\" style=\"display:block;border:0;width:275px;height:auto;margin:0 0 18px 0;\" />\n      <table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"border-collapse:collapse;\">\n        <tr>\n          <td style=\"padding-right:18px;\">\n            <img src=\"https://www.familijna.pl/uploads/drive/familijna_dolny_barycz.png\" width=\"100\" alt=\"Familijna Dolny Barycz\" style=\"display:block;border:0;width:100px;height:auto;\" />\n          </td>\n          <td>\n            <img src=\"https://www.familijna.pl/uploads/drive/familijna_cukiernia.png\" width=\"100\" alt=\"Familijna Cukiernia\" style=\"display:block;border:0;width:100px;height:auto;\" />\n          </td>\n        </tr>\n      </table>\n    </td>\n\n    <td width=\"3\" style=\"width:3px;background:#DF292F;font-size:0;line-height:0;\">&nbsp;</td>\n\n    <td width=\"437\" valign=\"top\" style=\"width:437px;padding:0 0 0 30px;font-size:10pt;line-height:145%;color:#303847;\">\n      <div style=\"font-size:17pt;line-height:120%;color:#DF292F;margin:0 0 4px 0;\">%%DisplayName%%</div>\n      <div style=\"font-size:10pt;color:#303847;margin:0 0 22px 0;\">%%Title%%</div>\n\n      <div style=\"font-size:10pt;color:#303847;margin:0 0 2px 0;\">\n        <a href=\"https://www.familijna.pl\" style=\"color:#303847;text-decoration:underline;\"><span style=\"color:#DF292F;\">www.</span>familijna.pl</a>\n        <span style=\"color:#DF292F;font-weight:bold;\"> email:</span>\n        <a href=\"mailto:%%Email%%\" style=\"color:#303847;text-decoration:underline;\">%%Email%%</a>\n      </div>\n\n      <div style=\"font-size:10pt;color:#303847;margin:0 0 22px 0;\">\n        %%PhoneHtml%%\n      </div>\n\n      <table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"border-collapse:collapse;margin-top:8px;\">\n        <tr>\n          <td style=\"padding-right:8px;\"><a href=\"https://www.facebook.com/familijna\"><img src=\"https://www.familijna.pl/uploads/drive/fb.png\" width=\"28\" height=\"28\" alt=\"Facebook\" style=\"display:block;border:0;\" /></a></td>\n          <td style=\"padding-right:8px;\"><a href=\"https://www.instagram.com/familijna/\"><img src=\"https://www.familijna.pl/uploads/drive/ig.png\" width=\"28\" height=\"28\" alt=\"Instagram\" style=\"display:block;border:0;\" /></a></td>\n          <td style=\"padding-right:8px;\"><a href=\"https://m.me/familijna\"><img src=\"https://www.familijna.pl/uploads/drive/ms.png\" width=\"28\" height=\"28\" alt=\"Messenger\" style=\"display:block;border:0;\" /></a></td>\n          <td style=\"padding-right:8px;\"><a href=\"https://goo.gl/maps/kpEMXw6deUcjidot9\"><img src=\"https://www.familijna.pl/uploads/drive/gm.png\" width=\"28\" height=\"28\" alt=\"Google Maps\" style=\"display:block;border:0;\" /></a></td>\n          <td style=\"padding-right:8px;\"><a href=\"https://www.youtube.com/@familijna1631/featured\"><img src=\"https://www.familijna.pl/uploads/drive/yt.png\" width=\"28\" height=\"28\" alt=\"YouTube\" style=\"display:block;border:0;\" /></a></td>\n          <td><a href=\"https://www.linkedin.com/company/familijna\"><img src=\"https://www.familijna.pl/uploads/drive/in.png\" width=\"28\" height=\"28\" alt=\"LinkedIn\" style=\"display:block;border:0;\" /></a></td>\n        </tr>\n      </table>\n    </td>\n  </tr>\n\n  <tr>\n    <td colspan=\"3\" style=\"padding:26px 10px 0 10px;font-size:7.5pt;line-height:135%;color:#303847;\">\n      <p style=\"margin:0 0 8px 0;\">\n        <span style=\"color:#DF292F;\">GRUPA FAMILIJNA</span> Spółka z ograniczoną odpowiedzialnością, Kuźnica Czeszycka 11, 56-320 Krośnice, tel. 71 384 56 13\n      </p>\n      <p style=\"margin:0 0 22px 0;\">\n        NIP: 9161351695, REGON: 020182505, BDO: 000084673.\n      </p>\n\n      <p style=\"margin:0 0 8px 0;\">\n        Informacja dla odbiorcy: Informacje zawarte w niniejszym email-u oraz załącznikach do niego mają charakter poufny, są przeznaczone wyłącznie dla wskazanych adresatów. Jeśli nie są Państwo adresatem tego email-a, prosimy niezwłocznie o jego skasowanie oraz poinformowanie nadawcy. Wykonywanie kopii, ujawnienie, dystrybucja lub używanie niniejszego email-a do innych celów jest zabronione. Spółka Grupa Familijna Sp. z o.o. nie ponosi żadnej odpowiedzialności za zmiany email-a dokonane po jego wysłaniu.\n      </p>\n      <p style=\"margin:0;\">\n        Administratorem danych osobowych jest Grupa Familijna sp. z o.o. z siedzibą w Kuźnicy Czeszyckiej. Dane osobowe zawarte w korespondencji mailowej są przetwarzane w celu odpowiadania na pytania, dokonywania ustaleń, zawierania i realizacji umów z kontrahentami, rozpoznawania reklamacji, jak również ustalenia, dochodzenia i obrony roszczeń. Mają Państwo w szczególności prawo dostępu do swoich danych osobowych, żądania ich usunięcia i wniesienia sprzeciwu wobec przetwarzania danych. Szczegóły dotyczące przetwarzania danych osobowych i przysługujących praw znajdują się w <a href=\"https://www.grupafamilijna.pl/pl/polityka-prywatnosci\" style=\"color:#0645AD;text-decoration:underline;\">Polityce prywatności</a>.\n      </p>\n    </td>\n  </tr>\n</table>\n";

  html = replaceAllSafe(html, "%%DisplayName%%", displayName);
  html = replaceAllSafe(html, "%%Email%%", email);
  html = replaceAllSafe(html, "%%Title%%", title);
  html = replaceAllSafe(html, "%%PhoneNumber%%", phoneNumber);
  html = replaceAllSafe(html, "%%MobileNumber%%", mobileNumber);
  html = replaceAllSafe(html, "%%PhoneHtml%%", phoneHtml);
  html = replaceAllSafe(html, "%%Department%%", department);
  html = replaceAllSafe(html, "%%OfficeLocation%%", officeLocation);
  html = replaceAllSafe(html, "%%CompanyName%%", companyName);

  return "<br><br>" + html;
}

async function insertSignature() {
  setButtonBusy(true);
  setStatus("Start...", false, false);

  try {
    const accessToken = await getAccessTokenWithDialog();
    const user = await getGraphUser(accessToken);
    const html = buildSignatureHtml(user);

    setStatus("Wstawiam stopkę do wiadomości...", false, false);

    Office.context.mailbox.item.body.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        setButtonBusy(false);

        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus("Stopka została wstawiona.", false, true);
          console.log("Stopka wstawiona.");
        } else {
          const msg = result.error && result.error.message
            ? result.error.message
            : "Nieznany błąd Outlook API.";

          setStatus("Nie udało się wstawić stopki: " + msg, true, false);
          console.error("Nie udało się wstawić stopki:", result.error);
        }
      }
    );
  } catch (e) {
    setButtonBusy(false);
    const message = e && e.message ? e.message : String(e);

    setStatus("Nie udało się pobrać danych użytkownika:\n" + message, true, false);
    console.error("Błąd dodatku:", e);
  }
}
