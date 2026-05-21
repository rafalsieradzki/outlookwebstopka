// Stopka Familijna - taskpane.js
// Poprawna wersja: uwierzytelnianie przez Office Dialog API.
// Panel otwiera auth.html, auth.html zwraca access token przez Office.context.ui.messageParent().

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

        authDialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          function (arg) {
            // 12006 bywa zwracane po zamknięciu okna; jeśli token już wrócił, ignorujemy.
            if (arg && arg.error) {
              console.warn("DialogEventReceived:", arg);
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

  let html = "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"max-width:600px;font-family:Calibri, Arial;\">\n  <tr>\n    <td style=\"font-size:9pt;line-height:140%;color:#595959;border-left:3px solid #DF292F;padding-left:15px;\">\n      <span style=\"font-size:14pt;color:#DF292F;\">%%DisplayName%%</span><br />\n      <span>%%Title%%</span><br /><br />\n      <a href=\"mailto:%%Email%%\" style=\"color:#595959;text-decoration:none;\">%%Email%%</a><br />\n      <span style=\"color:#DF292F;\">tel.</span> %%PhoneNumber%%\n      <span style=\"color:#DF292F;\">kom.</span> %%MobileNumber%%\n    </td>\n  </tr>\n</table>";

  html = replaceAllSafe(html, "%%DisplayName%%", displayName);
  html = replaceAllSafe(html, "%%Email%%", email);
  html = replaceAllSafe(html, "%%Title%%", title);
  html = replaceAllSafe(html, "%%PhoneNumber%%", phoneNumber);
  html = replaceAllSafe(html, "%%MobileNumber%%", mobileNumber);
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
