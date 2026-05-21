// Stopka Familijna - taskpane.js
// Wersja z Microsoft Graph, bez alert() i bez dynamicznego ładowania MSAL.
// MSAL jest ładowany statycznie w taskpane.html.

const CLIENT_ID = "4fbbe7eb-2819-4e83-be4d-6a96aa593088";
const REDIRECT_URI = "https://rafalsieradzki.github.io";
const GRAPH_SCOPES = ["User.Read", "User.ReadBasic.All"];

let msalApp = null;

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

async function getMsalApp() {
  if (msalApp) return msalApp;

  if (typeof msal === "undefined") {
    throw new Error("Biblioteka MSAL nie została załadowana.");
  }

  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: REDIRECT_URI
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false
    }
  });

  return msalApp;
}

async function getAccessToken() {
  const app = await getMsalApp();
  let accounts = app.getAllAccounts();

  if (accounts.length === 0) {
    setStatus("Logowanie do Microsoft 365...", false, false);

    const loginResult = await app.loginPopup({
      scopes: GRAPH_SCOPES,
      prompt: "select_account"
    });

    if (loginResult.account) {
      app.setActiveAccount(loginResult.account);
    }

    accounts = app.getAllAccounts();
  }

  const account = app.getActiveAccount() || accounts[0];

  try {
    const silentResult = await app.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: account
    });

    return silentResult.accessToken;
  } catch (e) {
    setStatus("Odnawianie dostępu do Microsoft Graph...", false, false);

    const popupResult = await app.acquireTokenPopup({
      scopes: GRAPH_SCOPES,
      account: account
    });

    return popupResult.accessToken;
  }
}

async function getGraphUser() {
  const token = await getAccessToken();

  setStatus("Pobieram dane użytkownika z Microsoft Graph...", false, false);

  const response = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,department,officeLocation,companyName",
    {
      headers: {
        Authorization: "Bearer " + token
      }
    }
  );

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

  let html = "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"max-width:600px;font-family:Calibri, Arial;\">\n  <tr>\n    <td>\n      <span style=\"font-size:14pt;color:#DF292F;\">%%DisplayName%%</span><br>\n      <span>%%Title%%</span><br><br>\n      <span>email:</span> %%Email%%<br>\n      <span>tel.</span> %%PhoneNumber%% <span>kom.</span> %%MobileNumber%%\n    </td>\n  </tr>\n</table>";

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
    const user = await getGraphUser();
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
