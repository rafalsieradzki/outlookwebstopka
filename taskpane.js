const CLIENT_ID = "4fbbe7eb-2819-4e83-be4d-6a96aa593088";
const REDIRECT_URI = "https://rafalsieradzki.github.io";
const GRAPH_SCOPES = ["User.Read", "User.ReadBasic.All"];

let msalApp = null;

Office.onReady(function () {
  const button = document.getElementById("insertSignature");
  if (button) button.onclick = insertSignature;
});

function loadScript(src) {
  return new Promise(function (resolve, reject) {
    const script = document.createElement("script");
    script.src = src;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

async function getMsalApp() {
  if (msalApp) return msalApp;

  await loadScript("https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js");

  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: REDIRECT_URI
    },
    cache: {
      cacheLocation: "sessionStorage"
    }
  });

  return msalApp;
}

async function getAccessToken() {
  const app = await getMsalApp();
  let accounts = app.getAllAccounts();

  if (accounts.length === 0) {
    const loginResult = await app.loginPopup({
      scopes: GRAPH_SCOPES,
      prompt: "select_account"
    });
    app.setActiveAccount(loginResult.account);
    accounts = app.getAllAccounts();
  }

  const account = app.getActiveAccount() || accounts[0];

  try {
    const silentResult = await app.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: account
    });
    return silentResult.accessToken;
  } catch {
    const popupResult = await app.acquireTokenPopup({
      scopes: GRAPH_SCOPES,
      account: account
    });
    return popupResult.accessToken;
  }
}

async function getGraphUser() {
  const token = await getAccessToken();

  const response = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,department,officeLocation,companyName",
    {
      headers: {
        Authorization: "Bearer " + token
      }
    }
  );

  if (!response.ok) {
    throw new Error(await response.text());
  }

  return await response.json();
}

function replaceAllSafe(text, token, value) {
  return text.split(token).join(value || "");
}

function firstBusinessPhone(user) {
  return user.businessPhones && user.businessPhones.length > 0
    ? user.businessPhones[0]
    : "";
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

  let html = `
<p>
%%DisplayName%%<br>
%%Title%%<br>
%%Department%%<br>
%%Email%%<br>
%%PhoneNumber%%<br>
%%MobileNumber%%
</p>
`;

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
  try {
    const user = await getGraphUser();
    const html = buildSignatureHtml(user);

    Office.context.mailbox.item.body.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          alert("Błąd: " + result.error.message);
        }
      }
    );
  } catch (e) {
    alert("Nie udało się pobrać danych użytkownika: " + e.message);
  }
}
