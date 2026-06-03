/* Stopka Familijna v3.1 - panel FIX 3.1.0.1 */
var GF_VERSION = typeof GF_VERSION !== "undefined" ? GF_VERSION : "3.1.0.0";
var GF_AUTO_KEY = typeof GF_AUTO_KEY !== "undefined" ? GF_AUTO_KEY : "autoSignatureEnabled";
var GF_PROFILE_KEY = typeof GF_PROFILE_KEY !== "undefined" ? GF_PROFILE_KEY : "signatureUserProfile";
var GF_MARKER = typeof GF_MARKER !== "undefined" ? GF_MARKER : 'data-familijna-signature="1"';
var GF_GRAPH_ME_URL = typeof GF_GRAPH_ME_URL !== "undefined" ? GF_GRAPH_ME_URL : "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,department,officeLocation,companyName";

/* Stopka Familijna v3.1 - panel */
const AUTH_URL = "https://rafalsieradzki.github.io/outlookwebstopka/auth.html";
let authDialog = null;

Office.onReady(function () {
  const insertButton = document.getElementById("insertSignature");
  const diagnosticButton = document.getElementById("runDiagnostics");
  const checkbox = document.getElementById("autoSignatureEnabled");

  if (insertButton) insertButton.onclick = insertSignatureManual;
  if (diagnosticButton) diagnosticButton.onclick = runDiagnostics;

  if (checkbox) {
    const enabled = roamingGet(GF_AUTO_KEY);
    checkbox.checked = enabled === true || enabled === "true";
    checkbox.onchange = onAutoSignatureChanged;
  }

  logDiag("Panel gotowy. Wersja 3.1.0.0.");
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

function logDiag(message, isError) {
  const out = document.getElementById("diagnosticOutput");
  if (!out) return;
  const time = new Date().toLocaleTimeString();
  const current = out.textContent && out.textContent !== "Diagnostyka pojawi się tutaj." ? out.textContent + "\n" : "";
  out.textContent = current + "[" + time + "] " + message;
  out.className = isError ? "error" : "";
}

function clearDiag() {
  const out = document.getElementById("diagnosticOutput");
  if (out) {
    out.textContent = "";
    out.className = "";
  }
}

function setBusy(id, isBusy, busyText, normalText) {
  const button = document.getElementById(id);
  if (!button) return;
  button.disabled = isBusy;
  if (isBusy && busyText) button.textContent = busyText;
  if (!isBusy && normalText) button.textContent = normalText;
}

function roamingSettings() {
  return Office.context && Office.context.roamingSettings ? Office.context.roamingSettings : null;
}

function roamingGet(key) {
  try {
    const settings = roamingSettings();
    return settings ? settings.get(key) : null;
  } catch (e) {
    return null;
  }
}

function roamingSave() {
  return new Promise(function (resolve, reject) {
    const settings = roamingSettings();
    if (!settings) {
      reject(new Error("Brak Office.context.roamingSettings."));
      return;
    }
    settings.saveAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error && result.error.message ? result.error.message : "Nieznany błąd saveAsync."));
      }
    });
  });
}

async function roamingSetAndSave(key, value) {
  const settings = roamingSettings();
  if (!settings) throw new Error("Brak Office.context.roamingSettings.");
  settings.set(key, value);
  await roamingSave();
  return settings.get(key);
}

function getAccessTokenWithDialog() {
  return new Promise(function (resolve, reject) {
    Office.context.ui.displayDialogAsync(
      AUTH_URL,
      { height: 65, width: 45, displayInIframe: false, promptBeforeOpen: false },
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Nie udało się otworzyć okna logowania: " + asyncResult.error.message));
          return;
        }

        authDialog = asyncResult.value;
        authDialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          try {
            const message = JSON.parse(arg.message);
            if (message.status === "success" && message.accessToken) {
              authDialog.close();
              authDialog = null;
              resolve(message.accessToken);
              return;
            }
            authDialog.close();
            authDialog = null;
            reject(new Error(message.message || "Błąd logowania Microsoft 365."));
          } catch (e) {
            if (authDialog) authDialog.close();
            authDialog = null;
            reject(e);
          }
        });
      }
    );
  });
}

async function getGraphUser(accessToken) {
  const response = await fetch(GF_GRAPH_ME_URL, {
    headers: { Authorization: "Bearer " + accessToken }
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error("Graph API: " + response.status + " " + errorText);
  }
  return await response.json();
}

async function ensureGraphProfile() {
  logDiag("Otwieram logowanie Microsoft 365...");
  const accessToken = await getAccessTokenWithDialog();
  logDiag("Token pobrany. Pobieram profil z Microsoft Graph...");
  const user = await getGraphUser(accessToken);
  logDiag("Graph OK: " + (user.mail || user.userPrincipalName || "brak maila") + ", " + (user.displayName || "brak nazwy"));
  await roamingSetAndSave(GF_PROFILE_KEY, JSON.stringify(user));
  logDiag("Profil zapisany w roamingSettings.");
  return user;
}

async function onAutoSignatureChanged() {
  const checkbox = document.getElementById("autoSignatureEnabled");
  const enabled = checkbox && checkbox.checked === true;
  clearDiag();

  try {
    logDiag("Zapisuję autoSignatureEnabled=" + enabled + "...");
    const saved = await roamingSetAndSave(GF_AUTO_KEY, enabled);
    logDiag("Zapis checkboxa OK: autoSignatureEnabled=" + saved);

    if (enabled) {
      await ensureGraphProfile();
      setStatus("Automatyczna stopka jest włączona.", false, true);
    } else {
      setStatus("Automatyczna stopka jest wyłączona.", false, true);
    }
  } catch (e) {
    if (checkbox) checkbox.checked = false;
    try { await roamingSetAndSave(GF_AUTO_KEY, false); } catch (_) {}
    const msg = e && e.message ? e.message : String(e);
    logDiag("Błąd: " + msg, true);
    setStatus("Nie udało się zapisać ustawienia. Checkbox cofnięty.", true, false);
  }
}

async function insertSignatureManual() {
  clearDiag();
  setBusy("insertSignature", true, "Wstawiam...", "Wstaw stopkę teraz");
  setStatus("Przygotowuję stopkę...", false, false);

  try {
    const user = await ensureGraphProfile();
    const html = gfBuildSignatureHtml(user, Office.context.mailbox.userProfile || {});
    logDiag("HTML stopki przygotowany. Wstawiam do wiadomości...");

    Office.context.mailbox.item.body.setSelectedDataAsync(
      "<br><br>" + html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        setBusy("insertSignature", false, null, "Wstaw stopkę teraz");
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          logDiag("Stopka została wstawiona ręcznie.");
          setStatus("Stopka została wstawiona.", false, true);
        } else {
          const msg = result.error && result.error.message ? result.error.message : "Nieznany błąd Outlook API.";
          logDiag("Błąd wstawiania: " + msg, true);
          setStatus("Nie udało się wstawić stopki.", true, false);
        }
      }
    );
  } catch (e) {
    setBusy("insertSignature", false, null, "Wstaw stopkę teraz");
    const msg = e && e.message ? e.message : String(e);
    logDiag("Błąd: " + msg, true);
    setStatus("Nie udało się przygotować stopki.", true, false);
  }
}

async function runDiagnostics() {
  clearDiag();
  setBusy("runDiagnostics", true, "Sprawdzam...", "Diagnostyka");
  setStatus("Uruchomiono diagnostykę...", false, false);

  try {
    logDiag("Wersja dodatku: 3.1.0.0");
    logDiag("Office.js: działa");

    const mailboxProfile = Office.context && Office.context.mailbox ? Office.context.mailbox.userProfile : null;
    if (mailboxProfile) {
      logDiag("Outlook profil: " + (mailboxProfile.displayName || "brak nazwy") + " / " + (mailboxProfile.emailAddress || "brak email"));
    } else {
      logDiag("Outlook profil: brak danych", true);
    }

    const settings = roamingSettings();
    logDiag("roamingSettings: " + (settings ? "dostępne" : "niedostępne"), !settings);
    const autoValue = roamingGet(GF_AUTO_KEY);
    const profileValue = roamingGet(GF_PROFILE_KEY);
    logDiag("autoSignatureEnabled: " + autoValue);

    const savedProfile = gfParseProfile(profileValue);
    if (savedProfile) {
      logDiag("Zapisany profil: " + (savedProfile.displayName || "brak nazwy") + " / " + (savedProfile.mail || savedProfile.userPrincipalName || "brak email"));
    } else {
      logDiag("Zapisany profil: brak lub nieczytelny");
    }

    const user = await ensureGraphProfile();
    logDiag("Test Graph zakończony OK.");
    logDiag("Dane do stopki: imię/nazwa=" + (user.displayName || "brak") + ", stanowisko=" + (user.jobTitle || "brak") + ", telefon=" + (gfFirstBusinessPhone(user) || "brak") + ", komórka=" + (user.mobilePhone || "brak"));

    setStatus("Diagnostyka zakończona.", false, true);
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    logDiag("Błąd diagnostyki: " + msg, true);
    setStatus("Diagnostyka wykazała błąd.", true, false);
  } finally {
    setBusy("runDiagnostics", false, null, "Diagnostyka");
  }
}
