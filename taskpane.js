// Stopka Familijna - taskpane.js
// Poprawka zapisu ustawień: bezpieczny wrapper OfficeRuntime.storage + fallback localStorage.

const AUTH_URL = "https://rafalsieradzki.github.io/outlookwebstopka/auth.html";
const GRAPH_ME_URL = "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,department,officeLocation,companyName";
const STORAGE_AUTO_KEY = "autoSignatureEnabled";
const STORAGE_PROFILE_KEY = "signatureUserProfile";

let authDialog = null;

Office.onReady(async function () {
  const button = document.getElementById("insertSignature");
  if (button) button.onclick = insertSignatureManual;

  const debugButton = document.getElementById("debugRoamingSettings");
  if (debugButton) debugButton.onclick = debugRoamingSettings;

  const checkbox = document.getElementById("autoSignatureEnabled");
  if (checkbox) {
    const enabled = await storageGet(STORAGE_AUTO_KEY);
    checkbox.checked = enabled === "true";
    checkbox.onchange = onAutoSignatureChanged;
  }

  setStatus("Dodatek gotowy 3.0.0.1 DEBUG.", false, true);
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
  button.textContent = isBusy ? "Pobieram dane..." : "Wstaw stopkę teraz";
}


async function storageGet(key) {
  try {
    if (window.OfficeRuntime && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      return await OfficeRuntime.storage.getItem(key);
    }
  } catch (e) {
    console.warn("OfficeRuntime.storage.getItem failed:", e);
  }

  try {
    return window.localStorage ? window.localStorage.getItem(key) : null;
  } catch (e) {
    console.warn("localStorage.getItem failed:", e);
    return null;
  }
}

async function storageSet(key, value) {
  const stringValue = value == null ? "" : String(value);

  try {
    if (window.OfficeRuntime && OfficeRuntime.storage && OfficeRuntime.storage.setItem) {
      await OfficeRuntime.storage.setItem(key, stringValue);
      return;
    }
  } catch (e) {
    console.warn("OfficeRuntime.storage.setItem failed:", e);
  }

  try {
    if (window.localStorage) {
      window.localStorage.setItem(key, stringValue);
      return;
    }
  } catch (e) {
    console.warn("localStorage.setItem failed:", e);
  }

  throw new Error("Brak dostępnego mechanizmu zapisu ustawień w tym środowisku Outlook.");
}


async function onAutoSignatureChanged() {
  const checkbox = document.getElementById("autoSignatureEnabled");
  const enabled = checkbox && checkbox.checked;

  try {
    await storageSet(STORAGE_AUTO_KEY, enabled ? "true" : "false");

    if (enabled) {
      setStatus("Pobieram dane użytkownika do automatu...", false, false);
      const accessToken = await getAccessTokenWithDialog();
      const user = await getGraphUser(accessToken);
      await storageSet(STORAGE_PROFILE_KEY, JSON.stringify(user));
      setStatus("Automatyczne dodawanie stopki jest włączone.", false, true);
    } else {
      setStatus("Automatyczne dodawanie stopki jest wyłączone.", false, true);
    }
  } catch (e) {
    if (checkbox) checkbox.checked = false;

    try {
      await storageSet(STORAGE_AUTO_KEY, "false");
    } catch (_) {}

    const message = e && e.message ? e.message : String(e);
    setStatus("Nie udało się włączyć automatu:\n" + message, true, false);
  }
}

function getAccessTokenWithDialog() {
  return new Promise(function (resolve, reject) {
    setStatus("Otwieram logowanie Microsoft 365...", false, false);

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
        });
      }
    );
  });
}

async function getGraphUser(accessToken) {
  setStatus("Pobieram dane użytkownika z Microsoft Graph...", false, false);

  const response = await fetch(GRAPH_ME_URL, {
    headers: { Authorization: "Bearer " + accessToken }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error("Graph API: " + response.status + " " + errorText);
  }

  return await response.json();
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


async function insertSignatureManual() {
  setButtonBusy(true);
  setStatus("Start...", false, false);

  try {
    const accessToken = await getAccessTokenWithDialog();
    const user = await getGraphUser(accessToken);

    await storageSet(STORAGE_PROFILE_KEY, JSON.stringify(user));

    const html = buildSignatureHtml(user);

    setStatus("Wstawiam stopkę do wiadomości...", false, false);

    Office.context.mailbox.item.body.setSelectedDataAsync(
      "<br><br>" + html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        setButtonBusy(false);

        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus("Stopka została wstawiona.", false, true);
        } else {
          const msg = result.error && result.error.message ? result.error.message : "Nieznany błąd Outlook API.";
          setStatus("Nie udało się wstawić stopki: " + msg, true, false);
        }
      }
    );
  } catch (e) {
    setButtonBusy(false);
    const message = e && e.message ? e.message : String(e);
    setStatus("Nie udało się pobrać danych użytkownika:\n" + message, true, false);
  }
}



async function debugRoamingSettings() {
  try {
    try {
      const status = document.getElementById("status");
      if (status) {
        status.textContent = "Kliknięto Pokaż roamingSettings...";
        status.className = "ok";
      }
    } catch (_) {}
    const settings = Office.context.roamingSettings;

    const autoValue = settings ? settings.get("autoSignatureEnabled") : null;
    const profileValue = settings ? settings.get("signatureUserProfile") : null;

    let officeRuntimeAuto = null;
    let officeRuntimeProfile = null;

    try {
      if (window.OfficeRuntime && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
        officeRuntimeAuto = await OfficeRuntime.storage.getItem("autoSignatureEnabled");
        officeRuntimeProfile = await OfficeRuntime.storage.getItem("signatureUserProfile");
      }
    } catch (e) {
      officeRuntimeAuto = "OfficeRuntime error: " + (e && e.message ? e.message : String(e));
    }

    let localAuto = null;
    let localProfile = null;

    try {
      if (window.localStorage) {
        localAuto = window.localStorage.getItem("autoSignatureEnabled");
        localProfile = window.localStorage.getItem("signatureUserProfile");
      }
    } catch (e) {
      localAuto = "localStorage error: " + (e && e.message ? e.message : String(e));
    }

    alert(
      "ROAMING SETTINGS\\n" +
      "autoSignatureEnabled = " + autoValue + "\\n" +
      "signatureUserProfile = " + (profileValue ? String(profileValue).substring(0, 500) : "null") + "\\n\\n" +
      "OFFICERUNTIME STORAGE\\n" +
      "autoSignatureEnabled = " + officeRuntimeAuto + "\\n" +
      "signatureUserProfile = " + (officeRuntimeProfile ? String(officeRuntimeProfile).substring(0, 500) : "null") + "\\n\\n" +
      "LOCAL STORAGE\\n" +
      "autoSignatureEnabled = " + localAuto + "\\n" +
      "signatureUserProfile = " + (localProfile ? String(localProfile).substring(0, 500) : "null")
    );
  } catch (e) {
    alert("Błąd debugRoamingSettings: " + (e && e.message ? e.message : String(e)));
  }
}



window.debugRoamingSettings = debugRoamingSettings;
