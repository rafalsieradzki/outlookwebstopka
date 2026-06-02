/* Stopka Familijna v3.1 - wspolne stale i generator HTML */
const GF_VERSION = "3.1.0.0";
const GF_AUTO_KEY = "autoSignatureEnabled";
const GF_PROFILE_KEY = "signatureUserProfile";
const GF_MARKER = 'data-familijna-signature="1"';
const GF_GRAPH_ME_URL = "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,department,officeLocation,companyName";

function gfText(value) {
  if (value === null || value === undefined) return "";
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;");
}

function gfFirstBusinessPhone(user) {
  return user && user.businessPhones && user.businessPhones.length > 0 ? (user.businessPhones[0] || "") : "";
}

function gfBuildPhoneHtml(phoneNumber, mobileNumber) {
  const parts = [];
  if (phoneNumber) parts.push('<span style="color:#DF292F;">tel.</span> ' + gfText(phoneNumber));
  if (mobileNumber) parts.push('<span style="color:#DF292F;">kom.</span> ' + gfText(mobileNumber));
  return parts.join(" ");
}

function gfBuildSignatureHtml(user, officeProfile) {
  user = user || {};
  officeProfile = officeProfile || {};

  const displayName = user.displayName || officeProfile.displayName || "";
  const email = user.mail || user.userPrincipalName || officeProfile.emailAddress || "";
  const title = user.jobTitle || "";
  const phoneNumber = gfFirstBusinessPhone(user);
  const mobileNumber = user.mobilePhone || "";
  const phoneHtml = gfBuildPhoneHtml(phoneNumber, mobileNumber);

  return '<div ' + GF_MARKER + '>' +
    '<table cellpadding="0" cellspacing="0" border="0" style="max-width:520px;font-family:Calibri, Arial;">' +
    '<tr>' +
    '<td style="margin:auto;width:220px;" align="center">' +
    '<img src="https://www.familijna.pl/uploads/drive/familijna_logotyp.png" width="80%" alt="GRUPA FAMILIJNA" />' +
    '</td>' +
    '<td style="font-size:9pt;line-height:140%;color:#595959;border-left:3px solid #DF292F;padding-left:15px;">' +
    '<span style="font-size:14pt;color:#DF292F;">' + gfText(displayName) + '</span><br />' +
    '<span>' + gfText(title) + '</span><br /><br />' +
    '<a href="https://familijna.pl" style="color:#595959;text-decoration:none;"><span style="color:#DF292F;">www.</span>familijna.pl</a> ' +
    '<span style="color:#DF292F;">email:</span> ' +
    '<a href="mailto:' + gfText(email) + '" style="color:#595959;text-decoration:none;">' + gfText(email) + '</a><br />' +
    phoneHtml +
    '<div style="padding-top:25px;">' +
    '<a href="https://www.facebook.com/familijna" style="display:inline-block;"><img src="https://www.familijna.pl/uploads/drive/fb.png" height="25" width="25" alt="facebook" style="margin-right:5px;" /></a>&nbsp;' +
    '<a href="https://www.instagram.com/familijna/" style="display:inline-block;"><img src="https://www.familijna.pl/uploads/drive/ig.png" height="25" width="25" alt="instagram" style="margin-right:5px;" /></a>&nbsp;' +
    '<a href="https://m.me/familijna" style="display:inline-block;"><img src="https://www.familijna.pl/uploads/drive/ms.png" height="25" width="25" alt="messenger" style="margin-right:5px;" /></a>&nbsp;' +
    '<a href="https://goo.gl/maps/kpEMXw6deUcjidot9" style="display:inline-block;"><img src="https://www.familijna.pl/uploads/drive/gm.png" height="25" width="25" alt="google maps" style="margin-right:5px;" /></a>&nbsp;' +
    '<a href="https://www.youtube.com/@familijna1631/featured" style="display:inline-block;"><img src="https://www.familijna.pl/uploads/drive/yt.png" height="25" width="25" alt="youtube" style="margin-right:5px;" /></a>&nbsp;' +
    '<a href="https://www.linkedin.com/company/familijna" style="display:inline-block;"><img src="https://www.familijna.pl/uploads/drive/in.png" height="25" width="25" alt="linkedin" style="margin-right:5px;" /></a>&nbsp;' +
    '</div></td></tr></table>' +
    '<table cellpadding="0" cellspacing="0" border="0" width="900" style="width:900px;max-width:900px;font-family:Calibri, Arial;margin-top:6px;">' +
    '<tr><td style="font-size:7pt;line-height:120%;color:#595959;">' +
    '<p style="margin:0 0 8px 0;"><span style="color:#DF292F;">GRUPA FAMILIJNA</span> Spółka z ograniczoną odpowiedzialnością, Kuźnica Czeszycka 11, 56-320 Krośnice, tel. 71 384 56 13</p>' +
    '<p style="margin:0 0 20px 0;">NIP: 9161351695, REGON: 020182505, BDO: 000084673.</p>' +
    '<p style="margin:0 0 8px 0;">Informacja dla odbiorcy: Informacje zawarte w niniejszym email-u oraz załącznikach do niego mają charakter poufny, są przeznaczone wyłącznie dla wskazanych adresatów. Jeśli nie są Państwo adresatem tego email-a, prosimy niezwłocznie o jego skasowanie oraz poinformowanie nadawcy. Wykonywanie kopii, ujawnienie, dystrybucja lub używanie niniejszego email-a do innych celów jest zabronione. Spółka Grupa Familijna Sp. z o.o. nie ponosi żadnej odpowiedzialności za zmiany email-a dokonane po jego wysłaniu.</p>' +
    '<p style="margin:0;">Administratorem danych osobowych jest Grupa Familijna sp. z o.o. z siedzibą w Kuźnicy Czeszyckiej. Dane osobowe zawarte w korespondencji mailowej są przetwarzane w celu odpowiadania na pytania, dokonywania ustaleń, zawierania i realizacji umów z kontrahentami, rozpoznawania reklamacji, jak również ustalenia, dochodzenia i obrony roszczeń. Mają Państwo w szczególności prawo dostępu do swoich danych osobowych, żądania ich usunięcia i wniesienia sprzeciwu wobec przetwarzania danych. Szczegóły dotyczące przetwarzania danych osobowych i przysługujących praw znajdują się w <a href="https://www.grupafamilijna.pl/pl/polityka-prywatnosci" style="color:#0645AD;text-decoration:underline;">Polityce prywatności</a>.</p>' +
    '</td></tr></table>' +
    '</div>';
}

function gfParseProfile(value) {
  if (!value) return null;
  if (typeof value === "object") return value;
  try { return JSON.parse(value); } catch (e) { return null; }
}


/* Stopka Familijna v3.1 - event-based activation */
function gfComplete(event) {
  try {
    if (event && typeof event.completed === "function") event.completed();
  } catch (e) {}
}

function gfGetRoaming(key) {
  try {
    if (!Office.context || !Office.context.roamingSettings) return null;
    return Office.context.roamingSettings.get(key);
  } catch (e) {
    return null;
  }
}

function gfBodyContainsMarker(callback) {
  try {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        callback(false);
        return;
      }
      callback(String(result.value || "").indexOf(GF_MARKER) >= 0);
    });
  } catch (e) {
    callback(false);
  }
}

function gfInsertSignature(html, done) {
  const body = Office.context.mailbox.item.body;

  if (body && typeof body.setSignatureAsync === "function") {
    body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        done();
        return;
      }
      gfFallbackInsert(html, done);
    });
    return;
  }

  gfFallbackInsert(html, done);
}

function gfFallbackInsert(html, done) {
  const body = Office.context.mailbox.item.body;

  if (body && typeof body.prependAsync === "function") {
    body.prependAsync("<br><br>" + html, { coercionType: Office.CoercionType.Html }, function () { done(); });
    return;
  }

  body.setSelectedDataAsync("<br><br>" + html, { coercionType: Office.CoercionType.Html }, function () { done(); });
}

function onNewMessageComposeHandlerV31(event) {
  try {
    const enabled = gfGetRoaming(GF_AUTO_KEY);
    if (!(enabled === true || enabled === "true")) {
      gfComplete(event);
      return;
    }

    const profile = gfParseProfile(gfGetRoaming(GF_PROFILE_KEY));
    if (!profile) {
      gfComplete(event);
      return;
    }

    gfBodyContainsMarker(function (exists) {
      if (exists) {
        gfComplete(event);
        return;
      }

      const html = gfBuildSignatureHtml(profile, Office.context.mailbox.userProfile || {});
      gfInsertSignature(html, function () { gfComplete(event); });
    });
  } catch (e) {
    gfComplete(event);
  }
}

try {
  Office.actions.associate("onNewMessageComposeHandlerV31", onNewMessageComposeHandlerV31);
} catch (e) {}
