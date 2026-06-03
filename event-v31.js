/* Stopka Familijna v3.1 - event-based activation, bundled for classic/new Outlook */
(function () {
  var GF_VERSION = "3.1.0.0";
  var GF_AUTO_KEY = "autoSignatureEnabled";
  var GF_PROFILE_KEY = "signatureUserProfile";
  var GF_EVENT_LOG_KEY = "signatureEventLastLog";
  var GF_MARKER = 'data-familijna-signature="1"';

  function complete(event) {
    try {
      if (event && typeof event.completed === "function") event.completed();
    } catch (_) {}
  }

  function text(value) {
    if (value === null || value === undefined) return "";
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;");
  }

  function firstBusinessPhone(user) {
    return user && user.businessPhones && user.businessPhones.length > 0 ? (user.businessPhones[0] || "") : "";
  }

  function buildPhoneHtml(phoneNumber, mobileNumber) {
    var parts = [];
    if (phoneNumber) parts.push('<span style="color:#DF292F;">tel.</span> ' + text(phoneNumber));
    if (mobileNumber) parts.push('<span style="color:#DF292F;">kom.</span> ' + text(mobileNumber));
    return parts.join(" ");
  }

  function buildSignatureHtml(user, officeProfile) {
    user = user || {};
    officeProfile = officeProfile || {};

    var displayName = user.displayName || officeProfile.displayName || "";
    var email = user.mail || user.userPrincipalName || officeProfile.emailAddress || "";
    var title = user.jobTitle || "";
    var phoneNumber = firstBusinessPhone(user);
    var mobileNumber = user.mobilePhone || "";
    var phoneHtml = buildPhoneHtml(phoneNumber, mobileNumber);

    return '<div ' + GF_MARKER + '>' +
      '<table cellpadding="0" cellspacing="0" border="0" style="max-width:520px;font-family:Calibri,Arial;border-collapse:collapse;line-height:120%;mso-line-height-rule:exactly;">' +
      '<tr>' +
      '<td style="margin:auto;width:220px;" align="center">' +
      '<img src="https://www.familijna.pl/uploads/drive/familijna_logotyp.png" width="80%" alt="GRUPA FAMILIJNA" />' +
      '</td>' +
      '<td style="font-size:9pt;line-height:115%;mso-line-height-rule:exactly;color:#595959;border-left:3px solid #DF292F;padding-left:15px;">' +
      '<span style="font-size:14pt;color:#DF292F;">' + text(displayName) + '</span><br />' +
      '<span>' + text(title) + '</span><br /><br />' +
      '<a href="https://familijna.pl" style="color:#595959;text-decoration:none;"><span style="color:#DF292F;">www.</span>familijna.pl</a> ' +
      '<span style="color:#DF292F;">email:</span> ' +
      '<a href="mailto:' + text(email) + '" style="color:#595959;text-decoration:none;">' + text(email) + '</a><br />' +
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
      '<tr><td style="font-size:7pt;line-height:110%;mso-line-height-rule:exactly;color:#595959;padding:0;margin:0;padding:0;line-height:120%;mso-line-height-rule:exactly;">' +
      '<p style="margin:0;padding:0;line-height:120%;mso-line-height-rule:exactly;"><span style="color:#DF292F;">GRUPA FAMILIJNA</span> Spółka z ograniczoną odpowiedzialnością, Kuźnica Czeszycka 11, 56-320 Krośnice, tel. 71 384 56 13</p>' +
      '<p style="margin:0;padding:0;line-height:120%;mso-line-height-rule:exactly;">NIP: 9161351695, REGON: 020182505, BDO: 000084673.</p>' +
      '<p style="margin:0;padding:0;line-height:120%;mso-line-height-rule:exactly;">Informacja dla odbiorcy: Informacje zawarte w niniejszym email-u oraz załącznikach do niego mają charakter poufny, są przeznaczone wyłącznie dla wskazanych adresatów. Jeśli nie są Państwo adresatem tego email-a, prosimy niezwłocznie o jego skasowanie oraz poinformowanie nadawcy. Wykonywanie kopii, ujawnienie, dystrybucja lub używanie niniejszego email-a do innych celów jest zabronione. Spółka Grupa Familijna Sp. z o.o. nie ponosi żadnej odpowiedzialności za zmiany email-a dokonane po jego wysłaniu.</p>' +
      '<p style="margin:0;padding:0;line-height:120%;mso-line-height-rule:exactly;">Administratorem danych osobowych jest Grupa Familijna sp. z o.o. z siedzibą w Kuźnicy Czeszyckiej. Dane osobowe zawarte w korespondencji mailowej są przetwarzane w celu odpowiadania na pytania, dokonywania ustaleń, zawierania i realizacji umów z kontrahentami, rozpoznawania reklamacji, jak również ustalenia, dochodzenia i obrony roszczeń. Mają Państwo w szczególności prawo dostępu do swoich danych osobowych, żądania ich usunięcia i wniesienia sprzeciwu wobec przetwarzania danych. Szczegóły dotyczące przetwarzania danych osobowych i przysługujących praw znajdują się w <a href="https://www.grupafamilijna.pl/pl/polityka-prywatnosci" style="color:#0645AD;text-decoration:underline;">Polityce prywatności</a>.</p>' +
      '</td></tr></table>' +
      '</div>';
  }

  function parseProfile(value) {
    if (!value) return null;
    if (typeof value === "object") return value;
    try { return JSON.parse(value); } catch (_) { return null; }
  }

  function settings() {
    try { return Office.context && Office.context.roamingSettings ? Office.context.roamingSettings : null; } catch (_) { return null; }
  }

  function getRoaming(key) {
    var s = settings();
    try { return s ? s.get(key) : null; } catch (_) { return null; }
  }

  function writeEventLog(message) {
    var s = settings();
    if (!s) return;
    try {
      s.set(GF_EVENT_LOG_KEY, new Date().toISOString() + " " + message);
      s.saveAsync(function () {});
    } catch (_) {}
  }

  function getBodyHtml(callback) {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) callback(String(result.value || ""));
        else callback("");
      });
    } catch (_) { callback(""); }
  }

  function insertHtml(html, done) {
    var body = Office.context.mailbox.item.body;
    if (body && typeof body.setSignatureAsync === "function") {
      body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) { done("setSignatureAsync OK"); return; }
        fallbackInsert(html, done);
      });
      return;
    }
    fallbackInsert(html, done);
  }

  function fallbackInsert(html, done) {
    var body = Office.context.mailbox.item.body;
    if (body && typeof body.prependAsync === "function") {
      body.prependAsync("<br><br>" + html, { coercionType: Office.CoercionType.Html }, function (result) {
        done(result && result.status === Office.AsyncResultStatus.Succeeded ? "prependAsync OK" : "prependAsync FAIL");
      });
      return;
    }
    body.setSelectedDataAsync("<br><br>" + html, { coercionType: Office.CoercionType.Html }, function (result) {
      done(result && result.status === Office.AsyncResultStatus.Succeeded ? "setSelectedDataAsync OK" : "setSelectedDataAsync FAIL");
    });
  }

  function handler(event) {
    writeEventLog("event start");
    try {
      var enabled = getRoaming(GF_AUTO_KEY);
      if (!(enabled === true || enabled === "true")) {
        writeEventLog("auto disabled: " + enabled);
        complete(event);
        return;
      }

      var profile = parseProfile(getRoaming(GF_PROFILE_KEY));
      if (!profile) {
        writeEventLog("profile missing");
        complete(event);
        return;
      }

      getBodyHtml(function (bodyHtml) {
        if (bodyHtml.indexOf(GF_MARKER) >= 0) {
          writeEventLog("signature already exists");
          complete(event);
          return;
        }

        var officeProfile = (Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) ? Office.context.mailbox.userProfile : {};
        var html = buildSignatureHtml(profile, officeProfile);
        insertHtml(html, function (insertStatus) {
          writeEventLog("signature inserted: " + insertStatus);
          complete(event);
        });
      });
    } catch (e) {
      writeEventLog("error: " + (e && e.message ? e.message : String(e)));
      complete(event);
    }
  }

  function register() {
    try {
      if (Office && Office.actions && typeof Office.actions.associate === "function") {
        Office.actions.associate("onNewMessageComposeHandlerV31", handler);
        writeEventLog("handler associated");
      }
    } catch (e) {
      writeEventLog("associate error: " + (e && e.message ? e.message : String(e)));
    }
  }

  register();
  try { Office.onReady(register); } catch (_) {}
  try { Office.initialize = register; } catch (_) {}
})();
