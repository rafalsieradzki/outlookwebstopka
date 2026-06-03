/* Stopka Familijna v3.1 - event-based activation, template-based, bundled for classic/new Outlook */
(function () {
  var GF_VERSION = "3.1.0.0";
  var GF_AUTO_KEY = "autoSignatureEnabled";
  var GF_PROFILE_KEY = "signatureUserProfile";
  var GF_EVENT_LOG_KEY = "signatureEventLastLog";
  var GF_MARKER = 'data-familijna-signature="1"';
  var GF_TEMPLATE_URL = "https://rafalsieradzki.github.io/outlookwebstopka/signature-template.html?v=3.1.2.0";
  var GF_TEMPLATE_CACHE = null;

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

  function replaceAll(source, search, replacement) {
    return String(source).split(search).join(replacement === null || replacement === undefined ? "" : String(replacement));
  }

  function firstBusinessPhone(user) {
    return user && user.businessPhones && user.businessPhones.length > 0 ? (user.businessPhones[0] || "") : "";
  }

  function buildPhoneLine(phoneNumber) {
    if (!phoneNumber) return "";
    return '<span style="color:#DF292F;">tel.</span> ' + text(phoneNumber);
  }

  function buildMobileLine(mobileNumber) {
    if (!mobileNumber) return "";
    return '<span style="color:#DF292F;">kom.</span> ' + text(mobileNumber);
  }

  function buildPhoneHtml(phoneNumber, mobileNumber) {
    var parts = [];
    var phoneLine = buildPhoneLine(phoneNumber);
    var mobileLine = buildMobileLine(mobileNumber);
    if (phoneLine) parts.push(phoneLine);
    if (mobileLine) parts.push(mobileLine);
    return parts.join(" ");
  }

  function loadTemplate(callback) {
    if (GF_TEMPLATE_CACHE) { callback(null, GF_TEMPLATE_CACHE); return; }

    if (typeof fetch === "function") {
      fetch(GF_TEMPLATE_URL, { cache: "no-store" })
        .then(function (response) {
          if (!response.ok) throw new Error("HTTP " + response.status);
          return response.text();
        })
        .then(function (template) {
          GF_TEMPLATE_CACHE = template;
          callback(null, template);
        })
        .catch(function (error) {
          loadTemplateWithXhr(callback, error);
        });
      return;
    }

    loadTemplateWithXhr(callback, null);
  }

  function loadTemplateWithXhr(callback, originalError) {
    try {
      var xhr = new XMLHttpRequest();
      xhr.open("GET", GF_TEMPLATE_URL, true);
      xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status >= 200 && xhr.status < 300) {
          GF_TEMPLATE_CACHE = xhr.responseText;
          callback(null, xhr.responseText);
        } else {
          callback(new Error("Nie udało się pobrać szablonu stopki" + (originalError ? ": " + originalError.message : ": HTTP " + xhr.status)));
        }
      };
      xhr.send();
    } catch (e) {
      callback(e);
    }
  }

  function buildSignatureHtml(user, officeProfile, callback) {
    user = user || {};
    officeProfile = officeProfile || {};

    var displayName = user.displayName || officeProfile.displayName || "";
    var email = user.mail || user.userPrincipalName || officeProfile.emailAddress || "";
    var title = user.jobTitle || "";
    var phoneNumber = firstBusinessPhone(user);
    var mobileNumber = user.mobilePhone || "";
    var phoneLine = buildPhoneLine(phoneNumber);
    var mobileLine = buildMobileLine(mobileNumber);
    var phoneHtml = buildPhoneHtml(phoneNumber, mobileNumber);

    loadTemplate(function (error, template) {
      if (error) { callback(error); return; }
      var html = template;
      html = replaceAll(html, "{{DISPLAY_NAME}}", text(displayName));
      html = replaceAll(html, "{{JOB_TITLE}}", text(title));
      html = replaceAll(html, "{{EMAIL}}", text(email));
      html = replaceAll(html, "{{PHONE}}", text(phoneNumber));
      html = replaceAll(html, "{{MOBILEPHONE}}", text(mobileNumber));
      html = replaceAll(html, "{{PHONE_LINE}}", phoneLine);
      html = replaceAll(html, "{{MOBILE_LINE}}", mobileLine);
      html = replaceAll(html, "{{PHONE_HTML}}", phoneHtml);
      callback(null, html);
    });
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
        buildSignatureHtml(profile, officeProfile, function (error, html) {
          if (error) {
            writeEventLog("template error: " + (error && error.message ? error.message : String(error)));
            complete(event);
            return;
          }
          insertHtml(html, function (insertStatus) {
            writeEventLog("signature inserted: " + insertStatus);
            complete(event);
          });
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
