/* Stopka Familijna v3.1 - wspolne stale, szablon HTML i generator */
var GF_VERSION = "3.1.0.0";
var GF_AUTO_KEY = "autoSignatureEnabled";
var GF_PROFILE_KEY = "signatureUserProfile";
var GF_MARKER = 'data-familijna-signature="1"';
var GF_GRAPH_ME_URL = "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,businessPhones,mobilePhone,department,officeLocation,companyName";
var GF_TEMPLATE_URL = "https://rafalsieradzki.github.io/outlookwebstopka/signature-template.html?v=3.1.2.0";
var GF_TEMPLATE_CACHE = null;

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

function gfBuildPhoneLine(phoneNumber) {
  if (!phoneNumber) return "";
  return gfText(phoneNumber);
}

function gfBuildMobileLine(mobileNumber) {
  if (!mobileNumber) return "";
  return '<span style="color:#DF292F;">kom.</span> ' + gfText(mobileNumber);
}

function gfBuildPhoneHtml(phoneNumber, mobileNumber) {
  var parts = [];
  var phoneLine = gfBuildPhoneLine(phoneNumber);
  var mobileLine = gfBuildMobileLine(mobileNumber);
  if (phoneLine) parts.push(phoneLine);
  if (mobileLine) parts.push(mobileLine);
  return parts.join(" ");
}

function gfReplaceAll(text, search, replacement) {
  return String(text).split(search).join(replacement === null || replacement === undefined ? "" : String(replacement));
}

async function gfLoadSignatureTemplate() {
  if (GF_TEMPLATE_CACHE) return GF_TEMPLATE_CACHE;
  var response = await fetch(GF_TEMPLATE_URL, { cache: "no-store" });
  if (!response.ok) throw new Error("Nie udało się pobrać szablonu stopki: HTTP " + response.status);
  GF_TEMPLATE_CACHE = await response.text();
  return GF_TEMPLATE_CACHE;
}

async function gfBuildSignatureHtml(user, officeProfile) {
  user = user || {};
  officeProfile = officeProfile || {};

  var displayName = user.displayName || officeProfile.displayName || "";
  var email = user.mail || user.userPrincipalName || officeProfile.emailAddress || "";
  var title = user.jobTitle || "";
  var phoneNumber = gfFirstBusinessPhone(user);
  var mobileNumber = user.mobilePhone || "";
  var phoneLine = gfBuildPhoneLine(phoneNumber);
  var mobileLine = gfBuildMobileLine(mobileNumber);
  var phoneHtml = gfBuildPhoneHtml(phoneNumber, mobileNumber);

  var html = await gfLoadSignatureTemplate();
  html = gfReplaceAll(html, "{{DISPLAY_NAME}}", gfText(displayName));
  html = gfReplaceAll(html, "{{JOB_TITLE}}", gfText(title));
  html = gfReplaceAll(html, "{{EMAIL}}", gfText(email));
  html = gfReplaceAll(html, "{{PHONE}}", gfText(phoneNumber));
  html = gfReplaceAll(html, "{{MOBILEPHONE}}", gfText(mobileNumber));
  html = gfReplaceAll(html, "{{PHONE_LINE}}", phoneLine);
  html = gfReplaceAll(html, "{{MOBILE_LINE}}", mobileLine);
  html = gfReplaceAll(html, "{{PHONE_HTML}}", phoneHtml);
  return html;
}

function gfParseProfile(value) {
  if (!value) return null;
  if (typeof value === "object") return value;
  try { return JSON.parse(value); } catch (e) { return null; }
}
