// ============================================================
// Merchant Onboarding Document Generator
// Single-file client app. No backend required.
// docx generation runs entirely in the browser via vendored docx.js.
// ============================================================
import * as docxLib from "./vendor/docx-8.2.2.js";

// ---- CONFIG ----
// Paste your Google Apps Script Web App URL here (see google-apps-script.gs).
// Logs ONLY: timestamp, agent mobile number, merchant display name / ID.
const SHEETS_URL = "https://script.google.com/a/macros/phonepe.com/s/AKfycbyOU_MKNPCXOPT7CYAAf2tmqZsQNSVTbP92QY-gw991gyJrEQdA7Tu_c_FlCsyxIn9_/exec"; // e.g. "https://script.google.com/macros/s/XXXXXXXX/exec"

const LS_MERCHANT = "docgen_merchant_v1";   // clears per onboarding
const LS_AGENT    = "docgen_agent_v1";      // permanent - agent mobile only

const PAN_REGEX = /^[A-Z]{5}\d{4}[A-Z]$/;
const AADHAAR_LAST4 = /^\d{4}$/;
const MOBILE_REGEX = /^[6-9]\d{9}$/;
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const TAN_REGEX = /^[A-Z]{4}\d{5}[A-Z]$/;
const GSTIN_REGEX = /^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;

// ---- STATE ----
const INITIAL_PARTNER = (id, designation) => ({
  id, name: "", pan: "", designation: designation || "Partner", dobRaw: "",
  address: "", poa: "AADHAAR", poaNum: "", nationality: "Indian", share: "", isAS: false,
});

const state = {
  // Step 0 - Agent + Merchant identification (the "only 2 datapoints" logged to Sheet)
  agentMobile: "",
  merchantStatus: "new",       // 'new' | 'existing'
  merchantMobile: "",          // required if new
  merchantId: "",              // required if existing
  loggedThisSession: false,

  // Step 1 - platform / entity
  onboardingType: null,   // 'ace' | 'salesforce'
  entityType: null,       // 'partnership' | 'company'
  docRequirement: null,   // set id
  step: 0,

  // Shared legal/entity profile (filled once, reused across every document)
  firmName: "",
  regAddress: "",
  includeLetterhead: true,
  principalSame: "same",
  principalAddress: "",
  partnershipRegType: "",
  resolutionVersion: "v5a",
  deedDate: "",

  // Partners / Directors
  partners: [INITIAL_PARTNER(1), INITIAL_PARTNER(2)],
  nextPartnerId: 3,
  presentPartnerIds: [],

  // Resolution
  resolutionDateRaw: "",
  resolutionTimeRaw: "",
  pepStatusRes: "no",

  // BO
  boDate: "",
  boCategory: "cat1",
  pepStatusBO: "no",
  companyListingStatus: "not_listed",
  stockExchangeName: "",
  boExternalAS: false,
  boExternalASName: "",
  boExternalASDesignation: "",

  // Company certification rules
  isOPC: false,
  opcApprovalConfirmed: false,

  // MDF
  mdfAuthName: "", mdfAuthDesignation: "", mdfAuthPan: "",
  mdfMobile: "", mdfEmail: "",
  mdfPwd: "no", mdfPwdType: "", mdfPwdPct: "",
  mdfFatherName: "", mdfKycDoc: "aadhaar",
  mdfEntityNature: "na",
  mdfTanStatus: "no_tan", mdfTanNum: "",
  mdfGstStatus: "no_gst", mdfGstNum: "",
  mdfPepStatus: "no", mdfDateRaw: "", mdfPlace: "",

  formValidated: false,
  inlineErrors: {},
  toastMessage: "",
};

let formData = {};

// ---- DOCUMENT SET CONFIGURATION ----
const docOptionsMap = {
  ace: {
    partnership: [
      { id: "set1", label: "Partner Resolution" },
      { id: "set2", label: "BO Declaration" },
      { id: "set3", label: "Partner Resolution + BO Declaration" },
      { id: "set4", label: "Partner Resolution + BO Declaration + ACE OSV (MDF)" },
      { id: "set17", label: "ACE OSV (MDF)" },
    ],
    company: [
      { id: "set5", label: "Board Resolution" },
      { id: "set6", label: "BO Declaration" },
      { id: "set7", label: "Board Resolution + BO Declaration" },
      { id: "set8", label: "Board Resolution + BO Declaration + ACE OSV (MDF)" },
      { id: "set18", label: "ACE OSV (MDF)" },
    ],
  },
  salesforce: {
    partnership: [
      { id: "set9", label: "Partner Resolution" },
      { id: "set10", label: "BO Declaration" },
      { id: "set11", label: "MDF" },
      { id: "set12", label: "Partner Resolution + BO Declaration + MDF" },
    ],
    company: [
      { id: "set13", label: "Board Resolution" },
      { id: "set14", label: "BO Declaration" },
      { id: "set15", label: "MDF" },
      { id: "set16", label: "Board Resolution + BO Declaration + MDF" },
    ],
  },
};
const needsResolution = (set) => ["set1","set3","set4","set5","set7","set8","set9","set12","set13","set16"].includes(set);
const needsBO         = (set) => ["set2","set3","set4","set6","set7","set8","set10","set12","set14","set16"].includes(set);
const needsMDF        = (set) => ["set4","set8","set11","set12","set15","set16","set17","set18"].includes(set);
const needsFullKYC    = (set) => needsBO(set);
const needsOwnership = (set) => needsBO(set) || needsResolution(set);

// ---- STORAGE ----
function loadAgent() {
  try {
    const raw = localStorage.getItem(LS_AGENT);
    if (!raw) return;
    const d = JSON.parse(raw);
    if (d.agentMobile) state.agentMobile = d.agentMobile;
  } catch (e) {}
}
function saveAgent() {
  try { localStorage.setItem(LS_AGENT, JSON.stringify({ agentMobile: state.agentMobile })); } catch (e) {}
}
function loadMerchant() {
  try {
    const raw = localStorage.getItem(LS_MERCHANT);
    if (!raw) return;
    const d = JSON.parse(raw);
    Object.assign(state, d, { agentMobile: state.agentMobile }); // keep agent separate
    // Migrate data saved by earlier versions that used favourable booleans.
    if (!state.pepStatusRes && typeof d.pepDeclarationRes === "boolean") state.pepStatusRes = d.pepDeclarationRes ? "no" : "yes";
    if (!state.pepStatusBO && typeof d.pepDeclarationBO === "boolean") state.pepStatusBO = d.pepDeclarationBO ? "no" : "yes";
    if (!state.mdfPepStatus && typeof d.mdfPepConfirm === "boolean") state.mdfPepStatus = d.mdfPepConfirm ? "no" : "yes";
    if (!state.mdfPepStatus) state.mdfPepStatus = "no";
    if (!state.resolutionVersion) state.resolutionVersion = "v5a";
    if (!state.pepStatusRes) state.pepStatusRes = "no";
    if (!state.pepStatusBO) state.pepStatusBO = "no";
    // v4 letterhead uses the legal firm name and main-office address directly.
    delete state.lhTagline; delete state.lhPhone; delete state.lhEmail; delete state.lhWebsite;
    const allowedSets = ["set1","set2","set3","set4","set5","set6","set7","set8","set9","set10","set11","set12","set13","set14","set15","set16","set17","set18"];
    if (!allowedSets.includes(state.docRequirement)) state.docRequirement = null;
  } catch (e) {}
}
function saveMerchant() {
  try {
    const { agentMobile, ...rest } = state;
    localStorage.setItem(LS_MERCHANT, JSON.stringify(rest));
  } catch (e) {}
}
function clearMerchantData() {
  if (!confirm("Start a NEW merchant onboarding? All current merchant/entity data will be cleared. Your mobile number will be kept.")) return;
  localStorage.removeItem(LS_MERCHANT);
  location.reload();
}
function globalReset() {
  if (!confirm("GLOBAL RESET will permanently clear all merchant data and the saved agent mobile number. Continue?")) return;
  localStorage.removeItem(LS_MERCHANT);
  localStorage.removeItem(LS_AGENT);
  location.reload();
}
function resetSection(section) {
  const labels = {
    tracking: "merchant tracking details",
    platform: "platform and entity selection",
    documents: "document selection",
    entity: "entity and letterhead details",
    members: "partners/directors",
    resolution: "resolution details",
    bo: "BO declaration details",
    mdf: "MDF details",
  };
  if (!confirm("Reset " + labels[section] + "? Other sections will be kept.")) return;
  if (section === "tracking") Object.assign(state, { merchantStatus: "new", merchantMobile: "", merchantId: "", loggedThisSession: false });
  if (section === "platform") Object.assign(state, { onboardingType: null, entityType: null, docRequirement: null });
  if (section === "documents") Object.assign(state, { docRequirement: null });
  if (section === "entity") Object.assign(state, { firmName: "", regAddress: "", includeLetterhead: true, principalSame: "same", principalAddress: "", partnershipRegType: "", deedDate: "" });
  if (section === "members") Object.assign(state, { partners: [INITIAL_PARTNER(1, state.entityType === "company" ? "Director" : "Partner"), INITIAL_PARTNER(2, state.entityType === "company" ? "Director" : "Partner")], nextPartnerId: 3, presentPartnerIds: [] });
  if (section === "resolution") Object.assign(state, { resolutionDateRaw: "", resolutionTimeRaw: "", resolutionVersion: "v5a", pepStatusRes: "no", presentPartnerIds: [], isOPC: false, opcApprovalConfirmed: false });
  if (section === "bo") Object.assign(state, { boDate: "", boCategory: "cat1", pepStatusBO: "no", companyListingStatus: "not_listed", stockExchangeName: "", boExternalAS: false, boExternalASName: "", boExternalASDesignation: "" });
  if (section === "mdf") Object.assign(state, { mdfAuthName: "", mdfAuthDesignation: state.entityType === "partnership" ? "Partner" : "Director", mdfAuthPan: "", mdfMobile: "", mdfEmail: "", mdfPwd: "no", mdfPwdType: "", mdfPwdPct: "", mdfFatherName: "", mdfKycDoc: "aadhaar", mdfEntityNature: "na", mdfTanStatus: "no_tan", mdfTanNum: "", mdfGstStatus: "no_gst", mdfGstNum: "", mdfPepStatus: "no", mdfDateRaw: "", mdfPlace: "" });
  state.inlineErrors = {};
  state.formValidated = false;
  saveMerchant();
  rerender();
}
function resetPage() {
  const sections = state.step === 0 ? ["tracking"] : state.step === 1 ? ["platform"] : state.step === 2 ? ["documents"] : state.step === 3 ? ["entity", "members", "resolution", "bo", "mdf"] : [];
  if (!sections.length || !confirm("Reset every field on this page? Data on other pages will be kept.")) return;
  // The individual resets are applied without repeated confirmations.
  if (state.step === 0) Object.assign(state, { merchantStatus: "new", merchantMobile: "", merchantId: "", loggedThisSession: false });
  if (state.step === 1) Object.assign(state, { onboardingType: null, entityType: null, docRequirement: null });
  if (state.step === 2) state.docRequirement = null;
  if (state.step === 3) {
    Object.assign(state, { firmName: "", regAddress: "", includeLetterhead: true, principalSame: "same", principalAddress: "", partnershipRegType: "", deedDate: "", partners: [INITIAL_PARTNER(1, state.entityType === "company" ? "Director" : "Partner"), INITIAL_PARTNER(2, state.entityType === "company" ? "Director" : "Partner")], nextPartnerId: 3, presentPartnerIds: [], resolutionDateRaw: "", resolutionTimeRaw: "", resolutionVersion: "v5a", pepStatusRes: "no", isOPC: false, opcApprovalConfirmed: false, boDate: "", boCategory: "cat1", pepStatusBO: "no", companyListingStatus: "not_listed", stockExchangeName: "", boExternalAS: false, boExternalASName: "", boExternalASDesignation: "", mdfAuthName: "", mdfAuthDesignation: state.entityType === "partnership" ? "Partner" : "Director", mdfAuthPan: "", mdfMobile: "", mdfEmail: "", mdfPwd: "no", mdfPwdType: "", mdfPwdPct: "", mdfFatherName: "", mdfKycDoc: "aadhaar", mdfEntityNature: "na", mdfTanStatus: "no_tan", mdfTanNum: "", mdfGstStatus: "no_gst", mdfGstNum: "", mdfPepStatus: "no", mdfDateRaw: "", mdfPlace: "" });
  }
  state.inlineErrors = {};
  state.formValidated = false;
  saveMerchant();
  rerender();
}

let saveTimer = null;
const scheduleSave = () => {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => { saveMerchant(); saveAgent(); flashSave(); }, 250);
};
function flashSave() {
  const el = document.getElementById("saveLabel");
  if (!el) return;
  el.textContent = "Saved \u2713";
  setTimeout(() => { if (el) el.textContent = "Auto-save ON"; }, 1400);
}

// ---- SHEET LOGGING (log-only, 2 datapoints) ----
// Submitted as a real hidden-form POST (not fetch) so that, on a device/browser
// already signed into the corporate @phonepe.com Google account, the request
// carries a valid Google session and succeeds even though the Workspace admin
// restricts Apps Script deployments to "Anyone within the organization" rather
// than fully public "Anyone". If the agent isn't signed in, the hidden iframe
// silently shows a Google sign-in page instead of logging \u2014 harmless, and it
// never blocks or interrupts document generation.
function logToSheet() {
  if (!SHEETS_URL || state.loggedThisSession) return;
  try {
    const payload = {
      timestamp: new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
      agentMobile: state.agentMobile,
      merchant: state.merchantStatus === "new" ? state.merchantMobile : state.merchantId,
      merchantStatus: state.merchantStatus,
    };
    submitViaHiddenForm(SHEETS_URL, payload);
    state.loggedThisSession = true;
  } catch (e) { /* silent */ }
}

function submitViaHiddenForm(url, data) {
  let iframe = document.getElementById("sheetLogFrame");
  if (!iframe) {
    iframe = document.createElement("iframe");
    iframe.id = "sheetLogFrame";
    iframe.name = "sheetLogFrame";
    iframe.style.display = "none";
    document.body.appendChild(iframe);
  }
  const form = document.createElement("form");
  form.method = "POST";
  form.action = url;
  form.target = "sheetLogFrame";
  form.style.display = "none";
  Object.entries(data).forEach(([key, value]) => {
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = key;
    input.value = value == null ? "" : String(value);
    form.appendChild(input);
  });
  document.body.appendChild(form);
  form.submit();
  setTimeout(() => { form.remove(); }, 3000);
}

// ---- TINY RENDER ENGINE (event delegation + focus-preserving rerender) ----
const appRoot = document.getElementById("app");
let handlerRegistry = {};
let handlerSeq = 0;
let isRendering = false;
const EVENT_TYPES = ["click", "input", "change", "blur"];

const on = (event, fn) => {
  const key = "h" + handlerSeq++;
  handlerRegistry[key] = fn;
  return 'data-on-' + event + '="' + key + '"';
};
const attr = (v = "") => String(v).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
const esc = (v = "") => attr(v).replaceAll("'", "&#39;");
const option = (value, label, current) => '<option value="' + attr(value) + '" ' + (value === current ? "selected" : "") + ">" + label + "</option>";

function bindHandlers(container) {
  EVENT_TYPES.forEach((event) => {
    container.querySelectorAll("[data-on-" + event + "]").forEach((node) => {
      const key = node.getAttribute("data-on-" + event);
      const fn = handlerRegistry[key];
      if (fn) node.addEventListener(event, fn);
    });
  });
}

function rerender() {
  if (isRendering) return;
  isRendering = true;
  try {
    const active = document.activeElement;
    const focusId = active && active.getAttribute ? active.getAttribute("data-fid") : null;
    const selStart = active && "selectionStart" in active ? active.selectionStart : null;
    handlerRegistry = {}; handlerSeq = 0;
    appRoot.innerHTML = renderApp();
    renderProgress();
    bindHandlers(appRoot);
    if (focusId) {
      const next = appRoot.querySelector('[data-fid="' + focusId + '"]');
      if (next) {
        next.focus();
        if (selStart != null && "setSelectionRange" in next) {
          try { next.setSelectionRange(selStart, selStart); } catch (e) {}
        }
      }
    }
  } finally { isRendering = false; }
}

let toastTimer = null;
function showToast(message, cls) {
  const el = document.getElementById("toast");
  el.textContent = message;
  el.className = "show" + (cls ? " " + cls : "");
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = ""; }, 3200);
}

// ---- FORMAT HELPERS ----
function parseDateParts(dateValue = "") {
  const norm = String(dateValue).trim().replaceAll("/", "-");
  const [a = "", b = "", c = ""] = norm.split("-").filter(Boolean);
  if (!a || !b || !c) return null;
  if (a.length === 4) return { year: a, month: b, day: c };
  if (c.length === 4) return { year: c, month: b, day: a };
  return null;
}
function formatDate(dateValue = "") {
  const p = parseDateParts(dateValue);
  return p ? p.day + "-" + p.month + "-" + p.year : "";
}
function formatTime(timeValue = "") {
  if (!timeValue) return "";
  const [h, m] = timeValue.split(":");
  const hr = Number(h);
  return (hr % 12 || 12) + ":" + m + (hr >= 12 ? "PM" : "AM");
}
function sanitizeShare(value = "") {
  let cleaned = String(value).replace(/[^0-9.]/g, "");
  const firstDot = cleaned.indexOf(".");
  if (firstDot !== -1) cleaned = cleaned.slice(0, firstDot + 1) + cleaned.slice(firstDot + 1).replace(/\./g, "");
  return cleaned === "" || cleaned === "." ? cleaned : (Number(cleaned) > 100 ? "100" : cleaned);
}

// ---- DERIVED DATA ----
function computeDerived() {
  const normalizedPartners = state.partners.map((p) => {
    const parts = parseDateParts(p.dobRaw);
    return {
      ...p,
      designation: state.entityType === "partnership" ? "Partner" : p.designation,
      pan: (p.pan || "").toUpperCase(),
      dob: parts ? parts.day + "/" + parts.month + "/" + parts.year : "",
      poaNumDisplay: p.poa === "AADHAAR" ? ("XXXX XXXX " + (p.poaNum || "")).trim() : p.poaNum,
    };
  });
  const presentPartners = normalizedPartners.filter((p) => state.presentPartnerIds.includes(p.id));
  const authSignatory = normalizedPartners.find((p) => p.isAS);
  formData = {
    ...state,
    resolutionDate: formatDate(state.resolutionDateRaw),
    resolutionTime: formatTime(state.resolutionTimeRaw),
    boDateFmt: formatDate(state.boDate),
    mdfDate: formatDate(state.mdfDateRaw),
    deedDateFmt: formatDate(state.deedDate),
    isRegistered: state.partnershipRegType === "registered",
    partners: normalizedPartners,
    presentPartners,
    authSignatoryName: authSignatory ? authSignatory.name : "",
    authSignatoryDesignation: authSignatory ? authSignatory.designation : "",
  };
}

// ---- PARTNER CRUD ----
function addPartner() {
  const designation = state.entityType === "company" ? "Director" : "Partner";
  state.partners = [...state.partners, INITIAL_PARTNER(state.nextPartnerId++, designation)];
  rerender();
}
function removePartner(id) {
  const minimum = state.entityType === "company" && state.isOPC ? 1 : 2;
  if (state.partners.length <= minimum) { showToast("Minimum " + minimum + " required", "er"); return; }
  state.partners = state.partners.filter((p) => p.id !== id);
  state.presentPartnerIds = state.presentPartnerIds.filter((pid) => pid !== id);
  rerender();
}
function onSelectAS(id) {
  state.partners = state.partners.map((p) => ({ ...p, isAS: p.id === id }));
  const selected = state.partners.find((p) => p.id === id);
  if (selected) {
    state.mdfAuthName = selected.name;
    state.mdfAuthDesignation = selected.designation;
    state.mdfAuthPan = selected.pan;
  }
  rerender();
}
function onPartnerChange(id, field, value) {
  state.partners = state.partners.map((p) => (p.id === id ? { ...p, [field]: value } : p));
  scheduleSave();
}
// When switching entity type, retro-fit the placeholder default designation
// ("Partner" <-> "Director") for members who haven't customised it yet.
function syncDefaultDesignations() {
  const swap = { Partner: "Director", Director: "Partner" };
  const target = state.entityType === "company" ? "Director" : "Partner";
  state.partners = state.partners.map((p) =>
    (p.designation === "Partner" || p.designation === "Director") ? { ...p, designation: target } : p
  );
}
function togglePresentPartner(id, checked) {
  state.presentPartnerIds = checked
    ? [...state.presentPartnerIds, id]
    : state.presentPartnerIds.filter((pid) => pid !== id);
  rerender();
}
function prefillDateTime() {
  const now = new Date();
  const d = now.getFullYear() + "-" + String(now.getMonth()+1).padStart(2,"0") + "-" + String(now.getDate()).padStart(2,"0");
  const t = String(now.getHours()).padStart(2,"0") + ":" + String(now.getMinutes()).padStart(2,"0");
  state.resolutionDateRaw = d; state.resolutionTimeRaw = state.entityType === "company" ? t : ""; state.boDate = d; state.mdfDateRaw = d;
  rerender();
}

// ---- VALIDATION ----
function validateStep0() {
  const errs = {};
  if (!MOBILE_REGEX.test(state.agentMobile)) errs.agentMobile = "Enter a valid 10-digit mobile number";
  if (state.merchantStatus === "new" && !MOBILE_REGEX.test(state.merchantMobile)) errs.merchantMobile = "Enter the merchant's valid 10-digit mobile number";
  if (state.merchantStatus === "existing" && !state.merchantId.trim()) errs.merchantId = "Merchant ID is required";
  state.inlineErrors = errs;
  rerender();
  return Object.keys(errs).length === 0;
}

function validateDataEntry() {
  const inlineErrors = {};
  const addError = (msg, field) => { if (field && !inlineErrors[field]) inlineErrors[field] = msg; };

  if (!state.firmName.trim()) addError("Legal name is required.", "firmName");
  if (!state.regAddress.trim()) addError("Registered address is required.", "regAddress");
  if (state.entityType === "partnership" && !state.partnershipRegType) addError("Select the partnership registration type.", "partnershipRegType");
  if (state.principalSame === "diff" && !state.principalAddress.trim()) addError("Principal office address is required.", "principalAddress");

  if (needsResolution(state.docRequirement) || needsBO(state.docRequirement)) {
    const minimumMembers = state.entityType === "company" && state.isOPC ? 1 : 2;
    if (state.partners.length < minimumMembers) addError("Minimum " + minimumMembers + " individual" + (minimumMembers > 1 ? "s are" : " is") + " required.", "partners");
    const internalASRequired = needsResolution(state.docRequirement) || (needsBO(state.docRequirement) && (!state.boExternalAS || state.boCategory === "cat2"));
    if (internalASRequired && !state.partners.some((p) => p.isAS)) addError('Tick "Authorised Signatory" for one member.', "authSignatory");

    let totalShare = 0;
    const requiresFull = needsFullKYC(state.docRequirement);
    const requiresOwnership = needsOwnership(state.docRequirement);
    state.partners.forEach((partner, index) => {
      const label = "Member " + (index + 1) + ": ";
      if (!partner.name.trim()) addError(label + "Name is required.", "name_" + partner.id);
      if (!partner.designation.trim()) addError(label + "Designation is required.", "designation_" + partner.id);
      if (requiresFull) {
        if (!partner.pan || !PAN_REGEX.test(partner.pan.toUpperCase())) addError(label + "Valid PAN required.", "pan_" + partner.id);
        if (!partner.dobRaw) addError(label + "DOB required.", "dobRaw_" + partner.id);
        if (!partner.address.trim()) addError(label + "Address required.", "address_" + partner.id);
        if (!partner.poaNum) addError(label + "Proof number required.", "poaNum_" + partner.id);
        else if (partner.poa === "AADHAAR" && !AADHAAR_LAST4.test(partner.poaNum)) addError(label + "Aadhaar proof must be exactly last 4 digits.", "poaNum_" + partner.id);
      }
      if (requiresOwnership) {
        if (partner.share === "") addError(label + "Share % required.", "share_" + partner.id);
        else totalShare += Number(partner.share || 0);
      }
    });
    if (requiresOwnership && Math.abs(totalShare - 100) > 0.6) {
      addError("Total shareholding must equal exactly 100% (currently " + totalShare.toFixed(1) + "%).", "shareTotal");
    }
  }

  if (needsResolution(state.docRequirement)) {
    if (!state.resolutionDateRaw) addError("Resolution date required.", "resolutionDateRaw");
    if (state.entityType === "company" && !state.resolutionTimeRaw) addError("Resolution time required.", "resolutionTimeRaw");
    if (state.presentPartnerIds.length === 0) addError("Select at least one present signatory.", "presentPartnerIds");
    if (state.entityType === "partnership") {
      if (!state.resolutionVersion) addError("Select Partner Resolution Version 5A or 5B.", "resolutionVersion");
      if (!state.pepStatusRes) addError("Select Yes or No for the resolution PEP declaration.", "pepStatusRes");
      const signingShare = state.partners.filter((partner) => state.presentPartnerIds.includes(partner.id)).reduce((sum, partner) => sum + Number(partner.share || 0), 0);
      if (signingShare < 50) addError("Selected signing partners must collectively hold at least 50% ownership (currently " + signingShare.toFixed(1) + "%).", "presentShare");
    }
    if (state.entityType === "company") {
      if (!state.presentPartnerIds.length) addError("Select at least one person to sign the Board Resolution.", "companySigners");
    }
  }

  if (needsBO(state.docRequirement)) {
    if (!state.boDate) addError("BO declaration date required.", "boDate");
    if (!state.pepStatusBO) addError("Select Yes or No for the BO PEP declaration.", "pepStatusBO");
    if (state.entityType === "company" && state.companyListingStatus !== "not_listed" && !state.stockExchangeName.trim()) {
      addError("Stock exchange name is required for the selected listing status.", "stockExchangeName");
    }
    if (state.boCategory === "cat1") {
      const threshold = state.entityType === "company" ? 10 : state.partnershipRegType === "registered" ? 10 : 15;
      if (!state.partners.some((partner) => Number(partner.share) > threshold)) {
        addError("Category 1 requires at least one owner above the applicable threshold.", "boCategory");
      }
    }
    if (state.entityType === "company" && state.boExternalAS) {
      if (!state.boExternalASName.trim()) addError("External BO authorised signatory name is required.", "boExternalASName");
      if (!state.boExternalASDesignation.trim()) addError("External BO authorised signatory designation is required.", "boExternalASDesignation");
    }
    if (needsResolution(state.docRequirement) && state.resolutionDateRaw && state.boDate && formatDate(state.resolutionDateRaw) !== formatDate(state.boDate)) {
      addError("BO date must match Resolution date.", "boDate");
    }
  }

  if (needsMDF(state.docRequirement)) {
    if (!state.mdfAuthName.trim()) addError("Signatory Name is required.", "mdfAuthName");
    if (!state.mdfAuthDesignation.trim()) addError("Signatory designation is required.", "mdfAuthDesignation");
    if (!state.mdfAuthPan || !PAN_REGEX.test(state.mdfAuthPan.toUpperCase())) addError("Valid PAN is required.", "mdfAuthPan");
    if (!MOBILE_REGEX.test(state.mdfMobile)) addError("Valid mobile number required.", "mdfMobile");
    if (state.mdfEmail && !EMAIL_REGEX.test(state.mdfEmail)) addError("Enter a valid email address or leave it blank.", "mdfEmail");
    if (state.onboardingType === "salesforce" && !state.mdfEmail) addError("Email address is required for Salesforce MDF.", "mdfEmail");
    if (state.mdfPwd === "yes" && (!state.mdfPwdType.trim() || !state.mdfPwdPct || Number(state.mdfPwdPct) <= 0 || Number(state.mdfPwdPct) > 100)) {
      addError("Enter disability type and a valid percentage.", "mdfPwdPct");
    }
    if (state.mdfTanStatus === "has_tan" && !TAN_REGEX.test(state.mdfTanNum.toUpperCase())) addError("Enter a valid TAN.", "mdfTanNum");
    if (state.onboardingType === "salesforce" && !state.mdfFatherName.trim()) addError("Father's name is required for Salesforce MDF.", "mdfFatherName");
    if (state.onboardingType === "salesforce" && state.mdfGstStatus === "has_gst" && !GSTIN_REGEX.test(state.mdfGstNum.toUpperCase())) addError("Enter a valid GSTIN.", "mdfGstNum");
    if (!state.mdfPepStatus) addError("Select Yes or No for the MDF PEP declaration.", "mdfPepStatus");
    if (!state.mdfDateRaw) addError("MDF date is required.", "mdfDateRaw");
    if (!state.mdfPlace.trim()) addError("MDF place is required.", "mdfPlace");
  }

  state.inlineErrors = inlineErrors;
  if (Object.keys(inlineErrors).length === 0) {
    state.formValidated = true;
    showToast("Validation successful. Documents ready to generate.", "ok");
  } else {
    state.formValidated = false;
    showToast("Please fix the highlighted fields.", "er");
  }
  rerender();
}

// ============================================================
// DOCX GENERATION (client-side, via docx.js)
// ============================================================
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, UnderlineType, PageBreak, Header,
} = docxLib;

const bdr = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const bdrs = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const noBdr = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBdrs = { top: noBdr, bottom: noBdr, left: noBdr, right: noBdr, insideH: noBdr, insideV: noBdr };

const t = (text, opts = {}) => new TextRun({ text: text || "", font: "Times New Roman", size: 22, ...opts });
const u = (text, opts = {}) => new TextRun({ text: text || "", font: "Times New Roman", size: 22, bold: true, underline: { type: UnderlineType.SINGLE }, ...opts });
const th = (text) => new TextRun({ text: text || "", font: "Times New Roman", size: 20, bold: true });
const p = (children, opts = {}) => new Paragraph({ children, spacing: { after: 120 }, ...opts });

function cell(content, w, opts = {}) {
  const children = Array.isArray(content) ? content
    : [new Paragraph({ children: [typeof content === "string" ? t(content) : content], spacing: { after: 0 } })];
  return new TableCell({ borders: bdrs, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: w, type: WidthType.DXA }, ...opts, children });
}
function ucell(text, w, opts = {}) {
  return new TableCell({ borders: bdrs, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: w, type: WidthType.DXA }, ...opts,
    children: [new Paragraph({ children: [u(text)], spacing: { after: 0 } })] });
}
function hCell(text, w) {
  return new TableCell({ borders: bdrs, shading: { fill: "D9D9D9", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
    width: { size: w, type: WidthType.DXA }, children: [new Paragraph({ children: [th(text)], spacing: { after: 0 } })] });
}
function sigTable(members) {
  const list = members && members.length ? members : [{ name: "", designation: "" }];
  const colW = Math.floor(9360 / list.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: list.map(() => colW), borders: noBdrs,
    rows: [
      new TableRow({ cantSplit: true, children: list.map(() => new TableCell({ borders: noBdrs, width: { size: colW, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [p([t("____________ (Sign)")])] })) }),
      new TableRow({ cantSplit: true, children: list.map((m) => new TableCell({ borders: noBdrs, width: { size: colW, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [p([u(m.name)]), p([t("(" ), u(m.designation||""), t(")")])] })) }),
    ],
  });
}
function buildLetterheadHeader(lh) {
  const firmName = lh.firmName || "";
  const officeLine = lh.regAddress ? "Regd Office: " + lh.regAddress + (lh.principalSame === "diff" && lh.principalAddress ? " | Principal Office: " + lh.principalAddress : "") : "";
  const firmSize = firmName.length > 32 ? 22 : firmName.length > 22 ? 24 : 28;
  const officeSize = officeLine.length > 110 ? 13 : officeLine.length > 80 ? 15 : 18;
  const lt = (text, opts={}) => new TextRun({ text: text||"", font: "Arial", size: officeSize, ...opts });
  const logoCell = new TableCell({ borders: noBdrs, noWrap: true, width: { size: 3200, type: WidthType.DXA }, margins: { top:0,bottom:0,left:0,right:120 },
    children: [
      new Paragraph({ children: [new TextRun({ text: firmName, font:"Arial", size:firmSize, bold:true, color:"000000" })], spacing:{after:0} }),
    ]});
  const contactCell = new TableCell({ borders: noBdrs, noWrap: true, width: { size: 6880, type: WidthType.DXA }, margins:{top:0,bottom:0,left:120,right:0},
    children: [new Paragraph({ children:[lt(officeLine,{bold:true,color:"000000"})], alignment: AlignmentType.RIGHT, spacing:{after:0} })] });
  const lhTable = new Table({ width:{size:10080,type:WidthType.DXA}, columnWidths:[3200,6880], borders: noBdrs, rows:[new TableRow({cantSplit:true,children:[logoCell,contactCell]})] });
  const hrPara = new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 5, color: "5F259F", space: 3 } }, spacing: { after: 80 } });
  return new Header({ children: [lhTable, hrPara] });
}
function pageSetup(hasLH) {
  return { size: { width: 12240, height: 15840 }, margin: { top: hasLH ? 1150 : 900, right: 900, bottom: 650, left: 900 } };
}
// ---- PARTNERSHIP: RESOLUTION ----
async function buildPartnershipResolution(d) {
  const hasLH = d.includeLetterhead && d.firmName;
  const principalText = d.principalSame === "same" ? "same \u2611 (tick if applicable)" : d.principalAddress;
  const versionFiveDeclaration = d.resolutionVersion === "v5b"
    ? "3. In case one or more Partners in the firm are non-Natural persons, the share holding of the Partners is as per the Deed. We are submitting with this declaration, the identity proofs of the Ultimate Beneficial Owner and of natural persons greater than above mentioned thresholds (applicable to us) in the firm."
    : "3. It is hereby confirmed that all partners of the firm are Natural persons only and the share holding of the Partners is as per the Deed.";
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(hasLH) },
    headers: hasLH ? { default: buildLetterheadHeader(d) } : undefined,
    children: [
      ...(!hasLH ? [p([t("(on partnership firm\u2019s letterhead)", { italics: true })], { alignment: AlignmentType.CENTER })] : []),
      p([t("TO WHOMSOEVER IT MAY CONCERN", { bold: true })], { alignment: AlignmentType.CENTER }),
      p([]),
      new Paragraph({ children: [
        t("RESOLUTION OF THE PARTNERS PASSED AT THE MEETING OF THE PARTNERS OF ", { bold: true }),
        u((d.firmName||"").toUpperCase(), { bold: true }),
        t(" (\u201cFIRM\u201d) HELD ON ", { bold: true }), u(d.resolutionDate, { bold: true }),
        t(" having its registered office address at "), u(d.regAddress),
        t(" and having its outlet or principal place of operation / office at "),
        ...(d.principalSame === "same" ? [t(principalText)] : [u(d.principalAddress)]), t("."),
      ], spacing: { after: 200 } }),
      p([]),
      p([t("PRESENT:", { bold: true })]),
      ...d.presentPartners.map((m, i) => p([t((i + 1) + ". "), u(m.name)])),
      p([t("(List of partners present during the resolution)", { italics: true })]),
      p([]),
      new Paragraph({ children: [
        t("RESOLVED THAT ", { bold: true }), t("Mr/Mrs "), u(d.authSignatoryName), t(" [Partner/Authorized Signatory], be and is hereby authorized, to act on behalf of the Firm and to execute/sign all necessary applications/documents for the purpose of opening and operating a business account with PhonePe Limited."),
      ], spacing: { after: 200 } }),
      p([]),
      new Paragraph({ children: [
        t("RESOLVED FURTHER THAT ", { bold: true }), t("all acts, deeds, and things done by the said Partner(s) in this regard shall be binding upon the Firm and all its partners and shall remain in force. "),
        t("RESOLVED FURTHER THAT ", { bold: true }), t("this resolution shall remain in force until a written notice of its withdrawal or amendment is served upon and acknowledged by PhonePe Limited, and that a certified true copy of this resolution be furnished to PhonePe Limited for their records."),
      ], spacing: { after: 240 } }),
      p([]),
      sigTable(d.presentPartners),
      p([]),
      p([t("Seal of the Firm")], { alignment: AlignmentType.CENTER }),

    ],
  }, {
    properties: { page: pageSetup(false) },
    headers: { default: new Header({ children: [] }) },
    children: [
      new Paragraph({ children: [new TextRun({ text: "Declaration", font: "Times New Roman", size: 24, bold: true })], spacing: { after: 160 } }),
      new Paragraph({ children: [t("I/we, the undersigned individuals, hereby personally, jointly, and severally undertake and declare that:")], spacing: { after: 160 } }),
      new Paragraph({ children: [t("1. Our firm "), u(d.firmName), t(" is constituted as a partnership firm and it is")], spacing: { after: 80 } }),
      p([t((d.isRegistered ? "\u2611" : "\u2610") + " Registered (LLP or Registry done in Registrar office of the Deed)")], { indent: { left: 720 } }),
      p([t((!d.isRegistered ? "\u2611" : "\u2610") + " Unregistered (Normal Deed Notarized or not Notarized)")], { indent: { left: 720 } }),
      new Paragraph({ children: [t("2. The partners holding more than 10% shares/control for registered partnership and 15% shares/control for un-registered partnership are beneficial owners of the entity. In case no natural person holds more than 10%/15% shares/control in the entity, the authorised signatory should be considered as senior management for the purpose of BO identification.")], spacing: { after: 100 } }),
      new Paragraph({ children: [t(versionFiveDeclaration)], spacing: { after: 100 } }),
      new Paragraph({ children: [t("4. Our personnel(s), partner(s), director(s), officer(s), or our family member(s) or our close associate(s) and beneficial owners, is a Politically Exposed Person. (\u201cPolitically Exposed Persons\u201d (PEPs) are individuals who are or have been entrusted with prominent public functions by a foreign country, including the Heads of States/Governments, senior politicians, senior government or judicial or military officers, senior executives of state-owned corporations and important political party officials.)")], spacing: { after: 80 } }),
      p([t((d.pepStatusRes === "yes" ? "\u2611" : "\u2610") + " YES     " + (d.pepStatusRes === "no" ? "\u2611" : "\u2610") + " NO")], { indent: { left: 720 } }),
      new Paragraph({ children: [t("5. The contents of the resolution of the partners provided to PhonePe are true and valid.")], spacing: { after: 100 } }),
      new Paragraph({ children: [t("6. The list of partners constituting the partnership firm and their respective details, as provided in the partnership deed, provided to PhonePe, are true, complete, current, and valid.")], spacing: { after: 100 } }),
      new Paragraph({ children: [t("The percentage of ownership and/or entitlement to capital or profits of the partners as specified in the Partnership Deed is true, complete, current, and valid. In consideration of PhonePe agreeing to rely on the Resolution, the Partnership Deed, and this declaration, I/we hereby personally, jointly, and severally undertake to indemnify and hold PhonePe harmless against all damages, liabilities, claims, demands, actions, proceedings, losses, costs (including legal costs), expenses, and all other liabilities of whatsoever nature or description, arising out of or in connection with PhonePe\u2019s reliance on the said Resolution, Partnership Deed, and this declaration.")], spacing: { after: 160 } }),
      p([]),
      sigTable(d.presentPartners),
    ],
  }] });
  return doc;
}

// ---- PARTNERSHIP: BO DECLARATION ----
async function buildPartnershipBO(d) {
  const hasLH = d.includeLetterhead && d.firmName;
  const principalCell = d.principalSame === "same" ? "\u2611 Same" : "\u2610 Same   OR " ;
  const boCols = [400, 1100, 1800, 680, 680, 980, 980, 680, 780];
  const threshold = d.isRegistered ? 10 : 15;
  const boRows = d.boCategory === "cat1"
    ? d.partners.filter((m) => Number(m.share) > threshold)
    : d.partners.filter((m) => m.isAS);
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(hasLH) },
    headers: hasLH ? { default: buildLetterheadHeader(d) } : undefined,
    children: [
      new Paragraph({ children: [new TextRun({ text: "DECLARATION OF BENEFICIAL OWNERSHIP (BO) and LIST OF PARTNERS", font: "Times New Roman", size: 26, bold: true, underline: { type: UnderlineType.SINGLE } })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: [500, 2500, 7080], rows: [
        new TableRow({ cantSplit: true, children: [cell("I", 500), cell("Name of the entity", 2500), ucell(d.firmName, 7080)] }),
        new TableRow({ cantSplit: true, children: [cell("II", 500), cell("Registered address", 2500), ucell(d.regAddress, 7080)] }),
        new TableRow({ cantSplit: true, children: [cell("III", 500), cell("Principal place of operation/office", 2500),
          new TableCell({ borders: bdrs, width: { size: 7080, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: d.principalSame === "same" ? [p([t("\u2611 Same")])] : [p([t("\u2610 Same   OR "), u(d.principalAddress)])] }) ] }),
        new TableRow({ cantSplit: true, children: [cell("IV", 500), cell("Type of entity", 2500),
          new TableCell({ borders: bdrs, width: { size: 7080, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [p([t(d.isRegistered ? "\u2611" : "\u2610"), t(" Partnership Firm / LLP")]), p([t(d.isRegistered ? "\u2610" : "\u2611"), t(" Unregistered Partnership Firm")])] }) ] }),
      ] }),
      p([]),
      new Paragraph({ children: [t("The Legal Entity as stated above hereby confirms and declares the following on the below date: ", { bold: true }), u(d.boDateFmt, { bold: true })], spacing: { after: 100 } }),
      new Paragraph({ children: [t((d.boCategory === "cat1" ? "\u2611" : "\u2610") + " Category 1", { bold: true }), t(" - We hereby declare that following persons/entity noted in the below table own 10%/15% or more interest or possess the right to control management/policy decisions (Refer Notes).")], spacing: { after: 70 } }),
      new Paragraph({ children: [t((d.boCategory === "cat2" ? "\u2611" : "\u2610") + " Category 2", { bold: true }), t(" - We hereby declare that no natural person is identified as per category 1 (above). (Mention the details of the natural person(s) holding the position of senior management official in the entity.)")], spacing: { after: 80 } }),
      new Paragraph({ children: [t((d.pepStatusBO === "yes" ? "\u2611" : "\u2610") + " Our personnel(s), partner(s), director(s), officer(s), or our family member(s) or our close associate(s) and beneficial owners, is a Politically Exposed Person. (\u201cPolitically Exposed Persons\u201d (PEPs) are individuals who are or have been entrusted with prominent public functions by a foreign country, including Heads of States/Governments, senior politicians, senior government or judicial or military officers, senior executives of state-owned corporations and important political party officials.)")], spacing: { after: 100 } }),
      p([t("The details of beneficial owner(s) is/are as follows:")]),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: boCols, rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [
          hCell("S.N.", boCols[0]), hCell("Name", boCols[1]), hCell("Residential Address and PIN code", boCols[2]),
          hCell("Designation", boCols[3]), hCell("DOB", boCols[4]), hCell("Proof of identity", boCols[5]),
          hCell("Proof of Address", boCols[6]), hCell("Nationality", boCols[7]), hCell("% of interest", boCols[8]),
        ]}),
        ...boRows.map((m, i) => new TableRow({ cantSplit: true, children: [
          cell(String(i + 1), boCols[0]), ucell(m.name, boCols[1]), ucell(m.address, boCols[2]),
          ucell(m.designation, boCols[3]), ucell(m.dob, boCols[4]), ucell(m.pan, boCols[5]),
          ucell(m.poa + " " + m.poaNumDisplay, boCols[6]), ucell(m.nationality, boCols[7]), ucell(m.share + "%", boCols[8]),
        ]})),
      ]}),
      p([]),
      p([t("List of the Current Partners of the firm operating at the aforementioned address:")]),
      ...d.partners.map((m, i) => p([t((i + 1) + ". "), u(m.name)])),
      p([]),
      p([t("Authorised Signatory/ies: (Refer note B for signature requirement)", { bold: true })]),
      p([t("Authorised Signatory: "), u(d.authSignatoryName)]),
      p([t("___________________________")]),
      p([t("(Name, Signature with Stamp)")]),
      new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),
      p([t("#Notes:-", { bold: true })]),
      p([t("A. RBI guidelines for identification of Beneficial owners", { bold: true })]),
      p([t("Category 1: Controlling ownership interest means:", { bold: true })]),
      new Table({ width: { size: 6500, type: WidthType.DXA }, columnWidths: [3800, 2700], rows: [
        new TableRow({ cantSplit: true, children: [hCell("Business entity", 3800), hCell("Shareholding* %", 2700)] }),
        new TableRow({ cantSplit: true, children: [cell("Partnership Firm", 3800), cell(">10%", 2700)] }),
        new TableRow({ cantSplit: true, children: [cell("Unregistered Partnership Firm", 3800), cell(">15%", 2700)] }),
      ]}),
      p([]),
      p([t("i. Ownership of/entitlement to more than 10%/15% of the capital or profits of the juridical person where the juridical person is a partnership firm, LLP. [\u2018Control\u2019 shall include the right to control the management or policy decision.]")]),
      p([t("Category 2:", { bold: true })]),
      p([t("Where no natural person is identified under (i) of category 1, the beneficial owner is the relevant natural person who holds the position of senior managing official in that entity.")]),
      p([t("B. Signature on the Declaration form:", { bold: true })]),
      p([t("A person who is authorised to sign BO declaration: Authorised signatory")]),
      p([t("C. Other Instructions", { bold: true })]),
      p([t("1. Proof of Identity -")]),
      new Table({ width: { size: 8500, type: WidthType.DXA }, columnWidths: [4000, 4500], rows: [
        new TableRow({ cantSplit: true, children: [hCell("BO Type", 4000), hCell("Details Required", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Individual (Indian / Foreign National) / Indian Entity", 4000), cell("PAN*", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Foreign entity", 4000), cell("Valid Establishment document issued in the country of incorporation/registration", 4500)] }),
      ]}),
      p([t("*If Individual PAN is not available, then form 60 should be provided.")]),
      p([t("2. Proof of Address -")]),
      new Table({ width: { size: 8500, type: WidthType.DXA }, columnWidths: [4000, 4500], rows: [
        new TableRow({ cantSplit: true, children: [hCell("BO Type", 4000), hCell("Details Required", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Individual (Indian / Foreign National)", 4000), cell("Voter ID/ Driving License / Passport/ Redacted Aadhar", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Entity (Indian or Foreign)", 4000), cell("Valid Establishment document", 4500)] }),
      ]}),
      p([t("3. PAN Number to be provided for Residents/ Entities registered in India.")]),
      p([t("4. In case if BO is a minor, and POI or POA as mentioned above is not available, then valid age proof to be provided.")]),
    ],
  }] });
  return doc;
}

// ---- COMPANY: BOARD RESOLUTION ----
async function buildCompanyResolution(d) {
  const hasLH = d.includeLetterhead && d.firmName;
  const present = d.presentPartners.length ? d.presentPartners : d.partners;
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(hasLH) },
    headers: hasLH ? { default: buildLetterheadHeader(d) } : undefined,
    children: [
      ...(!hasLH ? [p([t("(on company\u2019s letterhead)", { italics: true })], { alignment: AlignmentType.CENTER })] : []),
      p([t("TO WHOMSOEVER IT MAY CONCERN", { bold: true })], { alignment: AlignmentType.CENTER }),
      p([]),
      new Paragraph({ children: [
        t("CERTIFIED TRUE COPY OF THE RESOLUTION PASSED AT THE MEETING OF THE BOARD OF DIRECTORS OF ", { bold: true }),
        u((d.firmName||"").toUpperCase(), { bold: true }), t(" HELD ON ", { bold: true }), u(d.resolutionDate, { bold: true }),
        t(" AT ", { bold: true }), u(d.resolutionTime, { bold: true }), t(" AT "), u(d.regAddress), t("."),
      ], spacing: { after: 200 } }),
      p([]),
      p([t("PRESENT (All current Board members):", { bold: true })]),
      ...d.partners.map((m, i) => p([t((i + 1) + ". "), u(m.name), t(" \u2014 "), u(m.designation)])),
      p([]),
      new Paragraph({ children: [
        t("RESOLVED THAT ", { bold: true }), t("Mr/Mrs "), u(d.authSignatoryName), t(", with the designation of "), u(d.authSignatoryDesignation),
        t(", be and is hereby authorized, to act on behalf of the Company to execute/sign all necessary applications/documents for the purpose of opening and operating a business account with PhonePe Limited."),
      ], spacing: { after: 200 } }),
      new Paragraph({ children: [
        t("RESOLVED FURTHER THAT ", { bold: true }), t("all actions, acts, deeds, and things undertaken by the said Authorized Signatory in this capacity shall be binding upon the Company and shall remain in full force and effect until specifically rescinded by the Board of Directors."),
      ], spacing: { after: 240 } }),
      p([]),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3500, 2900, 2960], rows: [
        new TableRow({ cantSplit: true, children: [hCell("Name of the person who is authorised to sign the Board resolution.", 3500), hCell("Designation", 2900), hCell("Signature", 2960)] }),
        ...present.map((m) => new TableRow({ cantSplit: true, children: [ucell(m.name, 3500), ucell(m.designation, 2900), cell("[Sign Here]", 2960)] })),
      ]}),
      p([]),
    ],
  }] });
  return doc;
}

// ---- COMPANY: BO DECLARATION ----
async function buildCompanyBO(d) {
  const hasLH = d.includeLetterhead && d.firmName;
  const source = d.boCategory === "cat1"
    ? d.partners.filter((m) => Number(m.share) > 10)
    : d.partners.filter((m) => m.isAS);
  const boCols = [400, 1100, 1800, 680, 680, 980, 980, 680, 780];
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(hasLH) },
    headers: hasLH ? { default: buildLetterheadHeader(d) } : undefined,
    children: [
      p([t("Declarations")]),
      new Paragraph({ children: [new TextRun({ text: "A. DECLARATION OF BENEFICIAL OWNERSHIP (BO)", font: "Times New Roman", size: 26, bold: true, underline: { type: UnderlineType.SINGLE } })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: [500, 2500, 7080], rows: [
        new TableRow({ cantSplit: true, children: [cell("I", 500), cell("Name of the entity", 2500), ucell(d.firmName, 7080)] }),
        new TableRow({ cantSplit: true, children: [cell("II", 500), cell("Registered address", 2500), ucell(d.regAddress, 7080)] }),
        new TableRow({ cantSplit: true, children: [cell("III", 500), cell("Principal place of operation/office", 2500),
          new TableCell({ borders: bdrs, width: { size: 7080, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: d.principalSame === "same" ? [p([t("\u2611 Same")])] : [p([t("\u2610 Same   OR "), u(d.principalAddress)])] }) ] }),
        new TableRow({ cantSplit: true, children: [cell("IV", 500), cell("Type of entity", 2500), ucell("Company", 7080)] }),
        new TableRow({ cantSplit: true, children: [cell("V", 500), cell("Listing status", 2500), new TableCell({ borders: bdrs, width: { size: 7080, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [
          p([t((d.companyListingStatus === "listed_india" ? "\u2611" : "\u2610") + " i. An entity listed on a stock exchange in India"), ...(d.companyListingStatus === "listed_india" ? [t(" - "), u(d.stockExchangeName)] : [])]),
          p([t((d.companyListingStatus === "listed_foreign" ? "\u2611" : "\u2610") + " ii. An entity resident in a jurisdiction notified by the Central Government and listed there"), ...(d.companyListingStatus === "listed_foreign" ? [t(" - "), u(d.stockExchangeName)] : [])]),
          p([t((d.companyListingStatus === "subsidiary" ? "\u2611" : "\u2610") + " iii. A subsidiary of such listed entities (i & ii)")]),
          p([t((d.companyListingStatus === "not_listed" ? "\u2611" : "\u2610") + " B. A is not applicable; beneficial-owner information is provided below.")]),
        ]})] }),
      ]}),
      p([]),
      new Paragraph({ children: [t("The Legal Entity as stated above hereby confirms and declares the following on the date: ", { bold: true }), u(d.boDateFmt, { bold: true })], spacing: { after: 100 } }),
      new Paragraph({ children: [t((d.boCategory === "cat1" ? "\u2611" : "\u2610") + " Category 1", { bold: true }), t(" - Following persons/entities own 10% or more interest or possess the right to control management/policy decisions.")], spacing: { after: 80 } }),
      new Paragraph({ children: [t((d.boCategory === "cat2" ? "\u2611" : "\u2610") + " Category 2", { bold: true }), t(" - No natural person is identified as per Category 1 above.")], spacing: { after: 160 } }),
      p([t("The details of beneficial owner(s) is/are as follows:")]),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: boCols, rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [
          hCell("S.N.", boCols[0]), hCell("Name", boCols[1]), hCell("Residential Address and PIN code", boCols[2]),
          hCell("Designation", boCols[3]), hCell("DOB", boCols[4]), hCell("Proof of identity", boCols[5]),
          hCell("Proof of Address", boCols[6]), hCell("Nationality", boCols[7]), hCell("% of interest", boCols[8]),
        ]}),
        ...source.map((m, i) => new TableRow({ cantSplit: true, children: [
          cell(String(i + 1), boCols[0]), ucell(m.name, boCols[1]), ucell(m.address, boCols[2]),
          ucell(m.designation, boCols[3]), ucell(m.dob, boCols[4]), ucell(m.pan, boCols[5]),
          ucell(m.poa + " " + m.poaNumDisplay, boCols[6]), ucell(m.nationality, boCols[7]), ucell(m.share + "%", boCols[8]),
        ]})),
      ]}),
      p([]),
      new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),
      p([t("List of Senior Management Officials:")]),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: [5040, 5040], rows: [
        new TableRow({ cantSplit: true, children: [hCell("Name", 5040), hCell("Designation", 5040)] }),
        ...d.partners.map((m) => new TableRow({ cantSplit: true, children: [ucell(m.name, 5040), ucell(m.designation, 5040)] })),
      ]}),
      p([]),
      p([t("C. PEP Declaration (Mandatory)", { bold: true })]),
      p([t("Our personnel(s), partner(s), director(s), officer(s), or our family member(s) or our close associate(s) and beneficial owners, is a Politically Exposed Person (as defined by RBI)")]),
      p([t((d.pepStatusBO === "yes" ? "\u2611" : "\u2610") + " YES     " + (d.pepStatusBO === "no" ? "\u2611" : "\u2610") + " NO")]),
      p([t("Authorised Signatory/ies:", { bold: true })]),
      p([t("___________________________(Name, Signature with Stamp)")]),
      p([u(d.boExternalAS ? d.boExternalASName : d.authSignatoryName), t(d.boExternalAS ? " (" + d.boExternalASDesignation + ")" : "")]),
      p([t("#Notes:-", { bold: true })]),
      p([t("a. RBI guidelines for identification of Beneficial owners", { bold: true })]),
      p([t("Category 1: Controlling ownership interest means:", { bold: true })]),
      new Table({ width: { size: 7000, type: WidthType.DXA }, columnWidths: [4300, 2700], rows: [
        new TableRow({ cantSplit: true, children: [hCell("Business entity", 4300), hCell("Shareholding* %", 2700)] }),
        new TableRow({ cantSplit: true, children: [cell("Companies (Public, Private) & LLP", 4300), cell(">10%", 2700)] }),
      ]}),
      p([t("Ownership of/entitlement to more than 10% of the share or capital or profits of the juridical person, where the juridical person is a company. \u2018Control\u2019 shall include the right to appoint a majority of the directors or to control the management or policy decisions, including by virtue of shareholding, management rights, shareholders agreements or voting agreements.")]),
      p([t("Category 2: Where no natural person is identified under category 1, the beneficial owner is the relevant natural person who holds the position of senior managing official in that entity.")]),
      p([t("b. Signature on the Declaration form: person who is authorised to sign BO declaration: CS / Authorised signatory")]),
      p([t("c. Other Instructions", { bold: true })]),
      p([t("1. Proof of Identity -")]),
      new Table({ width: { size: 8500, type: WidthType.DXA }, columnWidths: [4000, 4500], rows: [
        new TableRow({ cantSplit: true, children: [hCell("BO Type", 4000), hCell("Details Required", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Individual (Indian / Foreign National) / Indian Entity", 4000), cell("PAN*", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Foreign entity", 4000), cell("Valid Establishment document issued in the country of incorporation/registration", 4500)] }),
      ]}),
      p([t("*If Individual PAN is not available, then form 60 should be provided.")]),
      p([t("2. Proof of Address -")]),
      new Table({ width: { size: 8500, type: WidthType.DXA }, columnWidths: [4000, 4500], rows: [
        new TableRow({ cantSplit: true, children: [hCell("BO Type", 4000), hCell("Details Required", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Individual (Indian / Foreign National)", 4000), cell("Voter ID/ Driving License / Passport/ Redacted Aadhar", 4500)] }),
        new TableRow({ cantSplit: true, children: [cell("Entity (Indian or Foreign)", 4000), cell("Valid Establishment document", 4500)] }),
      ]}),
      p([t("3. PAN Number to be provided for Residents/ Entities registered in India.")]),
      p([t("4. In case if BO is a minor, and POI or POA as mentioned above is not available, then valid age proof to be provided.")]),
    ],
  }] });
  return doc;
}

// ---- MDF (Merchant Declaration Form) ----
async function buildMDF(d) {
  const isACE = d.onboardingType === "ace";
  const hasLH = d.includeLetterhead && d.firmName;
  const principalText = d.principalSame === "same" ? "same (\u2611)" : d.principalAddress;
  const isCompany = d.entityType === "company";
  const isRegPartnership = d.entityType === "partnership" && d.isRegistered;
  const isUnregPartnership = d.entityType === "partnership" && !d.isRegistered;
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(hasLH) },
    headers: hasLH ? { default: buildLetterheadHeader(d) } : undefined,
    children: [
      p([t("Merchant Declaration", { bold: true })], { alignment: AlignmentType.CENTER }),
      p([t("To,")]),
      p([t("PhonePe Limited (Formerly known as 'PhonePe Private Limited') (hereinafter referred as \u201cPhonePe\u201d)")]),
      p([t("Office-2, Floor 5, Wing A, Block A, Salarpuria Softzone, Bellandur Village, Varthur Hobli, Outer Ring Road, Bangalore South, Bangalore, Karnataka, India, 560103")]),
      p([]),
      p([t("Subject: PhonePe Merchant Declaration", { bold: true })]),
      p([]),
      new Paragraph({ children: [
        t("I, "), u(d.mdfAuthName), t(", hereinafter referred to as \u201cMerchant\u201d being the "), u(d.mdfAuthDesignation),
        t(" of "), u(d.firmName), t(" having its registered office address at "), u(d.regAddress),
        t(" (\u201cEntity\u201d) and having its principal place of operation/office at "), u(principalText),
        t(", do hereby declare that I have been authorised to act as a designated authorised signatory for the Entity (including, but not limited to, registration/execution/renewal/amendment of the business related association(s)/partnership(s)/contract(s)/terms and conditions with PhonePe) and that the below mentioned details provided by me (including my specimen signature) are true, accurate, valid, legally binding and authenticated for the Entity, and can be used for the purposes of obtaining payment facilitation services, business related associations/partnership(s) with PhonePe."),
      ], spacing: { after: 200 } }),
      p([t("I hereby allow PhonePe to collect, store and use my KYC and/or other details as required by PhonePe, for the purposes of verifying my identity as the authorised signatory of the entity thereby enabling the entity to be onboarded as Merchant with PhonePe for availing PhonePe services, in accordance with PhonePe\u2019s Terms and Conditions and Privacy Policy.")]),
      p([t("Details provided under this declaration:", { bold: true })]),
      p([t("1. Mobile No. (registered with PhonePe for Onboarding): "), u(d.mdfMobile)]),
      p([t("2. Email ID (registered with PhonePe for Onboarding): "), u(d.mdfEmail)]),
      ...( !isACE ? [p([t("3. Individual KYC Documents:")])] : []),
      p([t((isACE ? "3" : "4") + ". In case of Person with Disability (PwD), please specify")]),
      p([t("   Type of Disability: "), u(d.mdfPwd === "yes" ? d.mdfPwdType : "N/A")]),
      p([t("   Percentage of Disability: "), u(d.mdfPwd === "yes" ? d.mdfPwdPct + "%" : "N/A")]),
      ...( !isACE ? [
        p([t("5. Father\u2019s Name (of the Authorized Signatory): "), u(d.mdfFatherName)]),
        new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),
        p([t("PAN CARD (Mandatory)                                                     \u2611")]),
        p([t("Any one of the following is mandatory (Please tick whichever submitted):")]),
        p([t((d.mdfKycDoc === "aadhaar" ? "\u2611" : "\u2610") + " Aadhaar (masked except the last 4 digits)")]),
        p([t((d.mdfKycDoc === "dl" ? "\u2611" : "\u2610") + " Driving License")]),
        p([t((d.mdfKycDoc === "voterid" ? "\u2611" : "\u2610") + " Voter ID")]),
        p([t("I hereby declare that the above information/details provided herein are true, valid and accurate as on date of submission and further that I would be liable for any incorrect/false information or for any untrue statement of details/information provided.")]),
        p([t("Signature with seal: ____________________      Name: "), u(d.mdfAuthName)]),
      ] : []),
      p([]),
      ...(isACE ? [new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } })] : []),
      p([t("I, on behalf of the Merchant, further declare that:", { bold: true })]),
      ...( !isACE ? [
        p([t((isCompany ? "\u2611" : "\u2610") + " Company    " + (isRegPartnership ? "\u2611" : "\u2610") + " Registered Partnership Firm    " + (isUnregPartnership ? "\u2611" : "\u2610") + " Un-Registered Partnership Firm")]),
      ] : []),
      new Paragraph({ children: [
        t((d.mdfTanStatus === "has_tan" ? "\u2611" : "\u2610") + " 1. The Merchant is registered under the Income Tax Act, 1961 and has obtained TAN Number "),
        ...(d.mdfTanStatus === "has_tan" ? [u(d.mdfTanNum)] : [t("_________________________")]),
        t(" against the registration. OR\n" + (d.mdfTanStatus === "no_tan" ? "\u2611" : "\u2610") + " The Merchant does not hold TAN as it is not liable to deduct tax at source or collect tax at source as per the provisions of Income Tax Act, 1961."),
      ], spacing: { after: 120 } }),
      ...( !isACE ? [ new Paragraph({ children: [
        t("2. The Merchant is registered and a GSTIN certificate/acknowledgement having provisional number "), u(d.mdfGstStatus === "has_gst" ? d.mdfGstNum : "N/A"),
        t(" is issued by GST authorities. OR The Merchant does not have any registration with GST authorities (" + (d.mdfGstStatus === "no_gst" ? "\u2611" : "\u2610") + ")."),
      ], spacing: { after: 120 } }) ] : []),
      ...( !isACE ? [
        new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),
        p([t("The entity is working in the nature of:")]),
        p([t((d.mdfEntityNature === "government" ? "\u2611" : "\u2610") + " Government organization")]),
        p([t((d.mdfEntityNature === "ngo" ? "\u2611" : "\u2610") + " NGO/Charitable institution")]),
        p([t((d.mdfEntityNature === "na" ? "\u2611" : "\u2610") + " NA")]),
      ] : []),
      new Paragraph({ children: [
        t("No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a Politically Exposed Person (PEP): "),
        t((d.mdfPepStatus === "yes" ? "YES \u2611" : "YES \u2610") + "     " + (d.mdfPepStatus === "no" ? "NO \u2611" : "NO \u2610")),
      ], spacing: { after: 160 } }),
      new Paragraph({ children: [t("I, having PAN number "), u(d.mdfAuthPan), t(", hereby declare that the above facts and information are true, complete and correct to the best of my knowledge. I understand and agree that in case it is found that the above-mentioned facts and information are incorrect, I will be personally held liable for the same.")], spacing: { after: 300 } }),
      p([t("Yours faithfully, For and behalf of the Merchant")]),
      p([]), p([]),
      p([t("___________________ (Signature with Seal)")]),
      p([t("Designation: "), u(d.mdfAuthDesignation)]),
      p([t("Date: "), u(d.mdfDate), t("   Place: "), u(d.mdfPlace)]),
      ...( !isACE ? [
        p([]),
        p([t("Picture of the Authorised Signatory (Countersign with face visible)", { italics: true })], { alignment: AlignmentType.RIGHT }),
        p([t("Note:", { bold: true })]),
        p([t("1. \u2018Government company\u2019 means any company in which not less than fifty-one per cent of the paid-up share capital is held by the Central Government, or by any State Government or Governments, or partly by the Central Government and partly by one or more State Governments, and includes a company which is a subsidiary company of such a Government company.")]),
        p([t("2. NGO (For Darpan applicability): For Company, a Section 8 company can be identified from its COI. For Societies and Trusts, non-government societies and trusts incorporated under applicable legislation or non-profit academic institutions may be classified as NPO.")]),
      ] : []),
    ],
  }] });
  return doc;
}

// ---- GENERATE HANDLERS ----
async function buildDocxArtifact(kind) {
  computeDerived();
  const fn = (formData.firmName || "document").replace(/[^a-zA-Z0-9]/g, "_");
  let doc, suffix;
  if (kind === "resolution") { doc = state.entityType === "company" ? await buildCompanyResolution(formData) : await buildPartnershipResolution(formData); suffix = "_Resolution.docx"; }
  else if (kind === "bo") { doc = state.entityType === "company" ? await buildCompanyBO(formData) : await buildPartnershipBO(formData); suffix = "_BO_Declaration.docx"; }
  else if (kind === "mdf") { doc = await buildMDF(formData); suffix = "_MDF.docx"; }
  return { blob: await Packer.toBlob(doc), filename: fn + suffix };
}
async function generateDocx(kind) {
  try {
    const artifact = await buildDocxArtifact(kind);
    logToSheet();
    downloadBlob(artifact.blob, artifact.filename);
    showToast("Download started", "ok");
  } catch (e) { console.error(e); showToast("Error: " + e.message, "er"); }
}
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.style.display = "none"; a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
}

let previewArtifact = null;
let previewLibPromise = null;
let pdfLibPromise = null;
function loadPreviewLibs() {
  if (previewLibPromise) return previewLibPromise;
  const load = (src) => new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src; script.onload = resolve; script.onerror = reject;
    document.head.appendChild(script);
  });
  previewLibPromise = (async () => {
    if (!window.JSZip) await load("./vendor/jszip-3.10.1.min.js");
    if (!window.docx) await load("./vendor/docx-preview-0.4.0.min.js");
    if (!window.docx || !window.docx.renderAsync) throw new Error("Word preview library did not load");
  })();
  return previewLibPromise;
}
function loadPdfLib() {
  if (pdfLibPromise) return pdfLibPromise;
  pdfLibPromise = new Promise((resolve, reject) => {
    if (window.html2pdf) { resolve(); return; }
    const script = document.createElement("script");
    script.src = "./vendor/html2pdf-0.10.2.bundle.min.js";
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
  return pdfLibPromise;
}
async function previewDocx(kind) {
  const modal = document.getElementById("previewModal");
  const body = document.getElementById("previewBody");
  document.getElementById("previewTitle").textContent = kind === "bo" ? "BO Declaration preview" : kind === "mdf" ? "MDF preview" : "Resolution preview";
  body.innerHTML = '<div class="preview-loading">Preparing Word preview…</div>';
  modal.classList.add("open");
  try {
    const [artifact] = await Promise.all([buildDocxArtifact(kind), loadPreviewLibs()]);
    previewArtifact = artifact;
    body.innerHTML = "";
    await window.docx.renderAsync(artifact.blob, body, null, { className: "docx", inWrapper: true, ignoreWidth: false, ignoreHeight: false, breakPages: true, renderHeaders: true, renderFooters: true });
    return true;
  } catch (e) {
    console.error(e);
    body.innerHTML = '<div class="error-box">Preview unavailable: '+esc(e.message)+'</div>';
    return false;
  }
}
function closePreview() {
  document.getElementById("previewModal").classList.remove("open");
  document.getElementById("previewBody").innerHTML = '<div class="preview-loading">Preparing preview…</div>';
  previewArtifact = null;
}
async function generatePrint(kind) {
  if (await previewDocx(kind)) {
    logToSheet();
    window.print();
  }
}
async function downloadPdf(kind) {
  if (!previewArtifact || !document.getElementById("previewModal").classList.contains("open")) {
    if (!await previewDocx(kind)) return;
  }
  try {
    await loadPdfLib();
    const pages = document.querySelectorAll("#previewBody section.docx");
    if (!pages.length) throw new Error("No preview pages found");
    const pdfName = previewArtifact.filename.replace(/\.docx$/i, ".pdf");
    const common = { margin: 0, image: { type: "jpeg", quality: 0.99 }, html2canvas: { scale: 2.2, useCORS: true, backgroundColor: "#ffffff", logging: false }, jsPDF: { unit: "mm", format: "a4", orientation: "portrait" } };
    const canvases = [];
    const exportRoot = document.createElement("div");
    exportRoot.className = "pdf-export-root";
    document.body.appendChild(exportRoot);
    for (const page of pages) {
      const clone = page.cloneNode(true);
      clone.style.position = "relative";
      clone.style.top = "0";
      clone.style.left = "0";
      clone.style.margin = "0";
      clone.style.transform = "none";
      exportRoot.appendChild(clone);
      canvases.push(await window.html2pdf().set(common).from(clone).toCanvas().get("canvas"));
      clone.remove();
    }
    const firstClone = pages[0].cloneNode(true);
    firstClone.style.cssText += ";position:relative;top:0;left:0;margin:0;transform:none";
    exportRoot.appendChild(firstClone);
    const firstWorker = window.html2pdf().set({ ...common, filename: pdfName }).from(firstClone).toPdf();
    const pdf = await firstWorker.get("pdf");
    // Replace html2pdf's first-page image with a page-fitted render, then add
    // one image per DOCX page so no table can be split by automatic pagination.
    for (let index = 0; index < canvases.length; index++) {
      const canvas = canvases[index];
      const image = canvas.toDataURL("image/jpeg", 0.99);
      const ratioHeight = 210 * canvas.height / canvas.width;
      const drawHeight = Math.min(297, ratioHeight);
      const drawWidth = drawHeight === 297 ? 297 * canvas.width / canvas.height : 210;
      const x = (210 - drawWidth) / 2;
      if (index === 0) {
        pdf.deletePage(1);
        pdf.addPage("a4", "portrait");
      } else {
        pdf.addPage("a4", "portrait");
      }
      pdf.addImage(image, "JPEG", x, 0, drawWidth, drawHeight);
    }
    if (pdf.getNumberOfPages() > canvases.length) pdf.deletePage(1);
    await pdf.save(pdfName);
    exportRoot.remove();
    logToSheet();
    showToast("PDF download started", "ok");
  } catch (error) {
    console.error(error);
    showToast("PDF error: " + error.message, "er");
  }
}

// ============================================================
// UI RENDERING
// ============================================================
function renderProgress() {
  const bar = document.getElementById("progressBar");
  if (!bar) return;
  const total = 5;
  let segs = "";
  for (let i = 0; i < total; i++) {
    const cls = i < state.step ? "done" : (i === state.step ? "cur" : "");
    segs += '<div class="seg ' + cls + '"></div>';
  }
  bar.innerHTML = segs;
}

function PartnerCardHTML(partner, index, fullKYC) {
  const err = (f) => state.inlineErrors[f + "_" + partner.id] || "";
  const label = state.entityType === "company" ? "Board Member" : "Partner";
  const isAadhaar = partner.poa === "AADHAAR";
  return '<div class="prow ' + (partner.isAS ? "as" : "") + '">' +
    '<div class="prow-hd"><span class="plbl">' + label + " " + (index + 1) + '</span>' +
    '<label class="as-label"><input type="radio" name="authSignatory" ' + (partner.isAS ? "checked" : "") + " " + on("change", () => onSelectAS(partner.id)) + " /> \u2705 Authorised Signatory</label>" +
    '<button type="button" class="rm-btn" ' + on("click", () => removePartner(partner.id)) + ">Remove</button></div>" +
    '<div class="' + (fullKYC ? "g3" : "g2") + '" style="margin-bottom:10px">' +
      '<div class="f"><label>Full Name *</label><input class="' + (err("name") ? "err" : "") + '" data-fid="name_' + partner.id + '" value="' + attr(partner.name) + '" ' + on("input", (e) => onPartnerChange(partner.id, "name", e.target.value)) + " /></div>" +
      (state.entityType === "company"
        ? '<div class="f"><label>Designation *</label><select ' + on("change", (e) => { onPartnerChange(partner.id, "designation", e.target.value); rerender(); }) + ">" + option("Director","Director",partner.designation) + option("Company Secretary","Company Secretary",partner.designation) + option("CEO","CEO",partner.designation) + option("Board Chairman","Board Chairman",partner.designation) + option("Managing Director","Managing Director",partner.designation) + "</select></div>"
        : '<div class="f"><label>Designation</label><input value="Partner" disabled /></div>') +
      (fullKYC ? (
        '<div class="f"><label>PAN *</label><input class="' + (err("pan") ? "err" : "") + '" maxlength="10" value="' + attr(partner.pan) + '" ' + on("input", (e) => onPartnerChange(partner.id, "pan", e.target.value.toUpperCase())) + " /></div>" +
        '<div class="f"><label>DOB *</label><input type="date" class="' + (err("dobRaw") ? "err" : "") + '" value="' + attr(partner.dobRaw) + '" ' + on("input", (e) => onPartnerChange(partner.id, "dobRaw", e.target.value)) + " /></div>" +
        '<div class="f s2 s3"><label>Residential Address & PIN *</label><input class="' + (err("address") ? "err" : "") + '" value="' + attr(partner.address) + '" ' + on("input", (e) => onPartnerChange(partner.id, "address", e.target.value)) + " /></div>" +
        '<div class="f"><label>Proof Type *</label><select ' + on("change", (e) => { onPartnerChange(partner.id, "poa", e.target.value); onPartnerChange(partner.id, "poaNum", ""); rerender(); }) + ">" + option("AADHAAR","Aadhaar",partner.poa) + option("VOTER ID","Voter ID",partner.poa) + option("DRIVING LICENSE","Driving License",partner.poa) + option("PASSPORT","Passport",partner.poa) + "</select></div>" +
        '<div class="f"><label>Proof Number *</label><div class="aadhar-wrap ' + (err("poaNum") ? "err" : "") + '"><input class="aadhar-input" placeholder="' + (isAadhaar ? "Last 4 digits" : "Reference number") + '" maxlength="' + (isAadhaar ? "4" : "30") + '" value="' + attr(partner.poaNum) + '" ' + on("input", (e) => { if (isAadhaar) e.target.value = e.target.value.replace(/\\D/g,""); onPartnerChange(partner.id, "poaNum", e.target.value); }) + " /></div></div>" +
        '<div class="f"><label>Nationality *</label><select ' + on("change", (e) => onPartnerChange(partner.id, "nationality", e.target.value)) + ">" + option("Indian","Indian",partner.nationality) + option("Foreign National","Foreign National",partner.nationality) + "</select></div>"
      ) : "") +
      (needsOwnership(state.docRequirement) ? '<div class="f"><label>% Ownership *</label><input class="' + (err("share") ? "err" : "") + '" inputmode="decimal" value="' + attr(partner.share) + '" ' + on("input", (e) => onPartnerChange(partner.id, "share", sanitizeShare(e.target.value))) + ' /><span class="hint">Refer Partnership Deed for Shareholding %</span></div>' : '') +
    "</div></div>";
}

function resolutionDeclarationsHTML(errs) {
  if (state.entityType !== "partnership") return "";
  return '<div class="divider">Partner Resolution Format *</div>' +
    '<div class="rg version-options"><label><input type="radio" name="resolutionVersion" ' + (state.resolutionVersion === "v5a" ? "checked" : "") + " " + on("change", () => { state.resolutionVersion = "v5a"; rerender(); }) + '> <span><strong>Version 5A (Recommended)</strong><small>Use when all partners are natural persons.</small></span></label>' +
    '<label><input type="radio" name="resolutionVersion" ' + (state.resolutionVersion === "v5b" ? "checked" : "") + " " + on("change", () => { state.resolutionVersion = "v5b"; rerender(); }) + '> <span><strong>Version 5B</strong><small>Valid when one or more partners are non-natural persons.</small></span></label></div>' +
    (errs.resolutionVersion ? '<span class="err-msg">' + esc(errs.resolutionVersion) + '</span>' : '') +
    '<div class="f" style="margin-top:10px"><label>Is any covered person a Politically Exposed Person (PEP)? *</label><div class="rg">' +
    '<label><input type="radio" name="pepRes" ' + (state.pepStatusRes === "yes" ? "checked" : "") + " " + on("change", () => { state.pepStatusRes = "yes"; rerender(); }) + '> Yes</label>' +
    '<label><input type="radio" name="pepRes" ' + (state.pepStatusRes === "no" ? "checked" : "") + " " + on("change", () => { state.pepStatusRes = "no"; rerender(); }) + '> No</label></div>' +
    (errs.pepStatusRes ? '<span class="err-msg">' + esc(errs.pepStatusRes) + '</span>' : '') + '</div>';
}

function boPepHTML(errs) {
  return '<div class="f" style="margin-top:10px"><label>Is any covered person a Politically Exposed Person (PEP)? *</label><div class="rg">' +
    '<label><input type="radio" name="pepBO" ' + (state.pepStatusBO === "yes" ? "checked" : "") + " " + on("change", () => { state.pepStatusBO = "yes"; rerender(); }) + '> Yes</label>' +
    '<label><input type="radio" name="pepBO" ' + (state.pepStatusBO === "no" ? "checked" : "") + " " + on("change", () => { state.pepStatusBO = "no"; rerender(); }) + '> No</label></div>' +
    (errs.pepStatusBO ? '<span class="err-msg">' + esc(errs.pepStatusBO) + '</span>' : '') + '</div>';
}
function signingGuidanceHTML(errs) {
  if (state.entityType !== "partnership") return "";
  const share = state.partners.filter((partner) => state.presentPartnerIds.includes(partner.id)).reduce((sum, partner) => sum + Number(partner.share || 0), 0);
  return '<div class="education-box"><strong>Partner Resolution signing guidance</strong><br>1. Refer Partnership Deed for Shareholding %.<br>2. After filling the document: Download &gt; Print &gt; Upload in Owner KYC section &gt; Partner Resolution.<br><strong>Selected signing ownership: ' + share.toFixed(1) + '% (minimum 50% required)</strong></div>' + (errs.presentShare ? '<div class="error-box">' + esc(errs.presentShare) + '</div>' : '');
}

function pageActions(back, next) {
  return '<div class="act">' + (back || '<div></div>') + '<div class="reset-actions"><button class="btn btn-reset" ' + on("click", resetPage) + '>Reset page</button><button class="btn btn-danger" ' + on("click", globalReset) + '>Global reset</button>' + (next || '') + '</div></div>';
}
function documentOptionHTML(opt) {
  const mandatory = (state.onboardingType === "salesforce") || ["set1","set5","set6","set7"].includes(opt.id);
  const status = mandatory ? '<span class="doc-status mandatory">Mandatory</span>' : '<span class="doc-status optional">Optional</span>';
  return '<label><input type="radio" name="docReq" ' + (state.docRequirement === opt.id ? "checked" : "") + " " + on("change", () => { state.docRequirement = opt.id; state.formValidated = false; rerender(); }) + ' /><span>' + opt.label + '</span>' + status + '</label>';
}

function renderApp() {
  computeDerived();
  const errs = state.inlineErrors;

  // ---- STEP 0: Agent + Merchant identification ----
  if (state.step === 0) {
    return '<div class="card"><div class="chd"><h2>\ud83d\udcf1 Agent & Merchant Details</h2><button class="reset-link" ' + on("click", () => resetSection("tracking")) + '>Reset section</button><span class="badge">Step 1 of 5</span></div><div class="cbd">' +
      '<div class="info-blue">Your mobile number is saved permanently on this device \u2014 you will not need to re-enter it next time. Merchant details reset for every new onboarding.</div>' +
      '<div class="g2">' +
        '<div class="f s2"><label>Your Mobile Number (Agent) *</label><input type="tel" inputmode="numeric" maxlength="10" class="' + (errs.agentMobile ? "err" : "") + '" value="' + attr(state.agentMobile) + '" ' + on("input", (e) => { state.agentMobile = e.target.value.replace(/\\D/g,"").slice(0,10); scheduleSave(); }) + ' />' + (errs.agentMobile ? '<span class="err-msg">' + errs.agentMobile + "</span>" : "") + "</div>" +
        '<div class="f s2"><label>Is this a new or existing merchant?</label><div class="rg">' +
          '<label><input type="radio" name="merchantStatus" ' + (state.merchantStatus === "new" ? "checked" : "") + " " + on("change", () => { state.merchantStatus = "new"; rerender(); }) + " /> New Merchant</label>" +
          '<label><input type="radio" name="merchantStatus" ' + (state.merchantStatus === "existing" ? "checked" : "") + " " + on("change", () => { state.merchantStatus = "existing"; rerender(); }) + " /> Existing Merchant</label>" +
        "</div></div>" +
        (state.merchantStatus === "new"
          ? '<div class="f s2"><label>Merchant Mobile Number *</label><input type="tel" inputmode="numeric" maxlength="10" class="' + (errs.merchantMobile ? "err" : "") + '" value="' + attr(state.merchantMobile) + '" ' + on("input", (e) => { state.merchantMobile = e.target.value.replace(/\D/g,"").slice(0,10); scheduleSave(); }) + ' />' + (errs.merchantMobile ? '<span class="err-msg">' + errs.merchantMobile + "</span>" : "") + "</div>"
          : '<div class="f s2"><label>Merchant ID *</label><input class="' + (errs.merchantId ? "err" : "") + '" value="' + attr(state.merchantId) + '" ' + on("input", (e) => { state.merchantId = e.target.value; scheduleSave(); }) + ' />' + (errs.merchantId ? '<span class="err-msg">' + errs.merchantId + "</span>" : "") + "</div>") +
      "</div></div></div>" +
      pageActions('', '<button class="btn btn-p" ' + on("click", () => { if (validateStep0()) { state.step = 1; rerender(); } }) + '>Next: Platform & Entity \u2192</button>');
  }

  // ---- STEP 1: Platform + Entity ----
  if (state.step === 1) {
    return '<div class="card"><div class="chd"><h2>\ud83c\udfe2 Onboarding Platform & Entity Type</h2><button class="reset-link" ' + on("click", () => resetSection("platform")) + '>Reset section</button><span class="badge">Step 2 of 5</span></div><div class="cbd">' +
      '<div class="divider">Onboarding Platform</div><div class="rg">' +
        '<label><input type="radio" name="platform" ' + (state.onboardingType === "ace" ? "checked" : "") + " " + on("change", () => { state.onboardingType = "ace"; state.docRequirement = null; rerender(); }) + " /> ACE</label>" +
        '<label><input type="radio" name="platform" ' + (state.onboardingType === "salesforce" ? "checked" : "") + " " + on("change", () => { state.onboardingType = "salesforce"; state.docRequirement = null; state.mdfPepStatus = "no"; state.mdfAuthDesignation = state.entityType === "company" ? "Director" : "Partner"; rerender(); }) + " /> Salesforce</label>" +
      "</div>" +
      '<div class="divider">Entity Type</div><div class="rg">' +
        '<label><input type="radio" name="entity" ' + (state.entityType === "partnership" ? "checked" : "") + " " + on("change", () => { state.entityType = "partnership"; state.docRequirement = null; state.mdfAuthDesignation = "Partner"; syncDefaultDesignations(); rerender(); }) + " /> Partnership Firm</label>" +
        '<label><input type="radio" name="entity" ' + (state.entityType === "company" ? "checked" : "") + " " + on("change", () => { state.entityType = "company"; state.docRequirement = null; state.mdfAuthDesignation = "Director"; syncDefaultDesignations(); rerender(); }) + " /> Company</label>" +
      "</div></div></div>" +
      pageActions('<button class="btn btn-s" ' + on("click", () => { state.step = 0; rerender(); }) + '>\u2190 Back</button>', '<button class="btn btn-p" ' + (state.onboardingType && state.entityType ? "" : "disabled") + " " + on("click", () => { if (state.onboardingType && state.entityType) { state.step = 2; rerender(); } }) + '>Next: Document Set \u2192</button>');
  }

  // ---- STEP 2: Document requirement ----
  if (state.step === 2) {
    const options = (state.onboardingType && state.entityType) ? docOptionsMap[state.onboardingType][state.entityType] : [];
    return '<div class="card"><div class="chd"><h2>\ud83d\udcc2 Which documents do you need?</h2><button class="reset-link" ' + on("click", () => resetSection("documents")) + '>Reset section</button><span class="badge">Step 3 of 5</span></div><div class="cbd">' +
      (state.onboardingType === "salesforce" && state.entityType === "partnership" ? '<div class="info-blue"><strong>MDF is mandatory for Salesforce Partnership onboarding.</strong> All requested individual and combined choices are available below.</div>' : state.onboardingType === "salesforce" && state.entityType === "company" ? '<div class="info-blue"><strong>Board Resolution, BO Declaration and MDF are mandatory for Salesforce Company onboarding.</strong> All requested choices are available below.</div>' : state.onboardingType === "ace" && state.entityType === "company" ? '<div class="info-blue"><strong>Board Resolution, BO Declaration, and Board Resolution + BO Declaration are marked mandatory.</strong> The ACE OSV combination is optional.</div>' : state.entityType === "partnership" ? '<div class="info-blue"><strong>Partner Resolution is mandatory for Partnership onboarding.</strong> Other documents are optional based on the onboarding requirement.</div>' : '') +
      '<div class="rg doc-options" style="flex-direction:column">' + options.map(documentOptionHTML).join("") + "</div></div></div>" +
      pageActions('<button class="btn btn-s" ' + on("click", () => { state.step = 1; rerender(); }) + '>\u2190 Back</button>', '<button class="btn btn-p" ' + (state.docRequirement ? "" : "disabled") + " " + on("click", () => { if (state.docRequirement) { state.step = 3; rerender(); } }) + '>Next: Fill Details \u2192</button>');
  }

  // ---- STEP 3: Data entry (firm, letterhead, partners, resolution, BO, MDF) ----
  if (state.step === 3) {
    const isFullKyc = needsFullKYC(state.docRequirement);
    const docLabel = (docOptionsMap[state.onboardingType][state.entityType].find(o => o.id === state.docRequirement) || {}).label || "";
    return (
      '<div class="card"><div class="chd"><h2>\ud83c\udfe2 Entity Details</h2><button class="reset-link" ' + on("click", () => resetSection("entity")) + '>Reset section</button><span class="badge">' + esc(docLabel) + '</span></div><div class="cbd">' +
        '<div class="info">This information is captured once and reused across every document you generate for this merchant.</div>' +
        '<div class="g2">' +
          '<div class="f s2"><label>Legal Entity / Firm Name (as on PAN) *</label><input class="' + (errs.firmName ? "err" : "") + '" value="' + attr(state.firmName) + '" ' + on("input", (e) => { state.firmName = e.target.value; scheduleSave(); }) + " /><span class=\"hint\">Must match Business PAN Card exactly.</span></div>" +
          '<div class="f s2"><label>Registered Office Address / Main Office Address *</label><textarea class="' + (errs.regAddress ? "err" : "") + '" ' + on("input", (e) => { state.regAddress = e.target.value; scheduleSave(); }) + ">" + attr(state.regAddress) + "</textarea></div>" +
          (state.entityType === "partnership" ? (
            '<div class="f"><label>Entity Registration *</label><select class="' + (errs.partnershipRegType ? "err" : "") + '" ' + on("change", (e) => { state.partnershipRegType = e.target.value; rerender(); }) + ">" + option("","Select registration type",state.partnershipRegType) + option("registered","LLP / Registered Partnership",state.partnershipRegType) + option("unregistered","Non LLP / Unregistered Partnership",state.partnershipRegType) + "</select><span class=\"hint\">Registered: LLP or deed registered in Registrar office. Unregistered: normal deed, notarized or not notarized.</span></div>"
          ) : "") +
          '<div class="f"><label>Principal Place of Operation *</label><div class="rg">' +
            '<label><input type="radio" name="pSame" ' + (state.principalSame === "same" ? "checked" : "") + " " + on("change", () => { state.principalSame = "same"; rerender(); }) + " /> Same</label>" +
            '<label><input type="radio" name="pSame" ' + (state.principalSame === "diff" ? "checked" : "") + " " + on("change", () => { state.principalSame = "diff"; rerender(); }) + " /> Different</label>" +
          "</div>" + (state.principalSame === "diff" ? '<input style="margin-top:8px" class="' + (errs.principalAddress ? "err" : "") + '" value="' + attr(state.principalAddress) + '" ' + on("input", (e) => { state.principalAddress = e.target.value; scheduleSave(); }) + " />" : "") + "</div>" +
          (state.entityType === "partnership" ? '<div class="f"><label>Partnership Deed Date</label><input type="date" value="' + attr(state.deedDate) + '" ' + on("change", (e) => { state.deedDate = e.target.value; scheduleSave(); }) + " /></div>" : "") +
        "</div>" +
        '<div class="divider">Letterhead (optional \u2014 embedded in generated documents)</div>' +
        '<div class="cg" style="margin-bottom:10px"><label><input type="checkbox" ' + (state.includeLetterhead ? "checked" : "") + " " + on("change", (e) => { state.includeLetterhead = e.target.checked; rerender(); }) + " /> Embed letterhead in generated documents</label></div>" +
        (state.includeLetterhead ? '<div class="info-blue">Letterhead Firm Name: <strong>' + esc(state.firmName || "Enter the Legal Entity / Firm Name above") + '</strong></div>' : '') +
      "</div></div>" +

      // Partners / Directors
      (needsResolution(state.docRequirement) || needsBO(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udc65 ' + (state.entityType === "company" ? "Directors" : "Partners") + '</h2><button class="reset-link" ' + on("click", () => resetSection("members")) + ' >Reset section</button></div><div class="cbd">' +
          '<div class="info">' + (isFullKyc ? "Full KYC required: PAN, DOB, Address, Proof of Address, % ownership. Total ownership must equal 100%." : "Basic details only for this document set.") + "</div>" +
          state.partners.map((pt, idx) => PartnerCardHTML(pt, idx, isFullKyc)).join("") +
          (errs.shareTotal ? '<div class="error-box">' + esc(errs.shareTotal) + "</div>" : "") +
          (errs.authSignatory ? '<div class="error-box">' + esc(errs.authSignatory) + "</div>" : "") +
          '<button type="button" class="add-btn" ' + on("click", addPartner) + ">\uff0b Add Another " + (state.entityType === "company" ? "Director" : "Partner") + "</button>" +
        "</div></div>"
      ) : "") +

      // Resolution details
      (needsResolution(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udccb ' + (state.entityType === "partnership" ? "Partner Resolution" : "Board Resolution") + ' Details</h2><button class="reset-link" ' + on("click", () => resetSection("resolution")) + '>Reset section</button></div><div class="cbd">' +
          '<div class="g2">' +
            '<div class="f"><label>Meeting Date *</label><input type="date" class="' + (errs.resolutionDateRaw ? "err" : "") + '" value="' + attr(state.resolutionDateRaw) + '" ' + on("change", (e) => { state.resolutionDateRaw = e.target.value; state.boDate = state.boDate || e.target.value; scheduleSave(); }) + " /></div>" +
            (state.entityType === "company" ? '<div class="f"><label>Meeting Time *</label><input type="time" class="' + (errs.resolutionTimeRaw ? "err" : "") + '" value="' + attr(state.resolutionTimeRaw) + '" ' + on("change", (e) => { state.resolutionTimeRaw = e.target.value; scheduleSave(); }) + " /></div>" : "") +
          "</div>" +
          '<div class="act" style="justify-content:flex-start"><button type="button" class="btn btn-s" ' + on("click", prefillDateTime) + ">Fill current " + (state.entityType === "company" ? "date/time" : "date") + "</button></div>" +
          '<div class="divider">Members Present & Signing</div>' +
          (errs.presentPartnerIds ? '<div class="error-box">' + esc(errs.presentPartnerIds) + "</div>" : "") +
          state.partners.map((pt) => '<label class="rg" style="justify-content:flex-start"><input type="checkbox" ' + (state.presentPartnerIds.includes(pt.id) ? "checked" : "") + " " + on("change", (e) => togglePresentPartner(pt.id, e.target.checked)) + " /> " + esc(pt.name || "(unnamed)") + (pt.isAS ? " \u2705" : "") + "</label>").join("") +
          signingGuidanceHTML(errs) +
          (state.entityType === "company" ? '<div class="education-box"><strong>Board Resolution signing:</strong> Select at least one person who will sign. Total ownership across all listed directors/officials must equal 100%.</div><div class="cg"><label><input type="checkbox" ' + (state.isOPC ? "checked" : "") + ' ' + on("change", (e) => { state.isOPC = e.target.checked; rerender(); }) + ' /> This is a One Person Company (allows one listed director)</label></div>' + (errs.companySigners ? '<div class="error-box">'+esc(errs.companySigners)+'</div>' : '') : '') +
          resolutionDeclarationsHTML(errs) +
        "</div></div>"
      ) : "") +

      // BO details
      (needsBO(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udcdc BO Declaration Details</h2><button class="reset-link" ' + on("click", () => resetSection("bo")) + '>Reset section</button></div><div class="cbd">' +
          '<div class="g2">' +
            '<div class="f"><label>BO Declaration Date *</label><input type="date" class="' + (errs.boDate ? "err" : "") + '" value="' + attr(state.boDate) + '" ' + on("change", (e) => { state.boDate = e.target.value; scheduleSave(); }) + " /></div>" +
          "</div>" +
          '<div class="divider">Category</div><div class="rg">' +
            '<label><input type="radio" name="boCat" ' + (state.boCategory === "cat1" ? "checked" : "") + " " + on("change", () => { state.boCategory = "cat1"; rerender(); }) + " /> Category 1 (natural persons above threshold)</label>" +
            '<label><input type="radio" name="boCat" ' + (state.boCategory === "cat2" ? "checked" : "") + " " + on("change", () => { state.boCategory = "cat2"; rerender(); }) + " /> Category 2 (senior managing official)</label>" +
          "</div>" + (errs.boCategory ? '<span class="err-msg">'+esc(errs.boCategory)+'</span>' : '') +
          (state.entityType === "company" ? '<div class="divider">Company Listing Status</div><div class="f"><select ' + on("change", (e) => { state.companyListingStatus = e.target.value; rerender(); }) + '>' + option("not_listed","Not listed / BO details applicable",state.companyListingStatus) + option("listed_india","Listed on an Indian stock exchange",state.companyListingStatus) + option("listed_foreign","Listed in a notified foreign jurisdiction",state.companyListingStatus) + option("subsidiary","Subsidiary of a listed entity",state.companyListingStatus) + '</select></div>' + (state.companyListingStatus !== "not_listed" ? '<div class="f" style="margin-top:8px"><label>Stock Exchange Name *</label><input class="'+(errs.stockExchangeName?"err":"")+'" value="'+attr(state.stockExchangeName)+'" '+on("input",(e)=>{state.stockExchangeName=e.target.value;scheduleSave();})+' /></div>' : '') : '') +
          boPepHTML(errs) +
          (state.entityType === "company" ? '<div class="divider">BO Authorised Signatory</div><div class="cg"><label><input type="checkbox" ' + (state.boExternalAS ? "checked" : "") + ' ' + on("change", (e) => { state.boExternalAS = e.target.checked; rerender(); }) + ' /> Authorised signatory is not a board member</label></div>' + (state.boExternalAS ? '<div class="g2" style="margin-top:10px"><div class="f"><label>Name *</label><input class="'+(errs.boExternalASName?"err":"")+'" value="'+attr(state.boExternalASName)+'" '+on("input",(e)=>{state.boExternalASName=e.target.value;scheduleSave();})+' /></div><div class="f"><label>Designation *</label><input class="'+(errs.boExternalASDesignation?"err":"")+'" value="'+attr(state.boExternalASDesignation)+'" '+on("input",(e)=>{state.boExternalASDesignation=e.target.value;scheduleSave();})+' /></div></div>' : '') : '') +
        "</div></div>"
      ) : "") +

      // MDF details
      (needsMDF(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udcc4 Merchant Declaration Form (MDF)</h2><button class="reset-link" ' + on("click", () => resetSection("mdf")) + '>Reset section</button></div><div class="cbd">' +
          '<div class="g2">' +
            '<div class="f"><label>Signatory Name *</label><input class="' + (errs.mdfAuthName ? "err" : "") + '" value="' + attr(state.mdfAuthName) + '" ' + on("input", (e) => { state.mdfAuthName = e.target.value; scheduleSave(); }) + " /></div>" +
            (state.onboardingType === "salesforce" && state.entityType === "company" ? '<div class="f"><label>Signatory Designation *</label><select class="' + (errs.mdfAuthDesignation ? "err" : "") + '" ' + on("change", (e) => { state.mdfAuthDesignation = e.target.value; rerender(); }) + '>' + option("Director","Director",state.mdfAuthDesignation || "Director") + option("Managing Director","Managing Director",state.mdfAuthDesignation) + option("Company Secretary","Company Secretary",state.mdfAuthDesignation) + option("CEO","CEO",state.mdfAuthDesignation) + '</select></div>' : state.onboardingType === "salesforce" && state.entityType === "partnership" ? '<div class="f"><label>Signatory Designation</label><input value="Partner" disabled /></div>' : '<div class="f"><label>Signatory Designation *</label><input class="' + (errs.mdfAuthDesignation ? "err" : "") + '" value="' + attr(state.mdfAuthDesignation) + '" ' + on("input", (e) => { state.mdfAuthDesignation = e.target.value; scheduleSave(); }) + " /></div>") +
            '<div class="f"><label>Signatory PAN *</label><input class="' + (errs.mdfAuthPan ? "err" : "") + '" maxlength="10" value="' + attr(state.mdfAuthPan) + '" ' + on("input", (e) => { state.mdfAuthPan = e.target.value.toUpperCase(); scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Mobile Number *</label><input type="tel" maxlength="10" class="' + (errs.mdfMobile ? "err" : "") + '" value="' + attr(state.mdfMobile) + '" ' + on("input", (e) => { state.mdfMobile = e.target.value.replace(/\\D/g,"").slice(0,10); scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Email' + (state.onboardingType === "salesforce" ? " *" : " (optional)") + '</label><input type="email" class="' + (errs.mdfEmail ? "err" : "") + '" value="' + attr(state.mdfEmail) + '" ' + on("input", (e) => { state.mdfEmail = e.target.value; scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Person with Disability?</label><select ' + on("change",(e)=>{state.mdfPwd=e.target.value;rerender();}) + '>' + option("no","No",state.mdfPwd)+option("yes","Yes",state.mdfPwd)+'</select></div>' +
            (state.mdfPwd === "yes" ? '<div class="f"><label>Disability Type *</label><input value="'+attr(state.mdfPwdType)+'" '+on("input",(e)=>{state.mdfPwdType=e.target.value;scheduleSave();})+' /></div><div class="f"><label>Disability Percentage *</label><input type="number" min="1" max="100" class="'+(errs.mdfPwdPct?"err":"")+'" value="'+attr(state.mdfPwdPct)+'" '+on("input",(e)=>{state.mdfPwdPct=e.target.value;scheduleSave();})+' /></div>' : '') +
            '<div class="f"><label>TAN Status</label><select ' + on("change", (e) => { state.mdfTanStatus = e.target.value; rerender(); }) + ">" + option("no_tan","Not liable for TAN",state.mdfTanStatus) + option("has_tan","Holds TAN",state.mdfTanStatus) + "</select></div>" +
            (state.mdfTanStatus === "has_tan" ? '<div class="f"><label>TAN Number *</label><input class="'+(errs.mdfTanNum?"err":"")+'" maxlength="10" value="' + attr(state.mdfTanNum) + '" ' + on("input", (e) => { state.mdfTanNum = e.target.value.toUpperCase(); scheduleSave(); }) + " /></div>" : "") +
            (state.onboardingType === "salesforce" ? (
              '<div class="f"><label>Father\'s Name *</label><input class="'+(errs.mdfFatherName?"err":"")+'" value="'+attr(state.mdfFatherName)+'" '+on("input",(e)=>{state.mdfFatherName=e.target.value;scheduleSave();})+' /></div>' +
              '<div class="f"><label>KYC Address Proof</label><select '+on("change",(e)=>{state.mdfKycDoc=e.target.value;rerender();})+'>'+option("aadhaar","Aadhaar (masked)",state.mdfKycDoc)+option("dl","Driving License",state.mdfKycDoc)+option("voterid","Voter ID",state.mdfKycDoc)+'</select></div>' +
              '<div class="f"><label>GST Status</label><select ' + on("change", (e) => { state.mdfGstStatus = e.target.value; rerender(); }) + ">" + option("no_gst","No GST Registration",state.mdfGstStatus) + option("has_gst","Holds GST",state.mdfGstStatus) + "</select></div>" +
              (state.mdfGstStatus === "has_gst" ? '<div class="f"><label>GSTIN *</label><input class="'+(errs.mdfGstNum?"err":"")+'" maxlength="15" value="' + attr(state.mdfGstNum) + '" ' + on("input", (e) => { state.mdfGstNum = e.target.value.toUpperCase(); scheduleSave(); }) + " /></div>" : "") +
              '<div class="f"><label>Entity Nature</label><select '+on("change",(e)=>{state.mdfEntityNature=e.target.value;rerender();})+'>'+option("government","Government organization",state.mdfEntityNature)+option("ngo","NGO / Charitable institution",state.mdfEntityNature)+option("na","NA",state.mdfEntityNature)+'</select></div>'
            ) : "") +
            '<div class="f"><label>Is any covered person a PEP? *</label><select class="'+(errs.mdfPepStatus?"err":"")+'" '+on("change",(e)=>{state.mdfPepStatus=e.target.value;rerender();})+'>'+option("","Select",state.mdfPepStatus)+option("yes","Yes",state.mdfPepStatus)+option("no","No",state.mdfPepStatus)+'</select></div>' +
            '<div class="f"><label>Place *</label><input class="'+(errs.mdfPlace?"err":"")+'" value="' + attr(state.mdfPlace) + '" ' + on("input", (e) => { state.mdfPlace = e.target.value; scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Date *</label><input type="date" class="'+(errs.mdfDateRaw?"err":"")+'" value="' + attr(state.mdfDateRaw) + '" ' + on("change", (e) => { state.mdfDateRaw = e.target.value; scheduleSave(); }) + " /></div>" +
          "</div>" +
        "</div></div>"
      ) : "") +

      pageActions('<button class="btn btn-s" ' + on("click", () => { state.step = 2; rerender(); }) + '>\u2190 Back</button>', '<button class="btn btn-p" ' + on("click", () => { validateDataEntry(); if (state.formValidated) { state.step = 4; rerender(); } }) + '>Review & Generate \u2192</button>')
    );
  }

  // ---- STEP 4: Review & Generate ----
  if (state.step === 4) {
    const docLabel = (docOptionsMap[state.onboardingType][state.entityType].find(o => o.id === state.docRequirement) || {}).label || "";
    return '<div class="card"><div class="chd"><h2>\u2705 Review</h2><span class="badge">Step 5 of 5</span></div><div class="cbd">' +
      '<table class="rtbl">' +
        '<tr class="hd"><td colspan="2">MERCHANT</td></tr>' +
        '<tr><td>Agent Mobile</td><td>' + esc(state.agentMobile) + "</td></tr>" +
        '<tr><td>' + (state.merchantStatus === "new" ? "Merchant Mobile" : "Merchant ID") + '</td><td>' + esc(state.merchantStatus === "new" ? state.merchantMobile : state.merchantId) + "</td></tr>" +
        '<tr class="hd"><td colspan="2">ENTITY</td></tr>' +
        '<tr><td>Documents</td><td><strong>' + esc(docLabel) + "</strong></td></tr>" +
        '<tr><td>Legal Name</td><td>' + esc(state.firmName) + "</td></tr>" +
        '<tr><td>Registered Address</td><td>' + esc(state.regAddress) + "</td></tr>" +
        (needsResolution(state.docRequirement) ? '<tr><td>Authorised Signatory</td><td>' + esc(formData.authSignatoryName) + "</td></tr>" : "") +
      "</table></div></div>" +
      '<div class="gen-box">' +
        '<h3>\ud83c\udf89 Ready to Generate</h3><p>Preview, download Word, or download a clean A4 PDF. Print remains available inside preview.</p>' +
        (needsResolution(state.docRequirement) ? '<div class="doc-group"><h4>' + (state.entityType === "partnership" ? "Partner Resolution" : "Board Resolution") + '</h4><div class="doc-btns"><button class="btn btn-s" ' + on("click", () => previewDocx("resolution")) + '>Preview</button><button class="btn btn-p" ' + on("click", () => generateDocx("resolution")) + '>Download .docx</button><button class="btn btn-g" ' + on("click", () => downloadPdf("resolution")) + '>Download PDF</button></div></div>' : '') +
        (needsBO(state.docRequirement) ? '<div class="doc-group"><h4>BO Declaration</h4><div class="doc-btns"><button class="btn btn-s" ' + on("click", () => previewDocx("bo")) + '>Preview</button><button class="btn btn-p" ' + on("click", () => generateDocx("bo")) + '>Download .docx</button><button class="btn btn-g" ' + on("click", () => downloadPdf("bo")) + '>Download PDF</button></div></div>' : '') +
        (needsMDF(state.docRequirement) ? '<div class="doc-group"><h4>' + (state.onboardingType === "ace" ? "ACE OSV / MDF" : "MDF") + '</h4><div class="doc-btns"><button class="btn btn-s" ' + on("click", () => previewDocx("mdf")) + '>Preview</button><button class="btn btn-p" ' + on("click", () => generateDocx("mdf")) + '>Download .docx</button><button class="btn btn-g" ' + on("click", () => downloadPdf("mdf")) + '>Download PDF</button></div></div>' : '') +
      "</div>" +
      '<div class="act"><button class="btn btn-s" ' + on("click", () => { state.step = 3; rerender(); }) + '>\u2190 Back to Edit</button><div class="reset-actions"><button class="btn btn-reset" ' + on("click", clearMerchantData) + '>Start new merchant</button><button class="btn btn-danger" ' + on("click", globalReset) + '>Global reset</button></div></div>';
  }

  return "";
}

// ---- INIT ----
loadAgent();
loadMerchant();
rerender();
document.getElementById("previewClose").addEventListener("click", closePreview);
document.getElementById("previewDownload").addEventListener("click", () => {
  if (!previewArtifact) return;
  logToSheet();
  downloadBlob(previewArtifact.blob, previewArtifact.filename);
  showToast("Download started", "ok");
});
document.getElementById("previewPrint").addEventListener("click", () => {
  if (!previewArtifact) return;
  logToSheet();
  window.print();
});
document.getElementById("previewPdf").addEventListener("click", () => downloadPdf());
document.getElementById("previewModal").addEventListener("click", (event) => {
  if (event.target.id === "previewModal") closePreview();
});
