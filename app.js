// ============================================================
// Merchant Onboarding Document Generator
// Single-file client app. No backend required.
// docx generation runs entirely in the browser via docx.js (CDN).
// ============================================================
import * as docxLib from "https://cdn.jsdelivr.net/npm/docx@8.2.2/build/index.js";

// ---- CONFIG ----
// Paste your Google Apps Script Web App URL here (see google-apps-script.gs).
// Logs ONLY: timestamp, agent mobile number, merchant display name / ID.
const SHEETS_URL = "https://script.google.com/a/macros/phonepe.com/s/AKfycbyOU_MKNPCXOPT7CYAAf2tmqZsQNSVTbP92QY-gw991gyJrEQdA7Tu_c_FlCsyxIn9_/exec"; // e.g. "https://script.google.com/macros/s/XXXXXXXX/exec"

const LS_MERCHANT = "docgen_merchant_v1";   // clears per onboarding
const LS_AGENT    = "docgen_agent_v1";      // permanent - agent mobile only

const PAN_REGEX = /^[A-Z]{5}\d{4}[A-Z]$/;
const AADHAAR_LAST4 = /^\d{4}$/;
const MOBILE_REGEX = /^[6-9]\d{9}$/;

// ---- STATE ----
const INITIAL_PARTNER = (id, designation) => ({
  id, name: "", pan: "", designation: designation || "Partner", dobRaw: "",
  address: "", poa: "AADHAAR", poaNum: "", nationality: "Indian", share: "", isAS: false,
});

const state = {
  // Step 0 - Agent + Merchant identification (the "only 2 datapoints" logged to Sheet)
  agentMobile: "",
  merchantStatus: "new",       // 'new' | 'existing'
  merchantDisplayName: "",     // required if new
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
  includeLetterhead: false,
  lhTagline: "", lhPhone: "", lhEmail: "", lhWebsite: "",
  principalSame: "same",
  principalAddress: "",
  partnershipRegType: "registered",
  naturalPersons: "yes",
  deedDate: "",

  // Partners / Directors
  partners: [INITIAL_PARTNER(1), INITIAL_PARTNER(2)],
  nextPartnerId: 3,
  presentPartnerIds: [],

  // Resolution
  resolutionDateRaw: "",
  resolutionTimeRaw: "",
  pepDeclarationRes: true,

  // BO
  boDate: "",
  boCategory: "cat1",
  pepDeclarationBO: true,

  // MDF
  mdfAuthName: "", mdfAuthDesignation: "", mdfAuthPan: "",
  mdfMobile: "", mdfEmail: "",
  mdfPwd: "no", mdfPwdType: "", mdfPwdPct: "",
  mdfTanStatus: "no_tan", mdfTanNum: "",
  mdfGstStatus: "no_gst", mdfGstNum: "",
  mdfPepConfirm: true, mdfDateRaw: "", mdfPlace: "",

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
      { id: "set3", label: "Resolution + BO Declaration" },
    ],
    company: [
      { id: "set5", label: "Board Resolution" },
      { id: "set6", label: "BO Declaration" },
      { id: "set7", label: "Resolution + BO Declaration" },
    ],
  },
  salesforce: {
    partnership: [
      { id: "set9", label: "Resolution" },
      { id: "set10", label: "BO Declaration" },
      { id: "set11", label: "MDF" },
      { id: "set12", label: "Resolution + BO Declaration + MDF" },
    ],
    company: [
      { id: "set13", label: "Board Resolution" },
      { id: "set14", label: "BO Declaration" },
      { id: "set15", label: "MDF" },
      { id: "set16", label: "Board Resolution + BO Declaration + MDF" },
    ],
  },
};
const needsResolution = (set) => ["set1","set3","set5","set7","set9","set12","set13","set16"].includes(set);
const needsBO         = (set) => ["set2","set3","set6","set7","set10","set12","set14","set16"].includes(set);
const needsMDF        = (set) => ["set11","set12","set15","set16"].includes(set);
const needsFullKYC    = (set) => needsBO(set) || needsResolution(set);

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
      merchant: state.merchantStatus === "new" ? state.merchantDisplayName : state.merchantId,
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
  if (state.partners.length <= 2) { showToast("Minimum 2 required", "er"); return; }
  state.partners = state.partners.filter((p) => p.id !== id);
  state.presentPartnerIds = state.presentPartnerIds.filter((pid) => pid !== id);
  rerender();
}
function onSelectAS(id) {
  state.partners = state.partners.map((p) => ({ ...p, isAS: p.id === id }));
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
  state.resolutionDateRaw = d; state.resolutionTimeRaw = t; state.boDate = d; state.mdfDateRaw = d;
  rerender();
}

// ---- VALIDATION ----
function validateStep0() {
  const errs = {};
  if (!MOBILE_REGEX.test(state.agentMobile)) errs.agentMobile = "Enter a valid 10-digit mobile number";
  if (state.merchantStatus === "new" && !state.merchantDisplayName.trim()) errs.merchantDisplayName = "Establishment / Merchant Display Name is required";
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

  if (needsResolution(state.docRequirement) || needsBO(state.docRequirement)) {
    if (state.partners.length < 2) addError("Minimum 2 individuals are required.", "partners");
    if (!state.partners.some((p) => p.isAS)) addError('Tick "Authorised Signatory" for one member.', "authSignatory");

    let totalShare = 0;
    const requiresFull = needsFullKYC(state.docRequirement);
    state.partners.forEach((partner, index) => {
      const label = "Member " + (index + 1) + ": ";
      if (!partner.name.trim()) addError(label + "Name is required.", "name_" + partner.id);
      if (requiresFull) {
        if (!partner.pan || !PAN_REGEX.test(partner.pan.toUpperCase())) addError(label + "Valid PAN required.", "pan_" + partner.id);
        if (!partner.dobRaw) addError(label + "DOB required.", "dobRaw_" + partner.id);
        if (!partner.address.trim()) addError(label + "Address required.", "address_" + partner.id);
        if (!partner.poaNum) addError(label + "Proof number required.", "poaNum_" + partner.id);
        else if (partner.poa === "AADHAAR" && !AADHAAR_LAST4.test(partner.poaNum)) addError(label + "Aadhaar proof must be exactly last 4 digits.", "poaNum_" + partner.id);
        if (partner.share === "") addError(label + "Share % required.", "share_" + partner.id);
        else totalShare += Number(partner.share || 0);
      }
    });
    if (requiresFull && Math.abs(totalShare - 100) > 0.6) {
      addError("Total shareholding must equal exactly 100% (currently " + totalShare.toFixed(1) + "%).", "shareTotal");
    }
  }

  if (needsResolution(state.docRequirement)) {
    if (!state.resolutionDateRaw) addError("Resolution date required.", "resolutionDateRaw");
    if (!state.resolutionTimeRaw) addError("Resolution time required.", "resolutionTimeRaw");
    if (state.presentPartnerIds.length === 0) addError("Select at least one present signatory.", "presentPartnerIds");
    if (state.entityType === "partnership") {
      const presentShare = state.partners.filter(p => state.presentPartnerIds.includes(p.id)).reduce((s,p)=>s+Number(p.share||0),0);
      if (presentShare <= 50) addError("Partners present must hold >50% combined share (currently " + presentShare.toFixed(1) + "%).", "presentShare");
    }
  }

  if (needsBO(state.docRequirement)) {
    if (!state.boDate) addError("BO declaration date required.", "boDate");
    if (needsResolution(state.docRequirement) && state.resolutionDateRaw && state.boDate && formatDate(state.resolutionDateRaw) !== formatDate(state.boDate)) {
      addError("BO date must match Resolution date.", "boDate");
    }
  }

  if (needsMDF(state.docRequirement)) {
    if (!state.mdfAuthName.trim()) addError("Signatory Name is required.", "mdfAuthName");
    if (!state.mdfAuthPan || !PAN_REGEX.test(state.mdfAuthPan.toUpperCase())) addError("Valid PAN is required.", "mdfAuthPan");
    if (!MOBILE_REGEX.test(state.mdfMobile)) addError("Valid mobile number required.", "mdfMobile");
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
      new TableRow({ children: list.map(() => new TableCell({ borders: noBdrs, width: { size: colW, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [p([t("____________Sign____________")])] })) }),
      new TableRow({ children: list.map((m) => new TableCell({ borders: noBdrs, width: { size: colW, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [p([u(m.name)]), p([t("(" ), u(m.designation||""), t(")")])] })) }),
    ],
  });
}
function buildLetterheadHeader(lh) {
  const lt = (text, opts={}) => new TextRun({ text: text||"", font: "Times New Roman", size: 18, ...opts });
  const logoCell = new TableCell({ borders: noBdrs, width: { size: 3600, type: WidthType.DXA }, margins: { top:0,bottom:0,left:0,right:200 },
    children: [
      new Paragraph({ children: [new TextRun({ text: lh.firmName||"", font:"Times New Roman", size:36, bold:true, color:"5f259f" })], spacing:{after:40} }),
      ...(lh.lhTagline ? [new Paragraph({ children:[lt(lh.lhTagline,{italics:true,color:"888888"})], spacing:{after:0} })] : []),
    ]});
  const contactLines = [
    lh.regAddress ? "Regd Office: " + lh.regAddress : null,
    lh.lhPhone ? "Phone: " + lh.lhPhone : null,
    lh.lhEmail ? "Email: " + lh.lhEmail : null,
    lh.lhWebsite ? "Website: " + lh.lhWebsite : null,
  ].filter(Boolean);
  const contactCell = new TableCell({ borders: noBdrs, width: { size: 6480, type: WidthType.DXA }, margins:{top:0,bottom:0,left:200,right:0},
    children: contactLines.map((line) => new Paragraph({ children:[lt(line)], alignment: AlignmentType.RIGHT, spacing:{after:20} })) });
  const lhTable = new Table({ width:{size:10080,type:WidthType.DXA}, columnWidths:[3600,6480], borders: noBdrs, rows:[new TableRow({children:[logoCell,contactCell]})] });
  const hrPara = new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "5f259f", space: 4 } }, spacing: { after: 200 } });
  return new Header({ children: [lhTable, hrPara] });
}
function pageSetup(hasLH) {
  return { size: { width: 12240, height: 15840 }, margin: { top: hasLH ? 1800 : 1440, right: 1440, bottom: 1440, left: 1440 } };
}
async function downloadDoc(doc, filename) {
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.style.display = "none"; a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); document.body.removeChild(a); }, 2000);
}

// ---- PARTNERSHIP: RESOLUTION ----
async function buildPartnershipResolution(d) {
  const hasLH = d.includeLetterhead && d.firmName;
  const principalText = d.principalSame === "same" ? "same \u2611 (tick if applicable)" : d.principalAddress;
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
        t(" AT ", { bold: true }), u(d.resolutionTime, { bold: true }),
        t(" having its registered office address at "), u(d.regAddress),
        t(" and having its principal place of operation/office at "),
        ...(d.principalSame === "same" ? [t(principalText)] : [u(d.principalAddress)]), t("."),
      ], spacing: { after: 200 } }),
      p([]),
      p([t("PRESENT:", { bold: true })]),
      ...d.partners.map((m, i) => p([t((i + 1) + ". "), u(m.name)])),
      p([t("(List of partners present during the resolution)", { italics: true })]),
      p([]),
      new Paragraph({ children: [
        t("RESOLVED THAT ", { bold: true }), t("Mr/Mrs "), u(d.authSignatoryName), t(" ["), u(d.authSignatoryDesignation), t("], be and is hereby authorized, to act on behalf of the Firm and to execute/sign all necessary applications/documents for the purpose of opening and operating a business account with PhonePe Limited."),
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

      new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),
      p([]),
      new Paragraph({ children: [new TextRun({ text: "Declaration", font: "Times New Roman", size: 24, bold: true, underline: { type: UnderlineType.SINGLE } })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
      new Paragraph({ children: [t("I/we, the undersigned individuals, hereby personally, jointly, and severally undertake and declare that:")], spacing: { after: 160 } }),
      new Paragraph({ children: [t("1. Our firm ["), u(d.firmName), t("] is constituted as a partnership firm and it is")], spacing: { after: 100 } }),
      p([t(d.isRegistered ? "\u2611 Registered" : "\u2610 Registered")], { indent: { left: 720 } }),
      new Paragraph({ children: [t("Whether there is/are any natural person(s), acting alone or through other juridical persons, with more than 10% ownership/entitlement to capital/profits, or who exercise control through other means.")], indent: { left: 1440 }, spacing: { after: 80 } }),
      ...(d.isRegistered ? [
        p([t(d.naturalPersons === "yes" ? "\u2611 Yes" : "\u2610 Yes")], { indent: { left: 1440 } }),
        p([t(d.naturalPersons === "no" ? "\u2611 No" : "\u2610 No")], { indent: { left: 1440 } }),
      ] : [p([t("\u2610 Yes   \u2610 No")], { indent: { left: 1440 } })]),
      p([t(d.isRegistered ? "\u2610 Unregistered" : "\u2611 Unregistered")], { indent: { left: 720 } }),
      new Paragraph({ children: [t("Whether there is/are any natural person(s), acting alone or through other juridical persons, with more than 15% ownership/entitlement to capital/profits, or who exercise control through other means.")], indent: { left: 1440 }, spacing: { after: 80 } }),
      ...(!d.isRegistered ? [
        p([t(d.naturalPersons === "yes" ? "\u2611 Yes" : "\u2610 Yes")], { indent: { left: 1440 } }),
        p([t(d.naturalPersons === "no" ? "\u2611 No" : "\u2610 No")], { indent: { left: 1440 } }),
      ] : [p([t("\u2610 Yes   \u2610 No")], { indent: { left: 1440 } })]),
      p([]),
      new Paragraph({ children: [t("2. If there is/are no natural person(s) as per the responses provided under declaration (1) hereinabove, then we declare that the authorised signatory identified in the Resolution is the senior managing official and be considered as the beneficial owner of our firm.")], spacing: { after: 120 } }),
      new Paragraph({ children: [t("3. " + (d.pepDeclarationRes ? "\u2611" : "\u2610") + " No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a "), t("Politically Exposed Person (PEP) (as defined by RBI).", { bold: true })], spacing: { after: 120 } }),
      new Paragraph({ children: [t("4. The contents of the resolution of the partners dated ["), u(d.resolutionDate), t("] (\u201cResolution\u201d) are true and valid.")], spacing: { after: 120 } }),
      new Paragraph({ children: [t("5. The list of partners constituting the partnership firm and their respective details, as provided in the partnership deed dated ["), u(d.deedDateFmt), t("] (\u201cPartnership Deed\u201d), are true, complete, current, and valid.")], spacing: { after: 120 } }),
      new Paragraph({ children: [t("6. The percentage of ownership and/or entitlement to capital or profits of the partners as specified in the Partnership Deed is true, complete, current, and valid.")], spacing: { after: 200 } }),
      new Paragraph({ children: [t("In consideration of PhonePe agreeing to rely on the Resolution, the Partnership Deed, and this declaration, I/we hereby personally, jointly, and severally undertake to indemnify and hold PhonePe harmless against all damages, liabilities, claims, demands, actions, proceedings, losses, costs (including legal costs), expenses, and all other liabilities of whatsoever nature or description, arising out of or in connection with PhonePe\u2019s reliance on the said Resolution, Partnership Deed, and this declaration.")], spacing: { after: 240 } }),
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
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(hasLH) },
    headers: hasLH ? { default: buildLetterheadHeader(d) } : undefined,
    children: [
      ...(!hasLH ? [p([t("(To be printed on Partnership Firm\u2019s Letterhead)", { italics: true })], { alignment: AlignmentType.CENTER })] : []),
      p([]),
      new Paragraph({ children: [new TextRun({ text: "DECLARATION OF BENEFICIAL OWNERSHIP (BO) and LIST OF PARTNERS", font: "Times New Roman", size: 26, bold: true, underline: { type: UnderlineType.SINGLE } })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: [500, 2500, 7080], rows: [
        new TableRow({ children: [cell("I", 500), cell("Name of the entity", 2500), ucell(d.firmName, 7080)] }),
        new TableRow({ children: [cell("II", 500), cell("Registered address", 2500), ucell(d.regAddress, 7080)] }),
        new TableRow({ children: [cell("III", 500), cell("Principal place of operation/office", 2500),
          new TableCell({ borders: bdrs, width: { size: 7080, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: d.principalSame === "same" ? [p([t("\u2611 Same")])] : [p([t("\u2610 Same   OR "), u(d.principalAddress)])] }) ] }),
        new TableRow({ children: [cell("IV", 500), cell("Type of entity", 2500),
          new TableCell({ borders: bdrs, width: { size: 7080, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [p([t(d.isRegistered ? "\u2611" : "\u2610"), t(" Partnership Firm / LLP")]), p([t(d.isRegistered ? "\u2610" : "\u2611"), t(" Unregistered Partnership Firm")])] }) ] }),
      ] }),
      p([]),
      new Paragraph({ children: [t("The Legal Entity as stated above hereby confirms and declares the following on the below date: ", { bold: true }), u(d.boDateFmt, { bold: true })], spacing: { after: 100 } }),
      new Paragraph({ children: [t((d.boCategory === "cat1" ? "\u2611" : "\u2610") + " Category 1", { bold: true }), t(" - Persons/entities noted below own " + threshold + "% or more interest or possess the right to control management/policy decisions.")], spacing: { after: 80 } }),
      new Paragraph({ children: [t((d.boCategory === "cat2" ? "\u2611" : "\u2610") + " Category 2", { bold: true }), t(" - No natural person is identified as per Category 1; senior managing official is the BO.")], spacing: { after: 120 } }),
      new Paragraph({ children: [t((d.pepDeclarationBO ? "\u2611" : "\u2610") + " No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a "), t("Politically Exposed Person (PEP) (as defined by RBI).", { bold: true })], spacing: { after: 160 } }),
      p([t("The details of beneficial owner(s) is/are as follows:")]),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: boCols, rows: [
        new TableRow({ tableHeader: true, children: [
          hCell("S.N.", boCols[0]), hCell("Name", boCols[1]), hCell("Residential Address and PIN code", boCols[2]),
          hCell("Designation", boCols[3]), hCell("DOB", boCols[4]), hCell("Proof of identity", boCols[5]),
          hCell("Proof of Address", boCols[6]), hCell("Nationality", boCols[7]), hCell("% of interest", boCols[8]),
        ]}),
        ...d.partners.map((m, i) => new TableRow({ children: [
          cell(String(i + 1), boCols[0]), ucell(m.name, boCols[1]), ucell(m.address, boCols[2]),
          ucell(m.designation, boCols[3]), ucell(m.dob, boCols[4]), ucell(m.pan, boCols[5]),
          ucell(m.poa + " " + m.poaNumDisplay, boCols[6]), ucell(m.nationality, boCols[7]), ucell(m.share + "%", boCols[8]),
        ]})),
      ]}),
      p([]),
      p([t("List of the Current Partners of the firm operating at the aforementioned address:")]),
      ...d.partners.map((m, i) => p([t((i + 1) + ". "), u(m.name)])),
      p([]),
      p([t("We acknowledge and confirm that the information provided above is true and correct to the best of our knowledge and belief.")]),
      p([]),
      p([t("Authorised Signatory/ies:", { bold: true })]),
      p([t("___________________________(Name, Signature with Stamp)")]),
      p([u(d.authSignatoryName)]),
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
        new TableRow({ children: [hCell("Name of authorised signatory", 3500), hCell("Designation", 2900), hCell("Signature", 2960)] }),
        ...present.map((m) => new TableRow({ children: [ucell(m.name, 3500), ucell(m.designation, 2900), cell("[Sign Here]", 2960)] })),
      ]}),
      p([]),
      p([t("Seal of the Company")], { alignment: AlignmentType.CENTER }),
    ],
  }] });
  return doc;
}

// ---- COMPANY: BO DECLARATION ----
async function buildCompanyBO(d) {
  const hasLH = d.includeLetterhead && d.firmName;
  const validRows = d.partners.filter((m) => Number(m.share) >= 10);
  const source = validRows.length ? validRows : d.partners;
  const boCols = [400, 1100, 1800, 680, 680, 980, 980, 680, 780];
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(hasLH) },
    headers: hasLH ? { default: buildLetterheadHeader(d) } : undefined,
    children: [
      new Paragraph({ children: [new TextRun({ text: "DECLARATION OF BENEFICIAL OWNERSHIP (BO)", font: "Times New Roman", size: 26, bold: true, underline: { type: UnderlineType.SINGLE } })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: [500, 2500, 7080], rows: [
        new TableRow({ children: [cell("I", 500), cell("Name of the entity", 2500), ucell(d.firmName, 7080)] }),
        new TableRow({ children: [cell("II", 500), cell("Registered address", 2500), ucell(d.regAddress, 7080)] }),
        new TableRow({ children: [cell("III", 500), cell("Principal place of operation/office", 2500),
          new TableCell({ borders: bdrs, width: { size: 7080, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: d.principalSame === "same" ? [p([t("\u2611 Same")])] : [p([t("\u2610 Same   OR "), u(d.principalAddress)])] }) ] }),
        new TableRow({ children: [cell("IV", 500), cell("Type of entity", 2500), ucell("Company", 7080)] }),
      ]}),
      p([]),
      new Paragraph({ children: [t("The Legal Entity as stated above hereby confirms and declares the following on the date: ", { bold: true }), u(d.boDateFmt, { bold: true })], spacing: { after: 100 } }),
      new Paragraph({ children: [t((d.boCategory === "cat1" ? "\u2611" : "\u2610") + " Category 1", { bold: true }), t(" - Following persons/entities own 10% or more interest or possess the right to control management/policy decisions.")], spacing: { after: 80 } }),
      new Paragraph({ children: [t((d.boCategory === "cat2" ? "\u2611" : "\u2610") + " Category 2", { bold: true }), t(" - No natural person is identified as per Category 1 above.")], spacing: { after: 160 } }),
      p([t("The details of beneficial owner(s) is/are as follows:")]),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: boCols, rows: [
        new TableRow({ tableHeader: true, children: [
          hCell("S.N.", boCols[0]), hCell("Name", boCols[1]), hCell("Residential Address and PIN code", boCols[2]),
          hCell("Designation", boCols[3]), hCell("DOB", boCols[4]), hCell("Proof of identity", boCols[5]),
          hCell("Proof of Address", boCols[6]), hCell("Nationality", boCols[7]), hCell("% of interest", boCols[8]),
        ]}),
        ...source.map((m, i) => new TableRow({ children: [
          cell(String(i + 1), boCols[0]), ucell(m.name, boCols[1]), ucell(m.address, boCols[2]),
          ucell(m.designation, boCols[3]), ucell(m.dob, boCols[4]), ucell(m.pan, boCols[5]),
          ucell(m.poa + " " + m.poaNumDisplay, boCols[6]), ucell(m.nationality, boCols[7]), ucell(m.share + "%", boCols[8]),
        ]})),
      ]}),
      p([]),
      p([t("List of Senior Management Officials:")]),
      new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: [5040, 5040], rows: [
        new TableRow({ children: [hCell("Name", 5040), hCell("Designation", 5040)] }),
        ...d.partners.map((m) => new TableRow({ children: [ucell(m.name, 5040), ucell(m.designation, 5040)] })),
      ]}),
      p([]),
      new Paragraph({ children: [t((d.pepDeclarationBO ? "\u2611" : "\u2610") + " No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a "), t("Politically Exposed Person (PEP) (as defined by RBI).", { bold: true })], spacing: { after: 160 } }),
      p([t("Authorised Signatory/ies:", { bold: true })]),
      p([t("___________________________(Name, Signature with Stamp)")]),
      p([u(d.authSignatoryName)]),
    ],
  }] });
  return doc;
}

// ---- MDF (Merchant Declaration Form) ----
async function buildMDF(d) {
  const isACE = d.onboardingType === "ace";
  const principalText = d.principalSame === "same" ? "same (\u2611)" : d.principalAddress;
  const isCompany = d.entityType === "company";
  const isRegPartnership = d.entityType === "partnership" && d.isRegistered;
  const isUnregPartnership = d.entityType === "partnership" && !d.isRegistered;
  const doc = new Document({ sections: [{
    properties: { page: pageSetup(false) },
    children: [
      p([t("Merchant Declaration Form (" + (isACE ? "ACE" : "Salesforce") + ")", { bold: true })], { alignment: AlignmentType.RIGHT }),
      p([t("To,")]),
      p([t("PhonePe Limited (Formerly known as 'PhonePe Private Limited')")]),
      p([t("Office-2, Floor 5, Wing A, Block A, Salarpuria Softzone, Bellandur Village, Varthur Hobli, Outer Ring Road, Bangalore South, Bangalore, Karnataka, India, 560103")]),
      p([]),
      p([t("Subject: PhonePe Merchant Declaration", { bold: true })]),
      p([]),
      new Paragraph({ children: [
        t("I, "), u(d.mdfAuthName), t(", hereinafter referred to as \u201cMerchant\u201d being the "), u(d.mdfAuthDesignation),
        t(" of "), u(d.firmName), t(" having its registered office address at "), u(d.regAddress),
        t(" (\u201cEntity\u201d) and having its principal place of operation/office at "), u(principalText),
        t(", do hereby declare that I have been authorised to act as a designated authorised signatory for the Entity."),
      ], spacing: { after: 200 } }),
      p([t("Details provided under this declaration:", { bold: true })]),
      p([t("1. Mobile No. (registered with PhonePe for Onboarding): "), u(d.mdfMobile)]),
      p([t("2. Email ID (registered with PhonePe for Onboarding): "), u(d.mdfEmail)]),
      ...( !isACE ? [
        p([t("3. In case of Person with Disability (PwD):")]),
        p([t("   Type of Disability: "), u(d.mdfPwd === "yes" ? d.mdfPwdType : "N/A")]),
        p([t("   Percentage of Disability: "), u(d.mdfPwd === "yes" ? d.mdfPwdPct : "N/A")]),
      ] : []),
      p([]),
      p([t("I, on behalf of the Merchant, further declare that:", { bold: true })]),
      ...( !isACE ? [
        p([t((isCompany ? "\u2611" : "\u2610") + " Company    " + (isRegPartnership ? "\u2611" : "\u2610") + " Registered Partnership Firm    " + (isUnregPartnership ? "\u2611" : "\u2610") + " Un-Registered Partnership Firm")]),
      ] : []),
      new Paragraph({ children: [
        t("1. The Merchant is registered under the Income Tax Act, 1961 and has obtained TAN Number "),
        u(d.mdfTanStatus === "has_tan" ? d.mdfTanNum : "N/A"),
        t(", OR does not hold TAN as it is not liable to deduct tax at source (" + (d.mdfTanStatus === "no_tan" ? "\u2611" : "\u2610") + ")."),
      ], spacing: { after: 120 } }),
      ...( !isACE ? [ new Paragraph({ children: [
        t("2. The Merchant is registered and has a GSTIN "), u(d.mdfGstStatus === "has_gst" ? d.mdfGstNum : "N/A"),
        t(", OR does not have any registration with GST authorities (" + (d.mdfGstStatus === "no_gst" ? "\u2611" : "\u2610") + ")."),
      ], spacing: { after: 120 } }) ] : []),
      new Paragraph({ children: [
        t("No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a Politically Exposed Person (PEP): "),
        t((d.mdfPepConfirm ? "NO \u2611" : "YES \u2611")),
      ], spacing: { after: 160 } }),
      new Paragraph({ children: [t("I, having PAN number "), u(d.mdfAuthPan), t(", hereby declare that the above facts and information are true, complete and correct to the best of my knowledge.")], spacing: { after: 300 } }),
      p([t("Yours faithfully, For and behalf of the Merchant")]),
      p([]), p([]),
      p([t("___________________ (Signature with Seal)")]),
      p([t("Designation: "), u(d.mdfAuthDesignation)]),
      p([t("Date: "), u(d.mdfDate), t("   Place: "), u(d.mdfPlace)]),
    ],
  }] });
  return doc;
}

// ---- PRINT / SAVE-AS-PDF FALLBACK (uses browser print dialog) ----
function printHTML(bodyHtml, title) {
  const w = window.open("", "_blank");
  if (!w) { showToast("Please allow pop-ups to print/save as PDF", "er"); return; }
  w.document.open();
  w.document.write(
    "<!doctype html><html><head><title>" + esc(title) + "</title><style>" +
    "@page{size:A4;margin:16mm;}body{font-family:Georgia,'Times New Roman',serif;color:#111;font-size:13px;line-height:1.55;}" +
    "table{width:100%;border-collapse:collapse;margin:10px 0;}td,th{border:1px solid #000;padding:6px;font-size:12px;}" +
    "u strong{text-decoration:underline;}h2{text-align:center;}" +
    "</style></head><body>" + bodyHtml + "</body></html>"
  );
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 300);
}
const cbx = (checked) => (checked ? "\u2611" : "\u2610");
function printablePartnershipResolution(d) {
  const present = d.presentPartners.length ? d.presentPartners : d.partners;
  const list = (arr) => arr.map((m,i)=>"<div>"+(i+1)+". <u><strong>"+esc(m.name)+"</strong></u></div>").join("");
  return "<h2>TO WHOMSOEVER IT MAY CONCERN</h2><p>RESOLUTION OF THE PARTNERS OF <u><strong>"+esc((d.firmName||"").toUpperCase())+"</strong></u> HELD ON <u><strong>"+esc(d.resolutionDate)+"</strong></u> AT <u><strong>"+esc(d.resolutionTime)+"</strong></u> at <u><strong>"+esc(d.regAddress)+"</strong></u>.</p><p><strong>PRESENT:</strong></p>"+list(present)+"<p>RESOLVED THAT Mr/Mrs <u><strong>"+esc(d.authSignatoryName)+"</strong></u> be authorised to sign all documents for opening a PhonePe business account.</p><table><tr>"+present.map(()=>"<td>Sign</td>").join("")+"</tr><tr>"+present.map(m=>"<td><u><strong>"+esc(m.name)+"</strong></u></td>").join("")+"</tr></table>";
}
function printablePartnershipBO(d) {
  const rows = d.partners.map((m,i)=>"<tr><td>"+(i+1)+"</td><td><u><strong>"+esc(m.name)+"</strong></u></td><td>"+esc(m.address)+"</td><td>"+esc(m.designation)+"</td><td>"+esc(m.dob)+"</td><td>"+esc(m.pan)+"</td><td>"+esc(m.poa)+" "+esc(m.poaNumDisplay)+"</td><td>"+esc(m.nationality)+"</td><td>"+esc(m.share)+"%</td></tr>").join("");
  return "<h2>DECLARATION OF BENEFICIAL OWNERSHIP</h2><table><tr><td>Entity</td><td><u><strong>"+esc(d.firmName)+"</strong></u></td></tr><tr><td>Registered Address</td><td>"+esc(d.regAddress)+"</td></tr></table><table><tr><th>S.N.</th><th>Name</th><th>Address</th><th>Designation</th><th>DOB</th><th>PAN</th><th>POA</th><th>Nationality</th><th>%</th></tr>"+rows+"</table><p>Authorised Signatory: <u><strong>"+esc(d.authSignatoryName)+"</strong></u></p>";
}
function printableCompanyResolution(d) {
  const present = d.presentPartners.length ? d.presentPartners : d.partners;
  const list = (arr) => arr.map((m,i)=>"<div>"+(i+1)+". <u><strong>"+esc(m.name)+"</strong></u> ("+esc(m.designation)+")</div>").join("");
  return "<h2>CERTIFIED TRUE COPY - BOARD RESOLUTION</h2><p>Meeting of the Board of Directors of <u><strong>"+esc((d.firmName||"").toUpperCase())+"</strong></u> held on <u><strong>"+esc(d.resolutionDate)+"</strong></u> at <u><strong>"+esc(d.resolutionTime)+"</strong></u>.</p>"+list(d.partners)+"<p>RESOLVED THAT <u><strong>"+esc(d.authSignatoryName)+"</strong></u> ("+esc(d.authSignatoryDesignation)+") be authorised to sign all documents for the PhonePe business account.</p>";
}
function printableCompanyBO(d) {
  const rows = d.partners.map((m,i)=>"<tr><td>"+(i+1)+"</td><td><u><strong>"+esc(m.name)+"</strong></u></td><td>"+esc(m.address)+"</td><td>"+esc(m.designation)+"</td><td>"+esc(m.dob)+"</td><td>"+esc(m.pan)+"</td><td>"+esc(m.poa)+" "+esc(m.poaNumDisplay)+"</td><td>"+esc(m.nationality)+"</td><td>"+esc(m.share)+"%</td></tr>").join("");
  return "<h2>DECLARATION OF BENEFICIAL OWNERSHIP</h2><table><tr><td>Entity</td><td><u><strong>"+esc(d.firmName)+"</strong></u></td></tr></table><table><tr><th>S.N.</th><th>Name</th><th>Address</th><th>Designation</th><th>DOB</th><th>PAN</th><th>POA</th><th>Nationality</th><th>%</th></tr>"+rows+"</table><p>Authorised Signatory: <u><strong>"+esc(d.authSignatoryName)+"</strong></u></p>";
}
function printableMDF(d) {
  return "<h2>Merchant Declaration Form</h2><p>I, <u><strong>"+esc(d.mdfAuthName)+"</strong></u> ("+esc(d.mdfAuthDesignation)+") of <u><strong>"+esc(d.firmName)+"</strong></u>, PAN <u><strong>"+esc(d.mdfAuthPan)+"</strong></u>, Mobile <u><strong>"+esc(d.mdfMobile)+"</strong></u>, declare the above facts true and correct.</p><p>Date: <u><strong>"+esc(d.mdfDate)+"</strong></u> Place: <u><strong>"+esc(d.mdfPlace)+"</strong></u></p>";
}

// ---- GENERATE HANDLERS ----
async function generateDocx(kind) {
  computeDerived();
  logToSheet();
  const fn = (formData.firmName || "document").replace(/[^a-zA-Z0-9]/g, "_");
  try {
    let doc, suffix;
    if (kind === "resolution") { doc = state.entityType === "company" ? await buildCompanyResolution(formData) : await buildPartnershipResolution(formData); suffix = "_Resolution.docx"; }
    else if (kind === "bo") { doc = state.entityType === "company" ? await buildCompanyBO(formData) : await buildPartnershipBO(formData); suffix = "_BO_Declaration.docx"; }
    else if (kind === "mdf") { doc = await buildMDF(formData); suffix = "_MDF.docx"; }
    await downloadDoc(doc, fn + suffix);
    showToast("Download started", "ok");
  } catch (e) { console.error(e); showToast("Error: " + e.message, "er"); }
}
function generatePrint(kind) {
  computeDerived();
  logToSheet();
  if (kind === "resolution") printHTML(state.entityType === "company" ? printableCompanyResolution(formData) : printablePartnershipResolution(formData), "Resolution");
  else if (kind === "bo") printHTML(state.entityType === "company" ? printableCompanyBO(formData) : printablePartnershipBO(formData), "BO Declaration");
  else if (kind === "mdf") printHTML(printableMDF(formData), "MDF");
}

// ============================================================
// UI RENDERING
// ============================================================
function renderProgress() {
  const bar = document.getElementById("progressBar");
  if (!bar) return;
  const total = 6;
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
      '<div class="f"><label>Designation *</label><select ' + on("change", (e) => { onPartnerChange(partner.id, "designation", e.target.value); rerender(); }) + ">" +
        (state.entityType === "company"
          ? option("Director","Director",partner.designation) + option("Company Secretary","Company Secretary",partner.designation) + option("CEO","CEO",partner.designation) + option("Managing Director","Managing Director",partner.designation)
          : option("Partner","Partner",partner.designation) + option("Owner","Owner",partner.designation) + option("Authorized Signatory","Authorized Signatory",partner.designation)
        ) + "</select></div>" +
      (fullKYC ? (
        '<div class="f"><label>PAN *</label><input class="' + (err("pan") ? "err" : "") + '" maxlength="10" value="' + attr(partner.pan) + '" ' + on("input", (e) => onPartnerChange(partner.id, "pan", e.target.value.toUpperCase())) + " /></div>" +
        '<div class="f"><label>DOB *</label><input type="date" class="' + (err("dobRaw") ? "err" : "") + '" value="' + attr(partner.dobRaw) + '" ' + on("input", (e) => onPartnerChange(partner.id, "dobRaw", e.target.value)) + " /></div>" +
        '<div class="f s2 s3"><label>Residential Address & PIN *</label><input class="' + (err("address") ? "err" : "") + '" value="' + attr(partner.address) + '" ' + on("input", (e) => onPartnerChange(partner.id, "address", e.target.value)) + " /></div>" +
        '<div class="f"><label>Proof Type *</label><select ' + on("change", (e) => { onPartnerChange(partner.id, "poa", e.target.value); onPartnerChange(partner.id, "poaNum", ""); rerender(); }) + ">" + option("AADHAAR","Aadhaar",partner.poa) + option("VOTER ID","Voter ID",partner.poa) + option("DRIVING LICENSE","Driving License",partner.poa) + option("PASSPORT","Passport",partner.poa) + "</select></div>" +
        '<div class="f"><label>Proof Number *</label><div class="aadhar-wrap ' + (err("poaNum") ? "err" : "") + '"><input class="aadhar-input" placeholder="' + (isAadhaar ? "Last 4 digits" : "Reference number") + '" maxlength="' + (isAadhaar ? "4" : "30") + '" value="' + attr(partner.poaNum) + '" ' + on("input", (e) => { if (isAadhaar) e.target.value = e.target.value.replace(/\\D/g,""); onPartnerChange(partner.id, "poaNum", e.target.value); }) + " /></div></div>" +
        '<div class="f"><label>Nationality *</label><select ' + on("change", (e) => onPartnerChange(partner.id, "nationality", e.target.value)) + ">" + option("Indian","Indian",partner.nationality) + option("Foreign National","Foreign National",partner.nationality) + "</select></div>" +
        '<div class="f"><label>% Ownership *</label><input class="' + (err("share") ? "err" : "") + '" inputmode="decimal" value="' + attr(partner.share) + '" ' + on("input", (e) => onPartnerChange(partner.id, "share", sanitizeShare(e.target.value))) + " /></div>"
      ) : "") +
    "</div></div>";
}

function renderApp() {
  computeDerived();
  const errs = state.inlineErrors;

  // ---- STEP 0: Agent + Merchant identification ----
  if (state.step === 0) {
    return '<div class="card"><div class="chd"><h2>\ud83d\udcf1 Agent & Merchant Details</h2><span class="badge">Step 1 of 6</span></div><div class="cbd">' +
      '<div class="info-blue">Your mobile number is saved permanently on this device \u2014 you will not need to re-enter it next time. Merchant details reset for every new onboarding.</div>' +
      '<div class="g2">' +
        '<div class="f s2"><label>Your Mobile Number (Agent) *</label><input type="tel" inputmode="numeric" maxlength="10" class="' + (errs.agentMobile ? "err" : "") + '" value="' + attr(state.agentMobile) + '" ' + on("input", (e) => { state.agentMobile = e.target.value.replace(/\\D/g,"").slice(0,10); scheduleSave(); }) + ' />' + (errs.agentMobile ? '<span class="err-msg">' + errs.agentMobile + "</span>" : "") + "</div>" +
        '<div class="f s2"><label>Is this a new or existing merchant?</label><div class="rg">' +
          '<label><input type="radio" name="merchantStatus" ' + (state.merchantStatus === "new" ? "checked" : "") + " " + on("change", () => { state.merchantStatus = "new"; rerender(); }) + " /> New Merchant</label>" +
          '<label><input type="radio" name="merchantStatus" ' + (state.merchantStatus === "existing" ? "checked" : "") + " " + on("change", () => { state.merchantStatus = "existing"; rerender(); }) + " /> Existing Merchant</label>" +
        "</div></div>" +
        (state.merchantStatus === "new"
          ? '<div class="f s2"><label>Establishment / Merchant Display Name *</label><input class="' + (errs.merchantDisplayName ? "err" : "") + '" value="' + attr(state.merchantDisplayName) + '" ' + on("input", (e) => { state.merchantDisplayName = e.target.value; scheduleSave(); }) + ' />' + (errs.merchantDisplayName ? '<span class="err-msg">' + errs.merchantDisplayName + "</span>" : "") + "</div>"
          : '<div class="f s2"><label>Merchant ID *</label><input class="' + (errs.merchantId ? "err" : "") + '" value="' + attr(state.merchantId) + '" ' + on("input", (e) => { state.merchantId = e.target.value; scheduleSave(); }) + ' />' + (errs.merchantId ? '<span class="err-msg">' + errs.merchantId + "</span>" : "") + "</div>") +
      "</div></div></div>" +
      '<div class="act"><div></div><button class="btn btn-p" ' + on("click", () => { if (validateStep0()) { state.step = 1; rerender(); } }) + ">Next: Platform & Entity \u2192</button></div>";
  }

  // ---- STEP 1: Platform + Entity ----
  if (state.step === 1) {
    return '<div class="card"><div class="chd"><h2>\ud83c\udfe2 Onboarding Platform & Entity Type</h2><span class="badge">Step 2 of 6</span></div><div class="cbd">' +
      '<div class="divider">Onboarding Platform</div><div class="rg">' +
        '<label><input type="radio" name="platform" ' + (state.onboardingType === "ace" ? "checked" : "") + " " + on("change", () => { state.onboardingType = "ace"; state.docRequirement = null; rerender(); }) + " /> ACE</label>" +
        '<label><input type="radio" name="platform" ' + (state.onboardingType === "salesforce" ? "checked" : "") + " " + on("change", () => { state.onboardingType = "salesforce"; state.docRequirement = null; rerender(); }) + " /> Salesforce</label>" +
      "</div>" +
      '<div class="divider">Entity Type</div><div class="rg">' +
        '<label><input type="radio" name="entity" ' + (state.entityType === "partnership" ? "checked" : "") + " " + on("change", () => { state.entityType = "partnership"; state.docRequirement = null; syncDefaultDesignations(); rerender(); }) + " /> Partnership Firm</label>" +
        '<label><input type="radio" name="entity" ' + (state.entityType === "company" ? "checked" : "") + " " + on("change", () => { state.entityType = "company"; state.docRequirement = null; syncDefaultDesignations(); rerender(); }) + " /> Company</label>" +
      "</div></div></div>" +
      '<div class="act"><button class="btn btn-s" ' + on("click", () => { state.step = 0; rerender(); }) + '>\u2190 Back</button>' +
      '<button class="btn btn-p" ' + (state.onboardingType && state.entityType ? "" : "disabled") + " " + on("click", () => { if (state.onboardingType && state.entityType) { state.step = 2; rerender(); } }) + ">Next: Document Set \u2192</button></div>";
  }

  // ---- STEP 2: Document requirement ----
  if (state.step === 2) {
    const options = (state.onboardingType && state.entityType) ? docOptionsMap[state.onboardingType][state.entityType] : [];
    return '<div class="card"><div class="chd"><h2>\ud83d\udcc2 Which documents do you need?</h2><span class="badge">Step 3 of 6</span></div><div class="cbd">' +
      '<div class="rg" style="flex-direction:column">' + options.map((opt) =>
        '<label><input type="radio" name="docReq" ' + (state.docRequirement === opt.id ? "checked" : "") + " " + on("change", () => { state.docRequirement = opt.id; state.formValidated = false; rerender(); }) + " /> " + opt.label + "</label>"
      ).join("") + "</div></div></div>" +
      '<div class="act"><button class="btn btn-s" ' + on("click", () => { state.step = 1; rerender(); }) + '>\u2190 Back</button>' +
      '<button class="btn btn-p" ' + (state.docRequirement ? "" : "disabled") + " " + on("click", () => { if (state.docRequirement) { state.step = 3; rerender(); } }) + ">Next: Fill Details \u2192</button></div>";
  }

  // ---- STEP 3: Data entry (firm, letterhead, partners, resolution, BO, MDF) ----
  if (state.step === 3) {
    const isFullKyc = needsFullKYC(state.docRequirement);
    const docLabel = (docOptionsMap[state.onboardingType][state.entityType].find(o => o.id === state.docRequirement) || {}).label || "";
    return (
      '<div class="card"><div class="chd"><h2>\ud83c\udfe2 Entity Details</h2><span class="badge">' + esc(docLabel) + '</span></div><div class="cbd">' +
        '<div class="info">This information is captured once and reused across every document you generate for this merchant.</div>' +
        '<div class="g2">' +
          '<div class="f s2"><label>Legal Entity / Firm Name (as on PAN) *</label><input class="' + (errs.firmName ? "err" : "") + '" value="' + attr(state.firmName) + '" ' + on("input", (e) => { state.firmName = e.target.value; scheduleSave(); }) + " /><span class=\"hint\">Must match Business PAN Card exactly.</span></div>" +
          '<div class="f s2"><label>Registered Office Address *</label><textarea class="' + (errs.regAddress ? "err" : "") + '" ' + on("input", (e) => { state.regAddress = e.target.value; scheduleSave(); }) + ">" + attr(state.regAddress) + "</textarea></div>" +
          (state.entityType === "partnership" ? (
            '<div class="f"><label>Entity Registration *</label><select ' + on("change", (e) => { state.partnershipRegType = e.target.value; rerender(); }) + ">" + option("registered","Registered Partnership / LLP",state.partnershipRegType) + option("unregistered","Unregistered Partnership Firm",state.partnershipRegType) + "</select></div>"
          ) : "") +
          '<div class="f"><label>Principal Place of Operation *</label><div class="rg">' +
            '<label><input type="radio" name="pSame" ' + (state.principalSame === "same" ? "checked" : "") + " " + on("change", () => { state.principalSame = "same"; rerender(); }) + " /> Same</label>" +
            '<label><input type="radio" name="pSame" ' + (state.principalSame === "diff" ? "checked" : "") + " " + on("change", () => { state.principalSame = "diff"; rerender(); }) + " /> Different</label>" +
          "</div>" + (state.principalSame === "diff" ? '<input style="margin-top:8px" value="' + attr(state.principalAddress) + '" ' + on("input", (e) => { state.principalAddress = e.target.value; scheduleSave(); }) + " />" : "") + "</div>" +
          (state.entityType === "partnership" ? '<div class="f"><label>Partnership Deed Date</label><input type="date" value="' + attr(state.deedDate) + '" ' + on("change", (e) => { state.deedDate = e.target.value; scheduleSave(); }) + " /></div>" : "") +
        "</div>" +
        '<div class="divider">Letterhead (optional \u2014 embedded in the Word document header)</div>' +
        '<div class="cg" style="margin-bottom:10px"><label><input type="checkbox" ' + (state.includeLetterhead ? "checked" : "") + " " + on("change", (e) => { state.includeLetterhead = e.target.checked; rerender(); }) + " /> Embed letterhead in generated documents</label></div>" +
        (state.includeLetterhead ? '<div class="g2">' +
          '<div class="f"><label>Tagline</label><input value="' + attr(state.lhTagline) + '" ' + on("input", (e) => { state.lhTagline = e.target.value; scheduleSave(); }) + " /></div>" +
          '<div class="f"><label>Phone</label><input value="' + attr(state.lhPhone) + '" ' + on("input", (e) => { state.lhPhone = e.target.value; scheduleSave(); }) + " /></div>" +
          '<div class="f"><label>Email</label><input value="' + attr(state.lhEmail) + '" ' + on("input", (e) => { state.lhEmail = e.target.value; scheduleSave(); }) + " /></div>" +
          '<div class="f"><label>Website</label><input value="' + attr(state.lhWebsite) + '" ' + on("input", (e) => { state.lhWebsite = e.target.value; scheduleSave(); }) + " /></div>" +
        "</div>" : "") +
      "</div></div>" +

      // Partners / Directors
      (needsResolution(state.docRequirement) || needsBO(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udc65 ' + (state.entityType === "company" ? "Directors" : "Partners") + '</h2></div><div class="cbd">' +
          '<div class="info">' + (isFullKyc ? "Full KYC required: PAN, DOB, Address, Proof of Address, % ownership. Total ownership must equal 100%." : "Basic details only for this document set.") + "</div>" +
          state.partners.map((pt, idx) => PartnerCardHTML(pt, idx, isFullKyc)).join("") +
          (errs.shareTotal ? '<div class="error-box">' + esc(errs.shareTotal) + "</div>" : "") +
          (errs.authSignatory ? '<div class="error-box">' + esc(errs.authSignatory) + "</div>" : "") +
          '<button type="button" class="add-btn" ' + on("click", addPartner) + ">\uff0b Add Another " + (state.entityType === "company" ? "Director" : "Partner") + "</button>" +
        "</div></div>"
      ) : "") +

      // Resolution details
      (needsResolution(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udccb Resolution Meeting Details</h2></div><div class="cbd">' +
          '<div class="g2">' +
            '<div class="f"><label>Meeting Date *</label><input type="date" class="' + (errs.resolutionDateRaw ? "err" : "") + '" value="' + attr(state.resolutionDateRaw) + '" ' + on("change", (e) => { state.resolutionDateRaw = e.target.value; state.boDate = state.boDate || e.target.value; scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Meeting Time *</label><input type="time" class="' + (errs.resolutionTimeRaw ? "err" : "") + '" value="' + attr(state.resolutionTimeRaw) + '" ' + on("change", (e) => { state.resolutionTimeRaw = e.target.value; scheduleSave(); }) + " /></div>" +
          "</div>" +
          '<div class="act" style="justify-content:flex-start"><button type="button" class="btn btn-s" ' + on("click", prefillDateTime) + ">Fill current date/time</button></div>" +
          '<div class="divider">Members Present & Signing</div>' +
          (errs.presentPartnerIds ? '<div class="error-box">' + esc(errs.presentPartnerIds) + "</div>" : "") +
          (errs.presentShare ? '<div class="error-box">' + esc(errs.presentShare) + "</div>" : "") +
          state.partners.map((pt) => '<label class="rg" style="justify-content:flex-start"><input type="checkbox" ' + (state.presentPartnerIds.includes(pt.id) ? "checked" : "") + " " + on("change", (e) => togglePresentPartner(pt.id, e.target.checked)) + " /> " + esc(pt.name || "(unnamed)") + (pt.isAS ? " \u2705" : "") + "</label>").join("") +
        "</div></div>"
      ) : "") +

      // BO details
      (needsBO(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udcdc BO Declaration Details</h2></div><div class="cbd">' +
          '<div class="g2">' +
            '<div class="f"><label>BO Declaration Date *</label><input type="date" class="' + (errs.boDate ? "err" : "") + '" value="' + attr(state.boDate) + '" ' + on("change", (e) => { state.boDate = e.target.value; scheduleSave(); }) + " /></div>" +
          "</div>" +
          '<div class="divider">Category</div><div class="rg">' +
            '<label><input type="radio" name="boCat" ' + (state.boCategory === "cat1" ? "checked" : "") + " " + on("change", () => { state.boCategory = "cat1"; rerender(); }) + " /> Category 1 (natural persons above threshold)</label>" +
            '<label><input type="radio" name="boCat" ' + (state.boCategory === "cat2" ? "checked" : "") + " " + on("change", () => { state.boCategory = "cat2"; rerender(); }) + " /> Category 2 (senior managing official)</label>" +
          "</div>" +
          '<div class="cg" style="margin-top:10px"><label><input type="checkbox" ' + (state.pepDeclarationBO ? "checked" : "") + " " + on("change", (e) => { state.pepDeclarationBO = e.target.checked; scheduleSave(); }) + ' /> Confirm: No one is a Politically Exposed Person (PEP)</label></div>' +
        "</div></div>"
      ) : "") +

      // MDF details
      (needsMDF(state.docRequirement) ? (
        '<div class="card"><div class="chd"><h2>\ud83d\udcc4 Merchant Declaration Form (MDF)</h2></div><div class="cbd">' +
          '<div class="g2">' +
            '<div class="f"><label>Signatory Name *</label><input class="' + (errs.mdfAuthName ? "err" : "") + '" value="' + attr(state.mdfAuthName) + '" ' + on("input", (e) => { state.mdfAuthName = e.target.value; scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Signatory Designation</label><input value="' + attr(state.mdfAuthDesignation) + '" ' + on("input", (e) => { state.mdfAuthDesignation = e.target.value; scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Signatory PAN *</label><input class="' + (errs.mdfAuthPan ? "err" : "") + '" maxlength="10" value="' + attr(state.mdfAuthPan) + '" ' + on("input", (e) => { state.mdfAuthPan = e.target.value.toUpperCase(); scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Mobile Number *</label><input type="tel" maxlength="10" class="' + (errs.mdfMobile ? "err" : "") + '" value="' + attr(state.mdfMobile) + '" ' + on("input", (e) => { state.mdfMobile = e.target.value.replace(/\\D/g,"").slice(0,10); scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Email</label><input type="email" value="' + attr(state.mdfEmail) + '" ' + on("input", (e) => { state.mdfEmail = e.target.value; scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>TAN Status</label><select ' + on("change", (e) => { state.mdfTanStatus = e.target.value; rerender(); }) + ">" + option("no_tan","Not liable for TAN",state.mdfTanStatus) + option("has_tan","Holds TAN",state.mdfTanStatus) + "</select></div>" +
            (state.mdfTanStatus === "has_tan" ? '<div class="f"><label>TAN Number</label><input value="' + attr(state.mdfTanNum) + '" ' + on("input", (e) => { state.mdfTanNum = e.target.value; scheduleSave(); }) + " /></div>" : "") +
            (state.onboardingType === "salesforce" ? (
              '<div class="f"><label>GST Status</label><select ' + on("change", (e) => { state.mdfGstStatus = e.target.value; rerender(); }) + ">" + option("no_gst","No GST Registration",state.mdfGstStatus) + option("has_gst","Holds GST",state.mdfGstStatus) + "</select></div>" +
              (state.mdfGstStatus === "has_gst" ? '<div class="f"><label>GSTIN</label><input value="' + attr(state.mdfGstNum) + '" ' + on("input", (e) => { state.mdfGstNum = e.target.value; scheduleSave(); }) + " /></div>" : "")
            ) : "") +
            '<div class="f"><label>Place</label><input value="' + attr(state.mdfPlace) + '" ' + on("input", (e) => { state.mdfPlace = e.target.value; scheduleSave(); }) + " /></div>" +
            '<div class="f"><label>Date</label><input type="date" value="' + attr(state.mdfDateRaw) + '" ' + on("change", (e) => { state.mdfDateRaw = e.target.value; scheduleSave(); }) + " /></div>" +
          "</div>" +
        "</div></div>"
      ) : "") +

      '<div class="act"><button class="btn btn-s" ' + on("click", () => { state.step = 2; rerender(); }) + '>\u2190 Back</button>' +
      '<button class="btn btn-p" ' + on("click", () => { validateDataEntry(); if (state.formValidated) { state.step = 4; rerender(); } }) + ">Review & Generate \u2192</button></div>"
    );
  }

  // ---- STEP 4: Review & Generate ----
  if (state.step === 4) {
    const docLabel = (docOptionsMap[state.onboardingType][state.entityType].find(o => o.id === state.docRequirement) || {}).label || "";
    return '<div class="card"><div class="chd"><h2>\u2705 Review</h2><span class="badge">Step 6 of 6</span></div><div class="cbd">' +
      '<table class="rtbl">' +
        '<tr class="hd"><td colspan="2">MERCHANT</td></tr>' +
        '<tr><td>Agent Mobile</td><td>' + esc(state.agentMobile) + "</td></tr>" +
        '<tr><td>' + (state.merchantStatus === "new" ? "Establishment Name" : "Merchant ID") + '</td><td>' + esc(state.merchantStatus === "new" ? state.merchantDisplayName : state.merchantId) + "</td></tr>" +
        '<tr class="hd"><td colspan="2">ENTITY</td></tr>' +
        '<tr><td>Documents</td><td><strong>' + esc(docLabel) + "</strong></td></tr>" +
        '<tr><td>Legal Name</td><td>' + esc(state.firmName) + "</td></tr>" +
        '<tr><td>Registered Address</td><td>' + esc(state.regAddress) + "</td></tr>" +
        (needsResolution(state.docRequirement) ? '<tr><td>Authorised Signatory</td><td>' + esc(formData.authSignatoryName) + "</td></tr>" : "") +
      "</table></div></div>" +
      '<div class="gen-box">' +
        '<h3>\ud83c\udf89 Ready to Generate</h3><p>Download real Word (.docx) files, or use Print to save as PDF. Print &amp; sign on the firm/company letterhead if not embedded.</p>' +
        (needsResolution(state.docRequirement) ? '<div class="doc-group"><h4>Resolution</h4><div class="doc-btns"><button class="btn btn-p" ' + on("click", () => generateDocx("resolution")) + '>\ud83d\udcc4 Download .docx</button><button class="btn btn-s" ' + on("click", () => generatePrint("resolution")) + ">\ud83d\uddb6 Print / PDF</button></div></div>" : "") +
        (needsBO(state.docRequirement) ? '<div class="doc-group"><h4>BO Declaration</h4><div class="doc-btns"><button class="btn btn-g" ' + on("click", () => generateDocx("bo")) + '>\ud83d\udccb Download .docx</button><button class="btn btn-s" ' + on("click", () => generatePrint("bo")) + ">\ud83d\uddb6 Print / PDF</button></div></div>" : "") +
        (needsMDF(state.docRequirement) ? '<div class="doc-group"><h4>MDF</h4><div class="doc-btns"><button class="btn btn-gold" ' + on("click", () => generateDocx("mdf")) + '>\ud83d\udcc4 Download .docx</button><button class="btn btn-s" ' + on("click", () => generatePrint("mdf")) + ">\ud83d\uddb6 Print / PDF</button></div></div>" : "") +
      "</div>" +
      '<div class="act"><button class="btn btn-s" ' + on("click", () => { state.step = 3; rerender(); }) + '>\u2190 Back to Edit</button>' +
      '<button class="btn btn-red" ' + on("click", clearMerchantData) + ">\ud83d\uddd1 Start New Merchant</button></div>";
  }

  return "";
}

// ---- INIT ----
loadAgent();
loadMerchant();
rerender();
