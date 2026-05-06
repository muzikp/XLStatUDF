const translations = {
  en: {
    title: "StatLab for Excel",
    subtitle: "Statistical functions, demos, options, and wizards in one Excel workspace.",
    searchLabel: "Search functions",
    searchPlaceholder: "Function name, parameter, category, or description",
    hint: "Tip: functions are registered without a namespace in the current Office.js build.",
    loading: "Loading function documentation...",
    empty: "No matching function found.",
    copied: "Formula copied.",
    copy: "Copy formula",
    pending: "Pending migration",
    optional: "optional",
    noParameters: "No parameters.",
    parameters: "Parameters",
    parameterName: "Name",
    parameterType: "Type",
    parameterDescription: "Description",
    output: "Output",
    noOutput: "Output is not documented yet.",
    formula: "Formula",
    tutorial: "Demo",
    tutorialReady: "Tutorial sheet created.",
    tutorialUnavailable: "Tutorial is available only inside Excel.",
    copyLog: "Copy log",
    logCopied: "Log copied.",
    required: "required",
    wizard: "Wizard",
    wizardClose: "Close",
    wizardInsert: "Insert formula",
    wizardUseSelection: "Use selection",
    wizardUseVariable: "Manual range",
    wizardDatasetLoaded: "Variables loaded from sheet.",
    wizardTarget: "Target cell",
    wizardReady: "Formula inserted.",
    wizardTitle: "Function wizard",
    wizardMissingRange: "Select the category and value ranges first.",
    wizardNoFormula: "The selected cell does not contain this function.",
    wizardAlphaInvalid: "Alpha must be a number from 0 to 1.",
    wizardNumberInvalid: "Enter a valid number.",
    wizardPositiveInvalid: "Enter a positive number.",
    wizardProbabilityInvalid: "Enter a number from 0 to 1.",
    wizardValueMissing: "This value is required.",
    wizardRangeInvalid: "This range is not valid in the active workbook.",
    wizardRangeTooLarge: "Select the actual data cells, not an entire row or column.",
    wizardTargetInvalid: "Target must be a single valid cell.",
    wizardRangeShapeInvalid: "Category and value ranges must contain the same number of cells.",
    functionsTab: "Functions",
    settingsTab: "Settings",
    settingsLanguage: "Language",
    settingsLanguageHint: "Changing language also recalculates workbook formulas.",
    settingsAccount: "Account",
    authDevLogin: "Simulate login",
    authLogout: "Sign out",
    authHint: "This local development login will be replaced by CVUT OAuth.",
    authSignedIn: "Signed in locally as",
    authSignedOut: "Not signed in. Protected functions return #UNAUTHORIZED.",
    authExpired: "Local login expired. Protected functions return #UNAUTHORIZED.",
    authRequired: "Login required"
  },
  cs: {
    title: "StatLab pro Excel",
    subtitle: "Statistické funkce, dema, volby a průvodci v jednom excelovém prostoru.",
    searchLabel: "Hledat funkce",
    searchPlaceholder: "Název funkce, parametr, kategorie nebo popis",
    hint: "Tip: v aktuálním Office.js buildu jsou funkce registrované bez namespace.",
    loading: "Načítám dokumentaci funkcí...",
    empty: "Nenalezena žádná odpovídající funkce.",
    copied: "Vzorec zkopírován.",
    copy: "Kopírovat vzorec",
    pending: "Čeká na migraci",
    optional: "volitelné",
    noParameters: "Bez parametrů.",
    parameters: "Parametry",
    parameterName: "Název",
    parameterType: "Typ",
    parameterDescription: "Popis",
    output: "Výstup",
    noOutput: "Výstup zatím není zdokumentovaný.",
    formula: "Vzorec",
    tutorial: "Demo",
    tutorialReady: "Tutorialový list byl vytvořen.",
    tutorialUnavailable: "Tutorial je dostupný jen uvnitř Excelu.",
    copyLog: "Kopírovat log",
    logCopied: "Log zkopírován.",
    required: "povinné",
    wizard: "Wizard",
    wizardClose: "Zavřít",
    wizardInsert: "Vložit vzorec",
    wizardUseSelection: "Použít výběr",
    wizardUseVariable: "Ruční rozsah",
    wizardDatasetLoaded: "Proměnné načteny z listu.",
    wizardTarget: "Cílová buňka",
    wizardReady: "Vzorec byl vložen.",
    wizardTitle: "Průvodce funkcí",
    wizardMissingRange: "Nejprve vyberte rozsah kategorií a hodnot.",
    wizardNoFormula: "Vybraná buňka neobsahuje tuto funkci.",
    wizardAlphaInvalid: "Alfa musí být číslo od 0 do 1.",
    wizardNumberInvalid: "Zadejte platné číslo.",
    wizardPositiveInvalid: "Zadejte kladné číslo.",
    wizardProbabilityInvalid: "Zadejte číslo od 0 do 1.",
    wizardValueMissing: "Tato hodnota je povinná.",
    wizardRangeInvalid: "Tento rozsah není v aktivním sešitu platný.",
    wizardRangeTooLarge: "Vyberte konkrétní buňky s daty, ne celý řádek nebo sloupec.",
    wizardTargetInvalid: "Cíl musí být jedna platná buňka.",
    wizardRangeShapeInvalid: "Rozsahy kategorií a hodnot musí mít stejný počet buněk.",
    functionsTab: "Funkce",
    settingsTab: "Nastavení",
    settingsLanguage: "Jazyk",
    settingsLanguageHint: "Prepnuti jazyka zaroven prepocita vzorce v sesitu.",
    settingsAccount: "Ucet",
    authDevLogin: "Simulovat prihlaseni",
    authLogout: "Odhlasit",
    authHint: "Toto lokalni vyvojove prihlaseni pozdeji nahradi CVUT OAuth.",
    authSignedIn: "Lokalne prihlasen jako",
    authSignedOut: "Neprihlaseno. Chranene funkce vraci #UNAUTHORIZED.",
    authExpired: "Lokalni prihlaseni vyprselo. Chranene funkce vraci #UNAUTHORIZED.",
    authRequired: "Vyžaduje přihlášení"
  }
};

const languageStorageKey = "evalytics.language";
const authStorageKey = "evalytics.auth.snapshot";

function normalizeLanguage(value) {
  if (!value) {
    return null;
  }
  const lower = String(value).toLowerCase();
  if (lower.startsWith("cs")) {
    return "cs";
  }
  if (lower.startsWith("en")) {
    return "en";
  }
  return null;
}

function initialLanguage() {
  try {
    const stored = normalizeLanguage(localStorage.getItem(languageStorageKey));
    if (stored) {
      return stored;
    }
  } catch {
    // localStorage can be unavailable in some Office webview modes.
  }
  return normalizeLanguage(navigator.language) ?? "en";
}

let language = initialLanguage();
let functions = [];
let localizedDocs = { cs: {}, en: {} };
let search = "";
let activeTab = "functions-tab";
let wizardInsertInProgress = false;
const buildStamp = window.EVALYTICS_BUILD_STAMP ?? "dev";
const appVersion = "1.0.0";
const buildVersion = /^\d+$/.test(String(buildStamp)) ? `${appVersion}.${buildStamp}` : appVersion;
const hiddenDocumentationFunctions = new Set(["VERSION", "PING"]);
const wizardExcludedFunctions = new Set();
const datasetWizardFunctions = new Set(["ANCOVA"]);
const excelMaxRows = 1048576;
const excelMaxColumns = 16384;
const actionIcons = {
  demo: "▶",
  wizard: "ƒ"
};

function t(key) {
  return translations[language][key];
}

function translateStaticText() {
  document.querySelectorAll("[data-i18n]").forEach((node) => {
    node.textContent = t(node.dataset.i18n);
  });

  document.querySelectorAll("[data-i18n-placeholder]").forEach((node) => {
    node.placeholder = t(node.dataset.i18nPlaceholder);
  });

  document.querySelectorAll("[data-lang]").forEach((node) => {
    node.classList.toggle("active", node.dataset.lang === language);
  });
}

function localizedInfo(fn) {
  return localizedDocs[language]?.[fn.name] ?? localizedDocs.en?.[fn.name] ?? {};
}

function switchTab(tabId) {
  activeTab = tabId;
  document.querySelectorAll("[data-tab-target]").forEach((button) => {
    button.classList.toggle("active", button.dataset.tabTarget === activeTab);
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    const isActive = panel.id === activeTab;
    panel.hidden = !isActive;
    panel.classList.toggle("active", isActive);
  });
}

async function persistLanguage(nextLanguage) {
  try {
    localStorage.setItem(languageStorageKey, nextLanguage);
  } catch {
    // The sidepanel still updates even if persistent storage is blocked.
  }

  await recalculateWorkbook();
}

async function recalculateWorkbook() {
  if (window.Excel) {
    await Excel.run(async (context) => {
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
  }
}

function readAuthSnapshot() {
  try {
    const raw = localStorage.getItem(authStorageKey);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function authExpiry(snapshot) {
  if (!snapshot?.expiresAt) {
    return 0;
  }
  if (typeof snapshot.expiresAt === "number") {
    return snapshot.expiresAt;
  }
  const parsed = Date.parse(snapshot.expiresAt);
  return Number.isFinite(parsed) ? parsed : 0;
}

async function setSharedAuthSnapshot(snapshot) {
  try {
    const runtime = window.OfficeRuntime;
    if (runtime?.storage?.setItem) {
      await runtime.storage.setItem(authStorageKey, JSON.stringify(snapshot));
    }
  } catch {
    // localStorage remains the fallback.
  }
}

async function clearSharedAuthSnapshot() {
  try {
    const runtime = window.OfficeRuntime;
    if (runtime?.storage?.removeItem) {
      await runtime.storage.removeItem(authStorageKey);
    }
  } catch {
    // localStorage remains the fallback.
  }
}

function isAuthActive(snapshot = readAuthSnapshot()) {
  return Boolean(snapshot && authExpiry(snapshot) > Date.now());
}

function updateAuthStatus() {
  const status = document.getElementById("auth-status");
  if (!status) {
    return;
  }

  const snapshot = readAuthSnapshot();
  if (!snapshot) {
    status.textContent = t("authSignedOut");
    return;
  }
  if (!isAuthActive(snapshot)) {
    status.textContent = t("authExpired");
    return;
  }

  const label = snapshot.user?.displayName || snapshot.user?.email || snapshot.user?.id || "local-dev-user";
  status.textContent = `${t("authSignedIn")} ${label}`;
}

async function writeAuthSnapshot(snapshot) {
  localStorage.setItem(authStorageKey, JSON.stringify(snapshot));
  await setSharedAuthSnapshot(snapshot);
  updateAuthStatus();
  render();
  await recalculateWorkbook();
}

async function simulateAuthLogin() {
  const expiresAt = Date.now() + (12 * 60 * 60 * 1000);
  await writeAuthSnapshot({
    provider: "local-dev",
    user: {
      id: "local-dev-user",
      displayName: "Local Dev User",
      email: "local-dev@example.test"
    },
    entitlements: {
      allowedFunctions: ["*"],
      deniedFunctions: []
    },
    issuedAt: Date.now(),
    expiresAt
  });
  debugLog("Auth dev login simulated", { expiresAt });
}

async function logoutAuth() {
  localStorage.removeItem(authStorageKey);
  await clearSharedAuthSnapshot();
  updateAuthStatus();
  render();
  await recalculateWorkbook();
  debugLog("Auth logged out");
}

function signatureFor(fn) {
  const args = fn.parameters.map((parameter) => (parameter.optional ? `[${parameter.name}]` : parameter.name));
  return `=${fn.name}(${args.join("; ")})`;
}

function normalizeFunction(fn) {
  return {
    ...fn,
    signature: signatureFor(fn),
    pending: fn.description.toLowerCase().includes("pending")
  };
}

function displaySummary(fn) {
  return localizedInfo(fn).summary ?? fn.description;
}

function displayCategory(fn) {
  const category = localizedInfo(fn).category ?? "Other";
  if (language === "cs") {
    if (category === "Obecn?" || category === "Obecne") {
      return "Obecné";
    }
    if (category === "Popisn?" || category === "Popisne") {
      return "Popisné";
    }
  }
  return category;
}

function parameterDescription(fn, index) {
  return localizedInfo(fn).parameters?.[index]?.description ?? "";
}

function parameterTypeLabel(parameter) {
  if (parameter.dimensionality === "matrix") {
    return language === "cs" ? "řada hodnot" : "value range";
  }

  const labels = {
    cs: {
      any: "hodnota",
      number: "číslo",
      string: "text",
      boolean: "pravda/nepravda"
    },
    en: {
      any: "value",
      number: "number",
      string: "text",
      boolean: "true/false"
    }
  };

  return labels[language][parameter.type] ?? parameter.type;
}

function outputDocs(fn) {
  return localizedInfo(fn).output ?? [];
}

function categoryRank(category) {
  const order = {
    en: ["General", "Descriptive", "Tests", "Pivot table", "Diagnostics", "Other"],
    cs: ["Obecné", "Popisné", "Testy", "Kontingenční tabulka", "Diagnostika", "Other"]
  };

  const index = order[language].indexOf(category);
  return index === -1 ? order[language].length : index;
}

function matchesSearch(fn) {
  if (!search) {
    return true;
  }

  const haystack = [
    fn.name,
    fn.description,
    displaySummary(fn),
    displayCategory(fn),
    fn.signature,
    ...fn.parameters.map((parameter) => parameter.name)
  ].join(" ").toLowerCase();

  return haystack.includes(search.toLowerCase());
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function debugLog(message, detail) {
  const target = document.getElementById("debug-log");
  if (!target) {
    return;
  }

  const time = new Date().toLocaleTimeString();
  const detailText = detail ? ` ${typeof detail === "string" ? detail : JSON.stringify(detail)}` : "";
  target.textContent = `[${time}] ${message}${detailText}\n${target.textContent}`.slice(0, 8000);
}

async function copyDebugLog() {
  const target = document.getElementById("debug-log");
  if (!target) {
    return;
  }

  await copyText(target.textContent ?? "");
  setStatus(t("logCopied"));
}

function sleep(milliseconds) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

function textLooksBusy(value) {
  return String(value ?? "").toUpperCase().includes("BUSY");
}

async function copyText(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.left = "-999px";
  document.body.append(textarea);
  textarea.select();
  document.execCommand("copy");
  textarea.remove();
}

function renderParameters(container, fn) {
  container.replaceChildren();

  const title = document.createElement("h3");
  title.textContent = t("parameters");
  container.append(title);

  if (fn.parameters.length === 0) {
    const empty = document.createElement("p");
    empty.textContent = t("noParameters");
    container.append(empty);
    return;
  }

  const table = document.createElement("table");
  table.className = "parameter-table";
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr><th>${t("parameterName")}</th><th>${t("parameterType")}</th><th>${t("parameterDescription")}</th></tr>`;
  table.append(thead);

  const tbody = document.createElement("tbody");
  fn.parameters.forEach((parameter, index) => {
    const row = document.createElement("tr");
    row.classList.toggle("optional", Boolean(parameter.optional));
    row.classList.toggle("required", !parameter.optional);

    const name = document.createElement("td");
    name.className = "parameter-name";
    const nameText = document.createElement("span");
    nameText.textContent = parameter.name;
    name.append(nameText);
    if (!parameter.optional) {
      const required = document.createElement("span");
      required.className = "required-star";
      required.title = t("required");
      required.textContent = "*";
      name.append(required);
    }
    row.append(name);

    const meta = document.createElement("td");
    meta.className = "parameter-meta";
    meta.textContent = parameterTypeLabel(parameter);
    row.append(meta);

    const description = document.createElement("td");
    description.className = "parameter-description";
    description.textContent = parameterDescription(fn, index);
    const enumValues = localizedInfo(fn).parameters?.[index]?.enumValues ?? [];
    if (enumValues.length > 0) {
      const list = document.createElement("dl");
      list.className = "enum-list";
      enumValues.forEach((item) => {
        const term = document.createElement("dt");
        term.classList.toggle("default-enum", Boolean(item.default));
        term.textContent = item.value;
        const definition = document.createElement("dd");
        definition.textContent = item.meaning;
        list.append(term, definition);
      });
      description.append(list);
    }
    row.append(description);
    tbody.append(row);
  });

  table.append(tbody);
  container.append(table);

}

function addTutorialButton(actions, fn) {
  if (!tutorialDefinitions[fn.name]) {
    return;
  }

  const button = document.createElement("button");
  button.className = "action-icon-button demo-button";
  button.type = "button";
  button.textContent = actionIcons.demo;
  button.title = t("tutorial");
  button.setAttribute("aria-label", t("tutorial"));
  button.addEventListener("click", () => runTutorial(fn.name));
  actions.append(button);
}

function addWizardButton(actions, fn) {
  if (fn.parameters.length === 0 || hiddenDocumentationFunctions.has(fn.name) || wizardExcludedFunctions.has(fn.name)) {
    return;
  }

  const button = document.createElement("button");
  button.className = "action-icon-button wizard-button";
  button.type = "button";
  button.textContent = actionIcons.wizard;
  button.title = t("wizard");
  button.setAttribute("aria-label", t("wizard"));
  button.addEventListener("click", () => {
    openFunctionWizard(fn).catch((error) => {
      setStatus(error instanceof Error ? error.message : String(error));
      debugLog("Wizard open failed", { functionName: fn.name, message: error instanceof Error ? error.message : String(error) });
    });
  });
  actions.append(button);
}

const tutorialDefinitions = {
  "VERSION": {
    header: "value",
    formula: "=VERSION()",
    range: "A2:A2"
  },
  "PING": {
    header: "value",
    formula: "=PING()",
    range: "A2:A2"
  },
  "GENERATE.NORM": {
    header: "value",
    formula: "=GENERATE.NORM(178,8)",
    range: "A2:A2"
  },
  "GENERATE.NORM.ARRAY": {
    header: "value",
    formula: "=GENERATE.NORM.ARRAY(100,178,8)",
    range: "A2:A101"
  },
  "GENERATE.INT": {
    header: "value",
    formula: "=GENERATE.INT(1,6)",
    range: "A2:A2"
  },
  "GENERATE.INT.ARRAY": {
    header: "value",
    formula: "=GENERATE.INT.ARRAY(100,1,6)",
    range: "A2:A101"
  },
  "FILL": {
    header: "group",
    formula: '=FILL("control",100)',
    range: "A2:A101"
  },
  "STACK": {
    customRun: runStackGTutorial
  },
  "UNSTACK": {
    customRun: runUnstackGTutorial
  },
  "PARSE.NUMBER": {
    customRun: runParseNumberTutorial
  },
  "NORM.DIST.RANGE": {
    header: "probability",
    formula: "=NORM.DIST.RANGE(0,1,-1,1)",
    range: "A2:A2"
  },
  "AVERAGE.W": {
    customRun: () => runDescriptiveTutorial("AVERAGE.W")
  },
  "HARMEAN.W": {
    customRun: () => runDescriptiveTutorial("HARMEAN.W")
  },
  "GEOMEAN.W": {
    customRun: () => runDescriptiveTutorial("GEOMEAN.W")
  },
  "VAR.P.W": {
    customRun: () => runDescriptiveTutorial("VAR.P.W")
  },
  "VAR.S.W": {
    customRun: () => runDescriptiveTutorial("VAR.S.W")
  },
  "STDEV.P.W": {
    customRun: () => runDescriptiveTutorial("STDEV.P.W")
  },
  "STDEV.S.W": {
    customRun: () => runDescriptiveTutorial("STDEV.S.W")
  },
  "VARCOEF": {
    customRun: () => runDescriptiveTutorial("VARCOEF")
  },
  "VARCOEF.S": {
    customRun: () => runDescriptiveTutorial("VARCOEF.S")
  },
  "VARCOEF.W": {
    customRun: () => runDescriptiveTutorial("VARCOEF.W")
  },
  "VARCOEF.S.W": {
    customRun: () => runDescriptiveTutorial("VARCOEF.S.W")
  },
  "PIVOT.COUNT": {
    customRun: () => runPivotTutorial("PIVOT.COUNT")
  },
  "PIVOT.SUM": {
    customRun: () => runPivotTutorial("PIVOT.SUM")
  },
  "PIVOT.AVERAGE": {
    customRun: () => runPivotTutorial("PIVOT.AVERAGE")
  },
  "PIVOT.MIN": {
    customRun: () => runPivotTutorial("PIVOT.MIN")
  },
  "PIVOT.MAX": {
    customRun: () => runPivotTutorial("PIVOT.MAX")
  },
  "PIVOT.MEDIAN": {
    customRun: () => runPivotTutorial("PIVOT.MEDIAN")
  },
  "PIVOT.PERCENTILE": {
    customRun: () => runPivotTutorial("PIVOT.PERCENTILE")
  },
  "PIVOT.STDEV.S": {
    customRun: () => runPivotTutorial("PIVOT.STDEV.S")
  },
  "PIVOT.STDEV.P": {
    customRun: () => runPivotTutorial("PIVOT.STDEV.P")
  },
  "PIVOT.VAR.S": {
    customRun: () => runPivotTutorial("PIVOT.VAR.S")
  },
  "PIVOT.VAR.P": {
    customRun: () => runPivotTutorial("PIVOT.VAR.P")
  },
  "PIVOT.VARCOEF.S": {
    customRun: () => runPivotTutorial("PIVOT.VARCOEF.S")
  },
  "PIVOT.VARCOEF.P": {
    customRun: () => runPivotTutorial("PIVOT.VARCOEF.P")
  },
  "PIVOT.CONF.T": {
    customRun: () => runPivotTutorial("PIVOT.CONF.T")
  },
  "PIVOT.CONF.NORM": {
    customRun: () => runPivotTutorial("PIVOT.CONF.NORM")
  },
  "PIVOT.MAD": {
    customRun: () => runPivotTutorial("PIVOT.MAD")
  },
  "PIVOT.IQR": {
    customRun: () => runPivotTutorial("PIVOT.IQR")
  },
  "SHAPIRO.WILK": {
    customRun: () => runMatrixTutorial("SHAPIRO.WILK")
  },
  "KOLMOGOROV.SMIRNOV": {
    customRun: () => runMatrixTutorial("KOLMOGOROV.SMIRNOV")
  },
  "T.TEST.1S": {
    customRun: () => runMatrixTutorial("T.TEST.1S")
  },
  "PROP.TEST.1S": {
    customRun: () => runMatrixTutorial("PROP.TEST.1S")
  },
  "WILCOXON.PAIRED": {
    customRun: () => runMatrixTutorial("WILCOXON.PAIRED")
  },
  "MANN.WHITNEY": {
    customRun: () => runMatrixTutorial("MANN.WHITNEY")
  },
  "CHISQ.GOF": {
    customRun: () => runMatrixTutorial("CHISQ.GOF")
  },
  "ANOVA": {
    customRun: () => runMatrixTutorial("ANOVA")
  },
  "CORREL.SPEARMAN": {
    customRun: () => runMatrixTutorial("CORREL.SPEARMAN")
  },
  "ANOVA.RM": {
    customRun: () => runMatrixTutorial("ANOVA.RM")
  },
  "ANCOVA": {
    customRun: () => runMatrixTutorial("ANCOVA")
  },
  "CONTINGENCY.T": {
    customRun: () => runMatrixTutorial("CONTINGENCY.T")
  },
  "CONTINGENCY": {
    customRun: () => runMatrixTutorial("CONTINGENCY")
  },
  "CORREL.MATRIX": {
    customRun: () => runMatrixTutorial("CORREL.MATRIX")
  },
  "WELCH.TEST.2S": {
    customRun: runWelchTutorialWithDiagnostics
  }
};

const descriptiveValueRows = [
  [12],
  [15],
  [18],
  [21],
  [23],
  [26],
  [29],
  [31],
  [34],
  [37]
];

const descriptiveWeightedRows = [
  [12, 1],
  [15, 2],
  [18, 1],
  [21, 3],
  [23, 2],
  [26, 2],
  [29, 1],
  [31, 3],
  [34, 2],
  [37, 1]
];

const descriptiveTutorialDefinitions = {
  "AVERAGE.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=AVERAGE.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "HARMEAN.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=HARMEAN.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "GEOMEAN.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=GEOMEAN.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "VAR.P.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=VAR.P.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "VAR.S.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=VAR.S.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "STDEV.P.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=STDEV.P.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "STDEV.S.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=STDEV.S.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "VARCOEF": {
    headers: ["value", "", "result"],
    rows: descriptiveValueRows,
    formulaCell: "C2",
    formula: "=VARCOEF(A2:A11)",
    previewRange: "C2:C2"
  },
  "VARCOEF.S": {
    headers: ["value", "", "result"],
    rows: descriptiveValueRows,
    formulaCell: "C2",
    formula: "=VARCOEF.S(A2:A11)",
    previewRange: "C2:C2"
  },
  "VARCOEF.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=VARCOEF.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
  "VARCOEF.S.W": {
    headers: ["value", "weight", "", "result"],
    rows: descriptiveWeightedRows,
    formulaCell: "D2",
    formula: "=VARCOEF.S.W(A2:A11,B2:B11)",
    previewRange: "D2:D2"
  },
};

const pivotTutorialRowBlocks = [
  { cell: "A2", formula: '=FILL("north",12)' },
  { cell: "A14", formula: '=FILL("south",12)' },
  { cell: "A26", formula: '=FILL("west",12)' }
];

const pivotTutorialColumnBlocks = [
  { cell: "B2", formula: '=FILL("online",6)' },
  { cell: "B8", formula: '=FILL("store",6)' },
  { cell: "B14", formula: '=FILL("online",6)' },
  { cell: "B20", formula: '=FILL("store",6)' },
  { cell: "B26", formula: '=FILL("online",6)' },
  { cell: "B32", formula: '=FILL("store",6)' }
];

const pivotTutorialValueFormula = "=GENERATE.NORM.ARRAY(36,100,15)";

function pivotTutorialFormula(functionName) {
  if (functionName === "PIVOT.PERCENTILE") {
    return "=PIVOT.PERCENTILE(A1:A37,B1:B37,C1:C37,0.75)";
  }
  if (functionName === "PIVOT.CONF.T" || functionName === "PIVOT.CONF.NORM") {
    return `=${functionName}(A1:A37,B1:B37,C1:C37,0.05,0)`;
  }
  return `=${functionName}(A1:A37,B1:B37,C1:C37)`;
}

const matrixTutorialDefinitions = {
  "SHAPIRO.WILK": {
    values: [
      ["value"],
      [171],
      [174],
      [176],
      [178],
      [179],
      [181],
      [183],
      [185],
      [187],
      [190]
    ],
    formula: "=SHAPIRO.WILK(A1:A11,1)",
    previewRange: "F2:G5"
  },
  "KOLMOGOROV.SMIRNOV": {
    values: [
      ["value"],
      [171],
      [174],
      [176],
      [178],
      [179],
      [181],
      [183],
      [185],
      [187],
      [190]
    ],
    formula: "=KOLMOGOROV.SMIRNOV(A1:A11,0,1)",
    previewRange: "F2:G5"
  },
  "T.TEST.1S": {
    values: [
      ["value"],
      [171],
      [174],
      [176],
      [178],
      [179],
      [181],
      [183],
      [185],
      [187],
      [190]
    ],
    formula: "=T.TEST.1S(A1:A11,175,0,0.05,1)",
    previewRange: "F2:G12"
  },
  "PROP.TEST.1S": {
    values: [
      ["success"],
      [1],
      [0],
      [1],
      [1],
      [0],
      [1],
      [1],
      [1],
      [0],
      [1]
    ],
    formula: "=PROP.TEST.1S(A1:A11,0.5,0,0.05,1)",
    previewRange: "F2:G11"
  },
  "WILCOXON.PAIRED": {
    values: [
      ["before", "after"],
      [18, 16],
      [21, 19],
      [20, 18],
      [24, 22],
      [23, 21],
      [25, 23],
      [22, 21],
      [26, 24]
    ],
    formula: "=WILCOXON.PAIRED(A1:A9,B1:B9,1,0.05,0)",
    previewRange: "F2:G13"
  },
  "MANN.WHITNEY": {
    values: [
      ["group", "value"],
      ["control", 12],
      ["control", 14],
      ["control", 15],
      ["control", 17],
      ["control", 18],
      ["treatment", 20],
      ["treatment", 21],
      ["treatment", 23],
      ["treatment", 24],
      ["treatment", 26]
    ],
    formula: "=MANN.WHITNEY(A1:A11,B1:B11,1,0.05,0)",
    previewRange: "F2:L16"
  },
  "CHISQ.GOF": {
    values: [
      ["category", "observed", "expected"],
      ["A", 18, 20],
      ["B", 22, 20],
      ["C", 15, 20],
      ["D", 25, 20]
    ],
    formula: "=CHISQ.GOF(B1:B5,C1:C5,A1:A5,0.05,1)",
    previewRange: "F2:I14"
  },
  "ANOVA": {
    values: [
      ["group", "value"],
      ["A", 10],
      ["A", 12],
      ["A", 13],
      ["A", 11],
      ["B", 16],
      ["B", 18],
      ["B", 19],
      ["B", 17],
      ["C", 20],
      ["C", 22],
      ["C", 24],
      ["C", 21]
    ],
    formula: "=ANOVA(A1:A13,B1:B13,1,0.05,0)",
    previewRange: "F2:L28"
  },
  "CORREL.SPEARMAN": {
    values: [
      ["x", "y"],
      [1, 12],
      [2, 15],
      [3, 16],
      [4, 19],
      [5, 18],
      [6, 24],
      [7, 25],
      [8, 29]
    ],
    formula: "=CORREL.SPEARMAN(A1:A9,B1:B9,0,0.05,1)",
    previewRange: "F2:G11"
  },
  "ANOVA.RM": {
    values: [
      ["baseline", "week 4", "week 8"],
      [10, 12, 13],
      [9, 10, 13],
      [11, 14, 15],
      [12, 13, 17],
      [10, 13, 14],
      [13, 15, 16]
    ],
    formula: "=ANOVA.RM(A1:C7,1,0.05,2)",
    previewRange: "E2:N24"
  },
  "ANCOVA": {
    values: [
      ["group", "score", "age"],
      ["A", 10, 21],
      ["A", 12, 25],
      ["A", 13, 28],
      ["A", 11, 32],
      ["B", 16, 22],
      ["B", 18, 27],
      ["B", 19, 31],
      ["B", 17, 35],
      ["C", 20, 24],
      ["C", 22, 29],
      ["C", 24, 33],
      ["C", 21, 37]
    ],
    formula: "=ANCOVA(A1:A13,B1:B13,C1:C13,2,0.05,1)",
    previewRange: "F2:O30"
  },
  "CONTINGENCY.T": {
    values: [
      ["", "male", "female"],
      ["yes", 30, 20],
      ["no", 10, 40]
    ],
    formula: "=CONTINGENCY.T(A1:C3,1,0.05)",
    previewRange: "E2:J26"
  },
  "CONTINGENCY": {
    values: [
      ["column", "row", "count"],
      ["male", "yes", 30],
      ["male", "no", 10],
      ["female", "yes", 20],
      ["female", "no", 40]
    ],
    formula: "=CONTINGENCY(A1:A5,B1:B5,C1:C5,0.05,1)",
    previewRange: "E2:J26"
  },
  "CORREL.MATRIX": {
    values: [
      ["x", "y", "z"],
      [1, 2, 8],
      [2, 4, 7],
      [3, 6, 5],
      [4, 8, 4],
      [5, 10, 3]
    ],
    formula: "=CORREL.MATRIX(A1:C6,0,3,1,1)",
    previewRange: "E2:J14"
  }
};

function docForParameter(fn, index) {
  return localizedInfo(fn).parameters?.[index] ?? {};
}

function enumForParameter(fn, index) {
  return docForParameter(fn, index).enumValues ?? [];
}

function wizardFieldId(index) {
  return `wizard-param-${index}`;
}

function parameterLabel(fn, parameter, index) {
  return docForParameter(fn, index).sourceName ?? parameter.name;
}

function defaultWizardValue(fn, parameter, index) {
  const configured = docForParameter(fn, index).wizard?.defaultValue;
  if (configured !== undefined) {
    return String(configured);
  }
  const name = parameter.name.toLowerCase();
  if (name === "alpha") {
    return "0.05";
  }
  if (name === "outlierrate") {
    return "0";
  }
  if (name === "quantile") {
    return "0.75";
  }
  if (name === "count") {
    return "100";
  }
  if (name === "minimum") {
    return "1";
  }
  if (name === "maximum") {
    return "6";
  }
  if (name === "standarddeviation") {
    return "1";
  }
  if (name === "mean" || name === "mu0" || name === "pi0") {
    return name === "pi0" ? "0.5" : "0";
  }
  return "";
}

function fieldKindForParameter(fn, parameter, index, enumValues) {
  const configured = docForParameter(fn, index).wizard?.kind;
  if (configured) {
    return configured;
  }
  if (enumValues.length > 0) {
    return "select";
  }
  if (parameter.dimensionality === "matrix") {
    return "range";
  }
  if (parameter.type === "number") {
    return "number";
  }
  return "text";
}

function applyNumberAttributes(control, wizard = {}) {
  control.type = "number";
  control.inputMode = "decimal";
  if (wizard.min !== undefined) {
    control.min = String(wizard.min);
  }
  if (wizard.max !== undefined) {
    control.max = String(wizard.max);
  }
  control.step = wizard.step === undefined ? "any" : String(wizard.step);
}

function createWizardField({ id, label, kind = "text", value = "", enumValues = [], range = false, optional = false, description = "", wizard = {}, datasetVariables = [] }) {
  const row = document.createElement("div");
  row.className = "wizard-field";

  const labelNode = document.createElement("label");
  labelNode.htmlFor = id;
  labelNode.textContent = label;
  row.append(labelNode);

  if (description) {
    const hint = document.createElement("p");
    hint.className = "wizard-field-hint";
    hint.textContent = description;
    row.append(hint);
  }

  const controlWrap = document.createElement("div");
  controlWrap.className = "wizard-control";

  let selectedVariableAddress = "";
  if (range && datasetVariables.length > 0) {
    const variableSelect = document.createElement("select");
    variableSelect.className = "wizard-variable-select";
    const manualOption = document.createElement("option");
    manualOption.value = "";
    manualOption.textContent = t("wizardUseVariable");
    variableSelect.append(manualOption);
    datasetVariables.forEach((variable) => {
      const option = document.createElement("option");
      option.value = variable.address;
      option.textContent = variable.type ? `${variable.name} (${variable.type})` : variable.name;
      if (rangesEquivalent(value, variable.address)) {
        option.selected = true;
        selectedVariableAddress = variable.address;
      }
      variableSelect.append(option);
    });
    variableSelect.addEventListener("change", () => {
      if (variableSelect.value) {
        const field = document.getElementById(id);
        if (field) {
          field.value = variableSelect.value;
        }
      }
    });
    controlWrap.append(variableSelect);
  }

  let control;
  if (kind === "select") {
    control = document.createElement("select");
    enumValues.forEach((item) => {
      const option = document.createElement("option");
      option.value = item.value;
      option.textContent = item.meaning;
      option.selected = value === "" ? Boolean(item.default) : String(item.value) === String(value);
      control.append(option);
    });
  } else {
    control = document.createElement("input");
    control.type = "text";
    control.value = selectedVariableAddress || value;
    if (kind === "number") {
      applyNumberAttributes(control, wizard);
    }
  }
  control.id = id;
  control.dataset.optional = optional ? "true" : "false";
  control.dataset.kind = kind;
  controlWrap.append(control);

  if (range) {
    const selectionButton = document.createElement("button");
    selectionButton.type = "button";
    selectionButton.className = "wizard-small-button";
    selectionButton.textContent = t("wizardUseSelection");
    selectionButton.addEventListener("click", () => useSelectionForField(id));
    controlWrap.append(selectionButton);
  }

  row.append(controlWrap);
  const message = document.createElement("p");
  message.className = "wizard-field-error";
  message.id = `${id}-error`;
  message.hidden = true;
  row.append(message);
  return row;
}

function setWizardFieldError(id, message) {
  const field = document.getElementById(id);
  const row = field?.closest(".wizard-field");
  const error = document.getElementById(`${id}-error`);
  if (!field || !row || !error) {
    return;
  }
  row.classList.toggle("invalid", Boolean(message));
  field.setAttribute("aria-invalid", message ? "true" : "false");
  error.textContent = message ?? "";
  error.hidden = !message;
}

function clearWizardErrors() {
  document.querySelectorAll(".wizard-field.invalid").forEach((row) => row.classList.remove("invalid"));
  document.querySelectorAll(".wizard-field-error").forEach((error) => {
    error.textContent = "";
    error.hidden = true;
  });
  document.querySelectorAll(".wizard-control input, .wizard-control select").forEach((field) => {
    field.setAttribute("aria-invalid", "false");
  });
}

function quoteSheetName(sheetName) {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}

function normalizeRangeAddress(address) {
  const bangIndex = address.lastIndexOf("!");
  if (bangIndex === -1) {
    return address.replace(/\$/g, "");
  }
  const sheetName = address.slice(0, bangIndex).replace(/^'|'$/g, "").replace(/''/g, "'");
  const range = address.slice(bangIndex + 1).replace(/\$/g, "");
  return `${quoteSheetName(sheetName)}!${range}`;
}

function parseQualifiedAddress(address) {
  const normalized = normalizeRangeAddress(address);
  const bangIndex = normalized.lastIndexOf("!");
  if (bangIndex === -1) {
    return { sheetName: null, range: normalized };
  }
  return {
    sheetName: normalized.slice(0, bangIndex).replace(/^'|'$/g, "").replace(/''/g, "'"),
    range: normalized.slice(bangIndex + 1)
  };
}

function rangesEquivalent(left, right) {
  if (!left || !right) {
    return false;
  }
  const normalizedLeft = normalizeRangeAddress(String(left)).toUpperCase();
  const normalizedRight = normalizeRangeAddress(String(right)).toUpperCase();
  if (normalizedLeft === normalizedRight) {
    return true;
  }
  const leftColumn = singleColumnKey(normalizedLeft);
  const rightColumn = singleColumnKey(normalizedRight);
  if (!leftColumn || !rightColumn) {
    return false;
  }
  return leftColumn.column === rightColumn.column
    && (!leftColumn.sheet || !rightColumn.sheet || leftColumn.sheet === rightColumn.sheet);
}

function singleColumnKey(address) {
  const parsed = parseQualifiedAddress(address);
  const match = parsed.range.match(/^([A-Z]+)(?::\1|\d+:\1\d+)$/i);
  return match ? { sheet: parsed.sheetName?.toUpperCase() ?? "", column: match[1].toUpperCase() } : null;
}

async function getSelectedAddress() {
  return Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    selected.load(["address", "rowCount", "columnCount"]);
    await context.sync();

    const isWholeColumn = selected.rowCount >= excelMaxRows;
    const isWholeRow = selected.columnCount >= excelMaxColumns;
    if (isWholeColumn || isWholeRow) {
      const usedRange = selected.worksheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) {
        return normalizeRangeAddress(selected.address);
      }
      const intersection = selected.getIntersectionOrNullObject(usedRange);
      intersection.load(["address", "isNullObject"]);
      await context.sync();
      if (!intersection.isNullObject) {
        return normalizeRangeAddress(intersection.address);
      }
    }

    return normalizeRangeAddress(selected.address);
  });
}

function splitFormulaArguments(argumentText) {
  const args = [];
  let current = "";
  let depth = 0;
  let inDoubleQuote = false;
  let inSingleQuote = false;

  for (let index = 0; index < argumentText.length; index += 1) {
    const char = argumentText[index];
    const next = argumentText[index + 1];
    if (char === '"' && !inSingleQuote) {
      inDoubleQuote = !inDoubleQuote;
      current += char;
      continue;
    }
    if (char === "'" && !inDoubleQuote) {
      current += char;
      if (next === "'") {
        current += next;
        index += 1;
      } else {
        inSingleQuote = !inSingleQuote;
      }
      continue;
    }
    if (!inDoubleQuote && !inSingleQuote) {
      if (char === "(") {
        depth += 1;
      } else if (char === ")" && depth > 0) {
        depth -= 1;
      } else if ((char === "," || char === ";") && depth === 0) {
        const previous = argumentText[index - 1] ?? "";
        const following = argumentText[index + 1] ?? "";
        const decimalComma = char === "," && /\d/.test(previous) && /\d/.test(following);
        if (!decimalComma) {
          args.push(current.trim());
          current = "";
          continue;
        }
      }
    }
    current += char;
  }

  if (current.trim() || argumentText.length > 0) {
    args.push(current.trim());
  }
  return args;
}

function parseFormulaCall(formula, functionName) {
  if (!formula) {
    return null;
  }
  const trimmed = formula.trim().replace(/^=/, "");
  const upper = trimmed.toUpperCase();
  const searchNames = [functionName.toUpperCase(), `EVALYTICS.${functionName.toUpperCase()}`];
  const matchedName = searchNames.find((name) => upper.startsWith(`${name}(`));
  if (!matchedName || !trimmed.endsWith(")")) {
    return null;
  }
  return splitFormulaArguments(trimmed.slice(matchedName.length + 1, -1));
}

async function readSelectedFormulaContext(functionName) {
  if (!window.Excel) {
    return null;
  }
  return Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    selected.load(["address", "formulas", "formulasLocal"]);
    await context.sync();
    const formula = selected.formulas?.[0]?.[0];
    const formulaLocal = selected.formulasLocal?.[0]?.[0];
    const args = parseFormulaCall(formulaLocal, functionName) ?? parseFormulaCall(formula, functionName);
    return {
      address: normalizeRangeAddress(selected.address),
      args
    };
  });
}

async function useSelectionForField(id) {
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    return;
  }
  try {
    const address = await getSelectedAddress();
    const field = document.getElementById(id);
    if (field) {
      field.value = address;
    }
    debugLog("Wizard selection captured", { field: id, address });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Wizard selection failed", { field: id, message: error instanceof Error ? error.message : String(error) });
  }
}

function guessVariableType(values) {
  const filled = values.filter((value) => value !== null && value !== undefined && String(value).trim() !== "");
  if (filled.length === 0) {
    return "";
  }
  const numeric = filled.filter((value) => tryParseDatasetNumber(value)).length;
  if (numeric === filled.length) {
    return language === "cs" ? "číslo" : "number";
  }
  const unique = new Set(filled.map((value) => String(value).trim())).size;
  if (unique <= Math.max(8, Math.ceil(filled.length / 2))) {
    return language === "cs" ? "kategorie" : "category";
  }
  return language === "cs" ? "text" : "text";
}

function tryParseDatasetNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return true;
  }
  if (typeof value !== "string") {
    return false;
  }
  return Number.isFinite(Number(value.trim().replace(",", ".")));
}

async function loadDatasetVariables(sheetName = null) {
  if (!window.Excel) {
    return [];
  }
  return Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load(["address", "columnCount", "rowCount", "values", "isNullObject"]);
    await context.sync();
    if (usedRange.isNullObject || usedRange.rowCount < 2 || usedRange.columnCount < 1) {
      return [];
    }

    const columns = [];
    for (let index = 0; index < usedRange.columnCount; index += 1) {
      const column = usedRange.getColumn(index);
      column.load("address");
      columns.push(column);
    }
    await context.sync();

    const firstRow = usedRange.values[0] ?? [];
    return columns.map((column, index) => {
      const rawName = firstRow[index];
      const name = rawName === null || rawName === undefined || String(rawName).trim() === ""
        ? `${language === "cs" ? "Proměnná" : "Variable"} ${index + 1}`
        : String(rawName).trim();
      const values = usedRange.values.slice(1).map((row) => row[index]);
      return {
        name,
        address: normalizeRangeAddress(column.address),
        type: guessVariableType(values)
      };
    });
  }).catch((error) => {
    debugLog("Wizard dataset load failed", { message: error instanceof Error ? error.message : String(error) });
    return [];
  });
}

async function openFunctionWizard(fn) {
  const panel = document.getElementById("wizard-panel");
  panel.hidden = false;
  panel.replaceChildren();

  const header = document.createElement("div");
  header.className = "wizard-header";
  const title = document.createElement("h2");
  title.textContent = `${t("wizardTitle")}: ${fn.name}`;
  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "wizard-close-button";
  closeButton.textContent = t("wizardClose");
  closeButton.addEventListener("click", () => {
    panel.hidden = true;
    panel.replaceChildren();
  });
  header.append(title, closeButton);
  panel.append(header);

  const selectedFormula = await readSelectedFormulaContext(fn.name).catch((error) => {
    debugLog("Wizard formula read failed", { functionName: fn.name, message: error instanceof Error ? error.message : String(error) });
    return null;
  });
  const selectedArgs = selectedFormula?.args ?? null;
  const selectedSheetName = selectedFormula?.address ? parseQualifiedAddress(selectedFormula.address).sheetName : null;
  const datasetVariables = datasetWizardFunctions.has(fn.name)
    ? await loadDatasetVariables(selectedSheetName)
    : [];

  const form = document.createElement("div");
  form.className = "wizard-form";
  if (datasetVariables.length > 0) {
    const datasetInfo = document.createElement("p");
    datasetInfo.className = "wizard-dataset-info";
    datasetInfo.textContent = t("wizardDatasetLoaded");
    form.append(datasetInfo);
  }
  form.append(createWizardField({ id: "wizard-target", label: t("wizardTarget"), value: selectedFormula?.address ?? "", range: true, optional: true }));

  fn.parameters.forEach((parameter, index) => {
    const enumValues = enumForParameter(fn, index);
    const kind = fieldKindForParameter(fn, parameter, index, enumValues);
    const selectedValue = selectedArgs?.[index];
    const value = selectedValue === undefined
      ? defaultWizardValue(fn, parameter, index)
      : selectedValue;
    form.append(createWizardField({
      id: wizardFieldId(index),
      label: parameterLabel(fn, parameter, index),
      value,
      kind,
      range: kind === "range",
      enumValues,
      optional: Boolean(parameter.optional),
      description: docForParameter(fn, index).description ?? "",
      wizard: docForParameter(fn, index).wizard ?? {},
      datasetVariables: kind === "range" ? datasetVariables : []
    }));
  });

  const actions = document.createElement("div");
  actions.className = "wizard-actions";
  const insertButton = document.createElement("button");
  insertButton.type = "button";
  insertButton.className = "copy-button";
  insertButton.textContent = t("wizardInsert");
  insertButton.addEventListener("click", () => insertFunctionFormulaFromWizard(fn));
  actions.append(insertButton);

  panel.append(form, actions);
  debugLog("Wizard opened", { functionName: fn.name, target: selectedFormula?.address ?? "", populatedFromFormula: Boolean(selectedArgs) });
}

function wizardValue(id) {
  return document.getElementById(id)?.value?.trim() ?? "";
}

function normalizeFormulaNumber(value) {
  return value.replace(",", ".");
}

function localizeFormulaNumber(value) {
  return language === "cs" ? String(value).replace(".", ",") : String(value);
}

function parseWizardNumber(value) {
  const normalized = normalizeFormulaNumber(String(value).trim());
  if (!normalized) {
    return Number.NaN;
  }
  return Number(normalized);
}

function rangeFromAddress(context, address) {
  const parsed = parseQualifiedAddress(address);
  const worksheet = parsed.sheetName
    ? context.workbook.worksheets.getItem(parsed.sheetName)
    : context.workbook.worksheets.getActiveWorksheet();
  return worksheet.getRange(parsed.range);
}

async function validateExcelRange(address, { singleCell = false } = {}) {
  return Excel.run(async (context) => {
    const range = rangeFromAddress(context, address);
    range.load(["address", "rowCount", "columnCount"]);
    await context.sync();
    if (singleCell && (range.rowCount !== 1 || range.columnCount !== 1)) {
      return { valid: false, rowCount: range.rowCount, columnCount: range.columnCount };
    }
    return {
      valid: true,
      address: normalizeRangeAddress(range.address),
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      cellCount: range.rowCount * range.columnCount,
      tooLarge: !singleCell && (range.rowCount >= excelMaxRows || range.columnCount >= excelMaxColumns)
    };
  }).catch((error) => {
    debugLog("Wizard range validation failed", { address, message: error instanceof Error ? error.message : String(error) });
    return { valid: false };
  });
}

function isBoundedProbabilityParameter(parameter) {
  return ["alpha", "outlierRate", "quantile", "pi0", "pMinimum"].includes(parameter.name);
}

function isPositiveNumberParameter(parameter) {
  return ["standardDeviation", "count"].includes(parameter.name);
}

function wizardValidationFor(fn, parameter, index) {
  return docForParameter(fn, index).wizard?.validation ?? "";
}

function formatWizardScalar(value, parameter, fn, index) {
  const trimmed = String(value ?? "").trim();
  if (!trimmed && parameter.optional) {
    const fallback = defaultWizardValue(fn, parameter, index);
    return fallback ? formatWizardScalar(fallback, { ...parameter, optional: false }, fn, index) : "";
  }
  if (parameter.type === "number") {
    return String(parseWizardNumber(trimmed));
  }
  if (parameter.type === "string") {
    const unquoted = trimmed.replace(/^"(.*)"$/, "$1");
    return `"${unquoted.replace(/"/g, '""')}"`;
  }
  return trimmed;
}

function inferredEqualCellGroups(fn) {
  const names = fn.parameters.map((parameter) => parameter.name);
  const matrixIndexes = fn.parameters
    .map((parameter, index) => ({ parameter, index }))
    .filter((item) => item.parameter.dimensionality === "matrix")
    .map((item) => item.index);

  if (["categories", "values"].every((name) => names.includes(name))) {
    return [[names.indexOf("categories"), names.indexOf("values")]];
  }
  if (["groups", "values", "covariates"].every((name) => names.includes(name))) {
    return [[names.indexOf("groups"), names.indexOf("values"), names.indexOf("covariates")]];
  }
  if (["xValues", "yValues"].every((name) => names.includes(name))) {
    return [[names.indexOf("xValues"), names.indexOf("yValues")]];
  }
  if (["rows", "columns", "values"].every((name) => names.includes(name))) {
    return [[names.indexOf("rows"), names.indexOf("columns"), names.indexOf("values")]];
  }
  if (["columnCategories", "rowCategories", "counts"].every((name) => names.includes(name))) {
    return [[names.indexOf("columnCategories"), names.indexOf("rowCategories"), names.indexOf("counts")]];
  }
  if (fn.name === "CHISQ.GOF") {
    return [[0, 1]];
  }
  if (matrixIndexes.length === 2 && !["ANOVA.RM", "CONTINGENCY.T", "CORREL.MATRIX"].includes(fn.name)) {
    return [matrixIndexes];
  }
  return [];
}

async function validateFunctionWizardInputs(fn) {
  clearWizardErrors();
  const errors = [];
  const values = fn.parameters.map((_, index) => wizardValue(wizardFieldId(index)));
  const target = wizardValue("wizard-target");
  const rangeChecks = new Map();

  fn.parameters.forEach((parameter, index) => {
    const value = values[index];
    const fieldId = wizardFieldId(index);
    if (!value && !parameter.optional) {
      errors.push([fieldId, parameter.dimensionality === "matrix" ? t("wizardMissingRange") : t("wizardValueMissing")]);
      return;
    }
    if (!value && parameter.optional) {
      return;
    }

    if (parameter.type === "number") {
      const number = parseWizardNumber(value);
      const validation = wizardValidationFor(fn, parameter, index);
      if (!Number.isFinite(number)) {
        errors.push([fieldId, t("wizardNumberInvalid")]);
      } else if ((validation === "probability" || isBoundedProbabilityParameter(parameter)) && (number < 0 || number > 1)) {
        errors.push([fieldId, parameter.name === "alpha" ? t("wizardAlphaInvalid") : t("wizardProbabilityInvalid")]);
      } else if ((validation === "positive" || isPositiveNumberParameter(parameter)) && number <= 0) {
        errors.push([fieldId, t("wizardPositiveInvalid")]);
      }
    }
  });

  if (errors.length === 0) {
    const targetCheck = target ? await validateExcelRange(target, { singleCell: true }) : { valid: true };
    if (!targetCheck.valid) {
      errors.push(["wizard-target", t("wizardTargetInvalid")]);
    }

    for (const [index, parameter] of fn.parameters.entries()) {
      const value = values[index];
      if (parameter.dimensionality !== "matrix" || !value) {
        continue;
      }
      const check = await validateExcelRange(value);
      rangeChecks.set(index, check);
      if (!check.valid) {
        errors.push([wizardFieldId(index), t("wizardRangeInvalid")]);
      } else if (check.tooLarge) {
        errors.push([wizardFieldId(index), t("wizardRangeTooLarge")]);
      }
    }

    inferredEqualCellGroups(fn).forEach((group) => {
      const checks = group.map((index) => rangeChecks.get(index)).filter(Boolean);
      if (checks.length < 2 || checks.some((check) => !check.valid)) {
        return;
      }
      const cellCount = checks[0].cellCount;
      if (checks.some((check) => check.cellCount !== cellCount)) {
        group.forEach((index) => errors.push([wizardFieldId(index), t("wizardRangeShapeInvalid")]));
      }
    });
  }

  const args = fn.parameters.map((parameter, index) => {
    const value = values[index];
    return parameter.dimensionality === "matrix" ? value : formatWizardScalar(value, parameter, fn, index);
  });
  while (args.length > 0 && !args[args.length - 1]) {
    args.pop();
  }

  errors.forEach(([id, message]) => setWizardFieldError(id, message));
  return {
    valid: errors.length === 0,
    args,
    target,
    message: errors[0]?.[1] ?? ""
  };
}

async function insertFunctionFormulaFromWizard(fn) {
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    return;
  }
  if (wizardInsertInProgress) {
    return;
  }

  wizardInsertInProgress = true;
  const insertButton = document.querySelector(".wizard-actions button");
  try {
    if (insertButton) {
      insertButton.disabled = true;
    }

    const validation = await validateFunctionWizardInputs(fn);
    if (!validation.valid) {
      setStatus(validation.message || t("wizardRangeInvalid"));
      return;
    }
    const formula = `=${fn.name}(${validation.args.join(",")})`;

    try {
      await Excel.run(async (context) => {
        let destination;
        if (validation.target) {
          destination = rangeFromAddress(context, validation.target);
        } else {
          destination = context.workbook.getSelectedRange();
        }
        destination.formulas = [[formula]];
        await context.sync();
      });
      setStatus(t("wizardReady"));
      debugLog("Wizard formula inserted", { functionName: fn.name, formula, target: validation.target });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
      debugLog("Wizard insert failed", { functionName: fn.name, formula, message: error instanceof Error ? error.message : String(error) });
    }
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Wizard validation failed", { functionName: fn.name, message: error instanceof Error ? error.message : String(error) });
  } finally {
    if (insertButton) {
      insertButton.disabled = false;
    }
    wizardInsertInProgress = false;
  }
}

async function runTutorial(functionName) {
  const definition = tutorialDefinitions[functionName];
  if (definition?.customRun) {
    await definition.customRun();
    return;
  }

  debugLog("Tutorial clicked", { functionName, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing", { functionName });
    return;
  }

  let sheetName = "";
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");

      sheet.getRange("A1").values = [[definition.header]];
      sheet.getRange("A1").format.font.bold = true;
      sheet.getRange("A2").formulas = [[definition.formula]];
      sheet.getRange("A:A").format.autofitColumns();
      await context.sync();
      sheetName = sheet.name;
      debugLog("Tutorial formulas written", {
        functionName,
        sheetName,
        formula: definition.formula
      });

      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { functionName, type: "full" });
    });

    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const preview = sheet.getRange(definition.range);
        preview.load("text");
        await context.sync();
        return {
          attempt,
          formula: definition.formula,
          text: preview.text
        };
      });

      const busy = state.text.flat().some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", { functionName, ...state });
      }
      if (!busy) {
        break;
      }
      await sleep(500);
    }

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed", { functionName, sheetName });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Tutorial failed", { functionName, message: error instanceof Error ? error.message : String(error) });
  }
}

async function runMatrixTutorial(functionName) {
  const definition = matrixTutorialDefinitions[functionName];
  let sheetName = "";

  debugLog("Tutorial clicked", { functionName, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing", { functionName });
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");

      sheet.getRangeByIndexes(0, 0, definition.values.length, definition.values[0].length).values = definition.values;
      sheet.getRangeByIndexes(0, 0, 1, definition.values[0].length).format.font.bold = true;
      if (definition.options) {
        sheet.getRangeByIndexes(0, 16, definition.options.length, 2).values = definition.options;
        sheet.getRange("Q1:R1").format.font.bold = true;
        if (functionName === "ANCOVA") {
          sheet.getRange("R2").dataValidation.rule = { list: { inCellDropDown: true, source: "0,1,2,3,4" } };
          sheet.getRange("R4").dataValidation.rule = { list: { inCellDropDown: true, source: "0,1,2" } };
        }
      }
      sheet.getRange("F2").formulas = [[definition.formula]];
      sheet.getRange("A:R").format.autofitColumns();
      await context.sync();

      sheetName = sheet.name;
      debugLog("Tutorial formulas written", { functionName, sheetName, formula: definition.formula, options: definition.options ?? null });
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { functionName, type: "full" });
    });

    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const preview = sheet.getRange(definition.previewRange);
        preview.load("text");
        await context.sync();
        return {
          attempt,
          formula: definition.formula,
          text: preview.text
        };
      });

      const busy = state.text.flat().some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", { functionName, ...state });
      }
      if (!busy) {
        break;
      }
      await sleep(500);
    }

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed", { functionName, sheetName });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Tutorial failed", { functionName, message: error instanceof Error ? error.message : String(error) });
  }
}

async function runParseNumberTutorial() {
  const functionName = "PARSE.NUMBER";
  const rawValues = [
    "1 234,56",
    "1,234.56",
    "CZK 1 234,56",
    "1.234,56 Kč",
    " - 42 ",
    "−12,5",
    "(987,65)",
    "2 500",
    "3.14",
    "4,5",
    "n/a"
  ];
  const formulas = rawValues.map((_, index) => [`=PARSE.NUMBER(A${index + 2})`]);
  let sheetName = "";

  debugLog("Tutorial clicked", { functionName, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing", { functionName });
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");

      sheet.getRange("A1:B1").values = [["raw", "value"]];
      sheet.getRange("A1:B1").format.font.bold = true;
      sheet.getRange(`A2:A${rawValues.length + 1}`).values = rawValues.map((value) => [value]);
      sheet.getRange(`B2:B${rawValues.length + 1}`).formulas = formulas;
      sheet.getRange("A:B").format.autofitColumns();
      await context.sync();

      sheetName = sheet.name;
      debugLog("Tutorial formulas written", { functionName, sheetName });
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { functionName, type: "full" });
    });

    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const preview = sheet.getRange(`B2:B${rawValues.length + 1}`);
        preview.load("text");
        await context.sync();
        return {
          attempt,
          text: preview.text
        };
      });

      const busy = state.text.flat().some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", { functionName, ...state });
      }
      if (!busy) {
        break;
      }
      await sleep(500);
    }

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed", { functionName, sheetName });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Tutorial failed", { functionName, message: error instanceof Error ? error.message : String(error) });
  }
}

async function runStackGTutorial() {
  const functionName = "STACK";
  const values = [
    ["control", "treatment", "placebo"],
    [10, 16, 9],
    [12, 18, 10],
    [13, 19, 11],
    [11, 17, 10]
  ];
  let sheetName = "";

  debugLog("Tutorial clicked", { functionName, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing", { functionName });
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");

      sheet.getRange("A1:C5").values = values;
      sheet.getRange("A1:C1").format.font.bold = true;
      sheet.getRange("E2").formulas = [["=STACK(A1:C5,1)"]];
      sheet.getRange("A:F").format.autofitColumns();
      await context.sync();

      sheetName = sheet.name;
      debugLog("Tutorial formulas written", { functionName, sheetName, formula: "=STACK(A1:C5,1)" });
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { functionName, type: "full" });
    });

    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const preview = sheet.getRange("E2:F14");
        preview.load("text");
        await context.sync();
        return { attempt, formula: "=STACK(A1:C5,1)", text: preview.text };
      });

      const busy = state.text.flat().some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", { functionName, ...state });
      }
      if (!busy) {
        break;
      }
      await sleep(500);
    }

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed", { functionName, sheetName });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Tutorial failed", { functionName, message: error instanceof Error ? error.message : String(error) });
  }
}

async function runUnstackGTutorial() {
  const functionName = "UNSTACK";
  const values = [
    ["group", "value"],
    ["control", 10],
    ["control", 12],
    ["control", 13],
    ["control", 11],
    ["treatment", 16],
    ["treatment", 18],
    ["treatment", 19],
    ["treatment", 17],
    ["placebo", 9],
    ["placebo", 10],
    ["placebo", 11],
    ["placebo", 10]
  ];
  let sheetName = "";

  debugLog("Tutorial clicked", { functionName, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing", { functionName });
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");

      sheet.getRange("A1:B13").values = values;
      sheet.getRange("A1:B1").format.font.bold = true;
      sheet.getRange("D2").formulas = [["=UNSTACK(A1:A13,B1:B13,1)"]];
      sheet.getRange("A:F").format.autofitColumns();
      await context.sync();

      sheetName = sheet.name;
      debugLog("Tutorial formulas written", { functionName, sheetName, formula: "=UNSTACK(A1:A13,B1:B13,1)" });
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { functionName, type: "full" });
    });

    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const preview = sheet.getRange("D2:F5");
        preview.load("text");
        await context.sync();
        return { attempt, formula: "=UNSTACK(A1:A13,B1:B13,1)", text: preview.text };
      });

      const busy = state.text.flat().some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", { functionName, ...state });
      }
      if (!busy) {
        break;
      }
      await sleep(500);
    }

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed", { functionName, sheetName });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Tutorial failed", { functionName, message: error instanceof Error ? error.message : String(error) });
  }
}

async function runDescriptiveTutorial(functionName) {
  const definition = descriptiveTutorialDefinitions[functionName];
  let sheetName = "";

  debugLog("Tutorial clicked", { functionName, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing", { functionName });
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");

      sheet.getRangeByIndexes(0, 0, 1, definition.headers.length).values = [definition.headers];
      sheet.getRangeByIndexes(0, 0, 1, definition.headers.length).format.font.bold = true;
      sheet.getRangeByIndexes(1, 0, definition.rows.length, definition.rows[0].length).values = definition.rows;
      sheet.getRange(definition.formulaCell).formulas = [[definition.formula]];
      sheet.getRange("A:D").format.autofitColumns();
      await context.sync();

      sheetName = sheet.name;
      debugLog("Tutorial formulas written", { functionName, sheetName, formula: definition.formula });
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { functionName, type: "full" });
    });

    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const preview = sheet.getRange(definition.previewRange);
        preview.load("text");
        await context.sync();
        return {
          attempt,
          formula: definition.formula,
          text: preview.text
        };
      });

      const busy = state.text.flat().some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", { functionName, ...state });
      }
      if (!busy) {
        break;
      }
      await sleep(500);
    }

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed", { functionName, sheetName });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Tutorial failed", { functionName, message: error instanceof Error ? error.message : String(error) });
  }
}

async function runPivotTutorial(functionName) {
  const formula = pivotTutorialFormula(functionName);
  let sheetName = "";

  debugLog("Tutorial clicked", { functionName, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing", { functionName });
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");

      sheet.getRange("A1:C1").values = [["region", "channel", "value"]];
      pivotTutorialRowBlocks.forEach((block) => {
        sheet.getRange(block.cell).formulas = [[block.formula]];
      });
      pivotTutorialColumnBlocks.forEach((block) => {
        sheet.getRange(block.cell).formulas = [[block.formula]];
      });
      sheet.getRange("C2").formulas = [[pivotTutorialValueFormula]];
      sheet.getRange("A1:C1").format.font.bold = true;
      sheet.getRange("E2").formulas = [[formula]];
      sheet.getRange("A:H").format.autofitColumns();
      await context.sync();

      sheetName = sheet.name;
      debugLog("Tutorial formulas written", {
        functionName,
        sheetName,
        formula,
        rowFormulas: pivotTutorialRowBlocks,
        columnFormulas: pivotTutorialColumnBlocks,
        valueFormula: pivotTutorialValueFormula
      });
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { functionName, type: "full" });
    });

    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const preview = sheet.getRange("E2:L9");
        preview.load("text");
        await context.sync();
        return {
          attempt,
          formula,
          text: preview.text
        };
      });

      const busy = state.text.flat().some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", { functionName, ...state });
      }
      if (!busy) {
        break;
      }
      await sleep(500);
    }

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed", { functionName, sheetName });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    debugLog("Tutorial failed", { functionName, message: error instanceof Error ? error.message : String(error) });
  }
}

function renderOutput(container, fn) {
  container.replaceChildren();

  const title = document.createElement("h3");
  title.textContent = t("output");
  container.append(title);

  const output = outputDocs(fn);
  if (output.length === 0) {
    const empty = document.createElement("p");
    empty.textContent = t("noOutput");
    container.append(empty);
    return;
  }

  const list = document.createElement("dl");
  list.className = "output-list";
  let currentSection = "";
  output.forEach((item) => {
    if (item.section && item.section !== currentSection) {
      currentSection = item.section;
      const section = document.createElement("div");
      section.className = "output-section-title";
      section.textContent = item.section;
      list.append(section);
    }
    const term = document.createElement("dt");
    term.textContent = item.name;
    const definition = document.createElement("dd");
    definition.textContent = item.description;
    list.append(term, definition);
  });
  container.append(list);
}

function render() {
  translateStaticText();

  const list = document.getElementById("function-list");
  const template = document.getElementById("function-card-template");
  list.replaceChildren();

  const visible = functions.filter(matchesSearch);
  setStatus(visible.length === 0 ? t("empty") : `${visible.length} / ${functions.length}`);

  const groups = new Map();
  visible.forEach((fn) => {
    const category = fn.pending ? t("pending") : displayCategory(fn);
    if (!groups.has(category)) {
      groups.set(category, []);
    }

    groups.get(category).push(fn);
  });

  [...groups.entries()]
    .sort(([left], [right]) => categoryRank(left) - categoryRank(right) || left.localeCompare(right))
    .forEach(([categoryName, groupFunctions]) => {
      const section = document.createElement("section");
      section.className = "function-section collapsed";

      const heading = document.createElement("button");
      heading.className = "section-heading";
      heading.type = "button";
      heading.textContent = `${categoryName} (${groupFunctions.length})`;
      section.append(heading);

      const sectionBody = document.createElement("div");
      sectionBody.className = "section-body";
      sectionBody.hidden = true;
      section.append(sectionBody);

      heading.addEventListener("click", () => {
        sectionBody.hidden = !sectionBody.hidden;
        section.classList.toggle("collapsed", sectionBody.hidden);
      });

      groupFunctions
        .slice()
        .sort((a, b) => a.name.localeCompare(b.name))
        .forEach((fn) => {
    const fragment = template.content.cloneNode(true);
    const card = fragment.querySelector(".function-card");
    const header = fragment.querySelector(".function-header");
    const name = fragment.querySelector(".function-name");
    const description = fragment.querySelector(".description");
    const formulaLabel = fragment.querySelector(".formula-label");
    const signature = fragment.querySelector(".signature");
    const parameters = fragment.querySelector(".parameters");
    const outputs = fragment.querySelector(".outputs");
    const actions = fragment.querySelector(".actions");

    card.classList.remove("open");
    name.textContent = fn.name;
    if (localizedInfo(fn).auth?.required && !isAuthActive()) {
      const badge = document.createElement("span");
      badge.className = "auth-required-badge";
      badge.textContent = t("authRequired");
      name.append(badge);
    }
    description.textContent = displaySummary(fn);
    formulaLabel.textContent = t("formula");
    signature.textContent = fn.signature;
    renderParameters(parameters, fn);
    renderOutput(outputs, fn);
    addTutorialButton(actions, fn);
    addWizardButton(actions, fn);

    actions.addEventListener("click", (event) => {
      event.stopPropagation();
    });

    header.addEventListener("click", () => {
      card.classList.toggle("open");
    });

          sectionBody.append(fragment);
        });

      list.append(section);
    });
}

async function runWelchTutorialWithDiagnostics() {
  debugLog("Tutorial clicked", { hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  if (!window.Excel) {
    setStatus(t("tutorialUnavailable"));
    debugLog("Tutorial unavailable: Excel API missing");
    return;
  }

  let sheetName = "";
  const formula = "=WELCH.TEST.2S(A1:A101,B1:B101,1,0.05,0)";
  const maleCategoryFormula = '=FILL("male",50)';
  const femaleCategoryFormula = '=FILL("female",50)';
  const maleValuesFormula = "=GENERATE.NORM.ARRAY(50,178,8)";
  const femaleValuesFormula = "=GENERATE.NORM.ARRAY(50,168,7)";
  const versionFormula = "=VERSION()";

  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.add();
      sheet.activate();
      sheet.load("name");
      await context.sync();
      sheetName = sheet.name;
      debugLog("Tutorial sheet created", { sheetName });

      sheet.getRange("A1:B1").values = [["group", "value"]];
      sheet.getRange("A2").formulas = [[maleCategoryFormula]];
      sheet.getRange("A52").formulas = [[femaleCategoryFormula]];
      sheet.getRange("B2").formulas = [[maleValuesFormula]];
      sheet.getRange("B52").formulas = [[femaleValuesFormula]];
      sheet.getRange("D1").values = [["Welch two-sample t-test"]];
      sheet.getRange("A:D").format.autofitColumns();
      sheet.getRange("1:1").format.font.bold = true;
      await context.sync();
      debugLog("Tutorial data formulas written", {
        dataRange: "A1:B101",
        maleCategoryCell: "A2",
        maleCategoryFormula,
        femaleCategoryCell: "A52",
        femaleCategoryFormula,
        maleValuesCell: "B2",
        maleValuesFormula,
        femaleValuesCell: "B52",
        femaleValuesFormula,
        hasHeader: 1
      });
    });

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.getRange("D3").formulas = [[formula]];
      sheet.getRange("L1").formulas = [[versionFormula]];
      await context.sync();
      debugLog("Tutorial formula inserted", { cell: "D3", formula, versionCell: "L1", versionFormula });

      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
      debugLog("Tutorial recalculation requested", { type: "full" });
    });

    let calculationSettled = false;
    for (let attempt = 1; attempt <= 30; attempt += 1) {
      const state = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const version = sheet.getRange("L1");
        const result = sheet.getRange("D3");
        const firstCategory = sheet.getRange("A2");
        const firstValue = sheet.getRange("B2");
        const lastCategory = sheet.getRange("A101");
        const lastValue = sheet.getRange("B101");
        version.load("text");
        result.load("text");
        firstCategory.load("text");
        firstValue.load("text");
        lastCategory.load("text");
        lastValue.load("text");
        await context.sync();
        return {
          attempt,
          version: version.text?.[0]?.[0],
          result: result.text?.[0]?.[0],
          firstCategory: firstCategory.text?.[0]?.[0],
          firstValue: firstValue.text?.[0]?.[0],
          lastCategory: lastCategory.text?.[0]?.[0],
          lastValue: lastValue.text?.[0]?.[0]
        };
      });

      const busy = [state.version, state.result, state.firstCategory, state.firstValue, state.lastCategory, state.lastValue].some(textLooksBusy);
      if (attempt === 1 || attempt % 5 === 0 || !busy) {
        debugLog("Tutorial calculation poll", state);
      }

      if (!busy) {
        calculationSettled = true;
        break;
      }

      await sleep(500);
    }

    if (!calculationSettled) {
      debugLog("Tutorial calculation poll timeout", { timeoutMs: 15000 });
    }

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const result = sheet.getRange("D3:J13");
      const firstRows = sheet.getRange("A1:B6");
      const lastRows = sheet.getRange("A97:B101");
      const version = sheet.getRange("L1");
      result.load(["values", "text", "formulas", "formulasLocal"]);
      firstRows.load(["values", "text", "formulas", "formulasLocal"]);
      lastRows.load(["values", "text", "formulas", "formulasLocal"]);
      version.load(["values", "text", "formulas", "formulasLocal"]);
      await context.sync();
      debugLog("Tutorial custom-functions runtime version", {
        formula: version.formulas?.[0]?.[0],
        formulaLocal: version.formulasLocal?.[0]?.[0],
        value: version.values?.[0]?.[0],
        text: version.text?.[0]?.[0]
      });
      debugLog("Tutorial result read-back", {
        formula: result.formulas?.[0]?.[0],
        formulaLocal: result.formulasLocal?.[0]?.[0],
        topLeftValue: result.values?.[0]?.[0],
        topLeftText: result.text?.[0]?.[0],
        previewText: result.text?.slice(0, 4)
      });
      debugLog("Tutorial data read-back", {
        firstRows: firstRows.text,
        lastRows: lastRows.text
      });
    });

    setStatus(t("tutorialReady"));
    debugLog("Tutorial completed");
  } catch (error) {
    const message = error?.message ?? String(error);
    setStatus(message);
    debugLog("Tutorial failed", {
      message,
      code: error?.code,
      debugInfo: error?.debugInfo
    });
  }
}

async function fetchJson(path) {
  const separator = path.includes("?") ? "&" : "?";
  const response = await fetch(`${path}${separator}v=${encodeURIComponent(buildStamp)}`, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Unable to load ${path}: ${response.status}`);
  }

  return response.json();
}

async function loadFunctions() {
  setStatus(t("loading"));
  const [metadata, docs] = await Promise.all([
    fetchJson("./functions.json"),
    fetchJson("./function-docs.json")
  ]);

  localizedDocs = docs;
  functions = metadata.functions
    .filter((fn) => !hiddenDocumentationFunctions.has(fn.name))
    .map(normalizeFunction);
  render();
}

function initialize() {
  const stamp = document.getElementById("build-stamp");
  if (stamp) {
    stamp.textContent = `Build ${buildVersion}`;
  }

  document.querySelectorAll("[data-lang]").forEach((button) => {
    button.addEventListener("click", () => {
      language = button.dataset.lang;
      render();
      persistLanguage(language).catch((error) => {
        debugLog("Language recalc failed", { message: error instanceof Error ? error.message : String(error) });
      });
    });
  });

  document.querySelectorAll("[data-tab-target]").forEach((button) => {
    button.addEventListener("click", () => {
      switchTab(button.dataset.tabTarget);
    });
  });

  document.getElementById("search").addEventListener("input", (event) => {
    search = event.target.value.trim();
    render();
  });

  document.getElementById("copy-debug-log")?.addEventListener("click", () => {
    copyDebugLog().catch((error) => {
      setStatus(error?.message ?? String(error));
    });
  });

  document.getElementById("auth-dev-login")?.addEventListener("click", () => {
    simulateAuthLogin().catch((error) => {
      setStatus(error?.message ?? String(error));
      debugLog("Auth dev login failed", { message: error?.message ?? String(error) });
    });
  });

  document.getElementById("auth-logout")?.addEventListener("click", () => {
    logoutAuth().catch((error) => {
      setStatus(error?.message ?? String(error));
      debugLog("Auth logout failed", { message: error?.message ?? String(error) });
    });
  });

  translateStaticText();
  updateAuthStatus();
  switchTab(activeTab);
  debugLog("Panel initialized", { buildStamp, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  loadFunctions().catch((error) => setStatus(error.message));
}

if (window.Office) {
  Office.onReady(initialize);
} else {
  initialize();
}
