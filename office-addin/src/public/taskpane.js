const translations = {
  en: {
    title: "Function Atlas",
    subtitle: "Searchable documentation for Evalytics worksheet functions.",
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
    tutorial: "Tutorial",
    tutorialReady: "Tutorial sheet created.",
    tutorialUnavailable: "Tutorial is available only inside Excel.",
    copyLog: "Copy log",
    logCopied: "Log copied.",
    required: "required"
  },
  cs: {
    title: "Atlas funkcí",
    subtitle: "Vyhledávatelná dokumentace funkcí Evalytics pro Excel.",
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
    tutorial: "Tutorial",
    tutorialReady: "Tutorialový list byl vytvořen.",
    tutorialUnavailable: "Tutorial je dostupný jen uvnitř Excelu.",
    copyLog: "Kopírovat log",
    logCopied: "Log zkopírován.",
    required: "povinné"
  }
};

let language = navigator.language.toLowerCase().startsWith("cs") ? "cs" : "en";
let functions = [];
let localizedDocs = { cs: {}, en: {} };
let search = "";
const buildStamp = window.EVALYTICS_BUILD_STAMP ?? "dev";
const hiddenDocumentationFunctions = new Set(["VERSION", "PING"]);

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
  return localizedInfo(fn).category ?? "Other";
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
    en: ["General", "Descriptive", "Tests", "Pivot", "Diagnostics", "Other"],
    cs: ["Obecné", "Popisné", "Testy", "PIVOT", "Diagnostika", "Other"]
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
  button.className = "tutorial-button";
  button.type = "button";
  button.textContent = t("tutorial");
  button.addEventListener("click", () => runTutorial(fn.name));
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
  "PARSE.NUMBER": {
    customRun: runParseNumberTutorial
  },
  "WELCH.TEST.2S.G": {
    customRun: runWelchTutorialWithDiagnostics
  }
};

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
    "4,5"
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
  output.forEach((item) => {
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
    .forEach(([categoryName, groupFunctions], groupIndex) => {
      const section = document.createElement("section");
      section.className = "function-section";

      const heading = document.createElement("button");
      heading.className = "section-heading";
      heading.type = "button";
      heading.textContent = `${categoryName} (${groupFunctions.length})`;
      section.append(heading);

      const sectionBody = document.createElement("div");
      sectionBody.className = "section-body";
      sectionBody.hidden = groupIndex !== 0;
      section.append(sectionBody);

      heading.addEventListener("click", () => {
        sectionBody.hidden = !sectionBody.hidden;
        section.classList.toggle("collapsed", sectionBody.hidden);
      });

      groupFunctions.forEach((fn, index) => {
    const fragment = template.content.cloneNode(true);
    const card = fragment.querySelector(".function-card");
    const header = fragment.querySelector(".function-header");
    const name = fragment.querySelector(".function-name");
    const category = fragment.querySelector(".function-category");
    const description = fragment.querySelector(".description");
    const formulaLabel = fragment.querySelector(".formula-label");
    const signature = fragment.querySelector(".signature");
    const parameters = fragment.querySelector(".parameters");
    const outputs = fragment.querySelector(".outputs");
    const copyButton = fragment.querySelector(".copy-button");
    const actions = fragment.querySelector(".actions");

    card.classList.toggle("open", groupIndex === 0 && index === 0);
    name.textContent = fn.name;
    category.textContent = categoryName;
    description.textContent = displaySummary(fn);
    formulaLabel.textContent = t("formula");
    signature.textContent = fn.signature;
    copyButton.textContent = t("copy");
    renderParameters(parameters, fn);
    renderOutput(outputs, fn);
    addTutorialButton(actions, fn);

    header.addEventListener("click", () => {
      card.classList.toggle("open");
    });

    copyButton.addEventListener("click", async () => {
      await copyText(fn.signature);
      setStatus(t("copied"));
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
  const formula = "=WELCH.TEST.2S.G(A1:A101,B1:B101,1,0.05,0)";
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
    stamp.textContent = `Build ${buildStamp}`;
  }

  document.querySelectorAll("[data-lang]").forEach((button) => {
    button.addEventListener("click", () => {
      language = button.dataset.lang;
      render();
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

  translateStaticText();
  debugLog("Panel initialized", { buildStamp, hasExcel: Boolean(window.Excel), hasOffice: Boolean(window.Office) });
  loadFunctions().catch((error) => setStatus(error.message));
}

if (window.Office) {
  Office.onReady(initialize);
} else {
  initialize();
}
