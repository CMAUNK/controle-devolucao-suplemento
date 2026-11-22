/* global Office, Excel */
"use strict";

const state = { records: [], headers: [], fileName: "", importing: false, plan: null, fillPlan: null };
const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  $("inputFile").addEventListener("change", readSelectedFile);
  $("base").addEventListener("change", () => {
    hideReview();
    hideFillReview();
    clearStatus();
    renderPreview();
    populateFillColumns();
  });
  $("actionMode").addEventListener("change", renderActionMode);
  $("fillColumns").addEventListener("change", () => {
    hideFillReview();
    updateFillControls();
  });
  $("overwriteValues").addEventListener("change", hideFillReview);
  $("noCteList").addEventListener("change", updateFillControls);
  $("addRows").addEventListener("click", prepareInsertion);
  $("insertDirect").addEventListener("click", insertDirectly);
  $("clearSelection").addEventListener("click", clearSelection);
  $("checkFill").addEventListener("click", prepareColumnFill);
  $("confirmFill").addEventListener("click", fillExistingColumn);
  $("cancelFill").addEventListener("click", hideFillReview);
  $("confirmNewOnly").addEventListener("click", () => insertRows("new"));
  $("confirmAll").addEventListener("click", () => insertRows("all"));
  $("cancelReview").addEventListener("click", () => {
    hideReview();
    clearStatus();
  });
});

function parseTSV(text) {
  const rows = [];
  let row = [], cell = "", quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (quoted && ch === '"' && text[i + 1] === '"') { cell += '"'; i += 1; }
    else if (ch === '"') quoted = !quoted;
    else if (ch === "\t" && !quoted) { row.push(cell); cell = ""; }
    else if ((ch === "\n" || ch === "\r") && !quoted) {
      if (ch === "\r" && text[i + 1] === "\n") i += 1;
      row.push(cell); cell = "";
      if (row.some((value) => value !== "")) rows.push(row);
      row = [];
    } else cell += ch;
  }
  row.push(cell);
  if (row.some((value) => value !== "")) rows.push(row);
  if (quoted) throw new Error("O arquivo TSV parece estar incompleto: há aspas sem fechamento.");
  if (rows.length < 2) throw new Error("O arquivo não contém linhas de dados.");
  const headers = rows[0].map((value) => value.replace(/^\uFEFF/, "").trim());
  const records = rows.slice(1).map((values) => Object.fromEntries(
    headers.map((header, index) => [header, values[index] ?? ""])
  ));
  return { headers, records };
}

async function readSelectedFile() {
  clearStatus();
  hideReview();
  hideFillReview();
  const file = $("inputFile").files[0];
  state.records = [];
  state.headers = [];
  state.fileName = "";
  $("addRows").disabled = true;
  $("insertDirect").disabled = true;
  $("clearSelection").disabled = !file;
  populateFillColumns();
  if (!file) { populateFillColumns(); renderPreview(); return; }
  try {
    const parsed = parseTSV(await file.text());
    state.records = parsed.records;
    state.headers = parsed.headers;
    state.fileName = file.name;
    $("addRows").disabled = false;
    $("insertDirect").disabled = false;
    $("clearSelection").disabled = false;
    populateFillColumns();
    renderPreview();
  } catch (error) {
    $("addRows").disabled = true;
    $("insertDirect").disabled = true;
    setStatus(error.message, true);
    $("preview").textContent = "Não foi possível ler o arquivo.";
  }
}

function renderPreview() {
  if (!state.records.length) {
    $("preview").textContent = "Selecione um arquivo para ver uma prévia.";
    $("previewRows").replaceChildren();
    $("previewRows").classList.add("hidden");
    return;
  }
  const base = $("base").value;
  const is100km = base === "100km";
  const identifier = is100km ? "SM" : "MANIFESTO";
  const quantity = is100km ? "KM" : "VOLUMES";
  const ids = [...new Set(state.records.map((row) => row[identifier]).filter(Boolean))];
  const countedRows = is100km
    ? state.records.filter((row, index, rows) => rows.findIndex((candidate) => candidate[identifier] === row[identifier]) === index)
    : state.records;
  const total = countedRows.reduce((sum, row) => sum + (Number(String(row[quantity] || "").replace(/[^0-9.]/g, "")) || 0), 0);
  const destination = is100km ? "BASE" : `BASE-${base.toUpperCase()}`;
  $("preview").textContent = `${state.fileName}\nDestino: ${destination}\nLinhas no arquivo: ${state.records.length}\n${is100km ? "SMs" : "Manifestos"}: ${ids.join(", ") || "não identificados"}\n${is100km ? "KM somados" : "Volumes somados"}: ${total}${is100km ? " Km" : ""}`;
  renderSampleRows();
}

function renderActionMode() {
  const fillMode = $("actionMode").value === "fill";
  $("addRows").classList.toggle("hidden", fillMode);
  $("insertDirect").classList.toggle("hidden", fillMode);
  $("fillOptions").classList.toggle("hidden", !fillMode);
  hideReview();
  hideFillReview();
  if (fillMode) populateFillColumns();
}

function populateFillColumns() {
  const keyName = $("base").value === "100km" ? "SM" : "MANIFESTO";
  const candidates = state.headers.filter((header) => normalizeHeader(header) !== normalizeHeader(keyName));
  const container = $("fillColumns");
  const previouslySelected = new Set([...container.querySelectorAll("input:checked")].map((input) => normalizeHeader(input.value)));
  container.replaceChildren();
  for (const header of candidates) {
    const label = document.createElement("label");
    label.className = "check-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = header;
    checkbox.disabled = !state.records.length || state.importing;
    const normalized = normalizeHeader(header);
    checkbox.checked = previouslySelected.size
      ? previouslySelected.has(normalized)
      : ["VALOR", "CTE S"].includes(normalized);
    const text = document.createElement("span");
    text.textContent = header;
    label.append(checkbox, text);
    container.appendChild(label);
  }
  updateFillControls();
}

function selectedFillHeaders() {
  return [...$("fillColumns").querySelectorAll("input:checked")].map((input) => input.value);
}

function updateFillControls() {
  const selected = selectedFillHeaders();
  $("checkFill").disabled = !state.records.length || !selected.length || state.importing;
  $("fillColumns").querySelectorAll("input").forEach((input) => { input.disabled = !state.records.length || state.importing; });
  $("noCteList").querySelectorAll("input").forEach((input) => { input.disabled = state.importing; });
  $("confirmFill").disabled = state.importing;
}

function renderSampleRows() {
  const container = $("previewRows");
  container.replaceChildren();
  const visibleHeaders = state.headers.slice(0, 6);
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  for (const header of visibleHeaders) {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  for (const record of state.records.slice(0, 5)) {
    const row = document.createElement("tr");
    for (const header of visibleHeaders) {
      const cell = document.createElement("td");
      cell.textContent = String(record[header] ?? "");
      row.appendChild(cell);
    }
    tbody.appendChild(row);
  }
  table.appendChild(tbody);
  container.appendChild(table);
  if (state.records.length > 5) {
    const note = document.createElement("div");
    note.className = "sample-note";
    note.textContent = `Exibindo as primeiras 5 de ${state.records.length} linhas.`;
    container.appendChild(note);
  }
  container.classList.remove("hidden");
}

function normalizeHeader(value) {
  const normalized = String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
  if (["CTE S FALTANTE", "CTE S NAO RECEBIDO", "CTE S NAO RECEBIDOS"].includes(normalized)) return "CTE S";
  if (["DATA DE TRASFERENCIA", "DATA DE TRANSFERENCIA"].includes(normalized)) return "DATA DE TRANSFERENCIA";
  if (["DT DE TRASFERENCIA", "DT DE TRANSFERENCIA"].includes(normalized)) return "DT DE TRANSFERENCIA";
  return normalized;
}

function personRecordKey(number, cpf, name = "") {
  let cpfDigits = String(cpf || "").replace(/\D/g, "");
  if (cpfDigits.length === 10) cpfDigits = cpfDigits.padStart(11, "0");
  const person = cpfDigits || String(name || "").trim().toLocaleLowerCase("pt-BR");
  return `${String(number || "").trim()}|${person}`;
}

function applyColumnTextStyle(value, sample, header) {
  if (value === null || value === undefined) return value;
  let text = String(value);
  if (normalizeHeader(header) === "CPF") {
    let digits = text.replace(/\D/g, "");
    // A CPF beginning with zero can lose that zero if an upstream export
    // treated the identifier as a number instead of text.
    if (digits.length === 10) digits = digits.padStart(11, "0");
    if (digits.length === 11) text = digits.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
  }
  if (!normalizeHeader(header).includes("MOTORISTA")) return text;
  const letters = text.replace(/[^\p{L}]/gu, "");
  if (!letters) return text;
  // Match the common Portuguese name style while retaining lowercase connectors.
  const lowercaseWords = new Set(["a", "as", "ao", "aos", "da", "das", "de", "do", "dos", "e", "na", "nas", "no", "nos"]);
  return text.toLocaleLowerCase("pt-BR").replace(/[\p{L}]+/gu, (word, offset) => {
    if (offset > 0 && lowercaseWords.has(word)) return word;
    return word.charAt(0).toLocaleUpperCase("pt-BR") + word.slice(1);
  });
}

function columnLetter(zeroBased) {
  let n = zeroBased + 1, result = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function excelDateSerial(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/);
  if (!match) return value;
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  return (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}

function excelAmount(value) {
  if (typeof value === "number") return value;
  let token = String(value ?? "").trim().replace(/[^\d,.-]/g, "");
  if (!token) return value;
  if (token.includes(",") && token.includes(".")) {
    token = token.lastIndexOf(",") > token.lastIndexOf(".")
      ? token.replace(/\./g, "").replace(",", ".")
      : token.replace(/,/g, "");
  } else if (token.includes(",")) {
    token = token.replace(",", ".");
  }
  const amount = Number(token);
  return Number.isFinite(amount) ? amount : value;
}

function hideReview() {
  state.plan = null;
  $("review").classList.add("hidden");
  $("duplicateWarning").classList.add("hidden");
  $("duplicateText").textContent = "";
  $("reviewSummary").replaceChildren();
}

async function syncExcelStage(context, stage) {
  try {
    await context.sync();
  } catch (error) {
    const message = error?.message || String(error);
    const code = error?.code ? ` · código ${error.code}` : "";
    const location = error?.debugInfo?.errorLocation ? ` · operação ${error.debugInfo.errorLocation}` : "";
    throw new Error(`${message} · etapa: ${stage}${code}${location}`);
  }
}

function clearSelection() {
  if (state.importing) return;
  $("inputFile").value = "";
  state.records = [];
  state.headers = [];
  state.fileName = "";
  hideReview();
  hideFillReview();
  $("addRows").disabled = true;
  $("insertDirect").disabled = true;
  $("clearSelection").disabled = true;
  clearStatus();
  renderPreview();
  populateFillColumns();
}

function ensureFileMatchesBase(base) {
  if (base === "100km") {
    const required = ["SM", "DATA", "CPF", "KM"];
    const missing = required.filter((header) => !state.headers.some((candidate) => normalizeHeader(candidate) === normalizeHeader(header)));
    if (missing.length) throw new Error(`Este arquivo não parece ser da planilha 100km. Não encontrei: ${missing.join(", ")}.`);
    return;
  }
  const expected = base === "recebendo" ? "DATA DE TRANSFERENCIA" : "DT DE TRANSFERENCIA";
  if (!state.headers.some((header) => normalizeHeader(header) === normalizeHeader(expected))) {
    throw new Error(`O arquivo não corresponde à base ${base === "recebendo" ? "RECEBENDO" : "EXPEDINDO"}. Selecione o TSV dessa base.`);
  }
}

async function inspectOrInsert(mode = null) {
  const base = $("base").value;
  ensureFileMatchesBase(base);
  const sheetName = base === "100km" ? "BASE" : (base === "recebendo" ? "BASE-RECEBENDO" : "BASE-EXPEDINDO");
  const inputKeyName = base === "100km" ? "SM" : "MANIFESTO";
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await syncExcelStage(context, "listar as abas da pasta de trabalho");
    const expectedSheetName = normalizeHeader(sheetName);
    const matchingSheets = worksheets.items.filter((item) => normalizeHeader(item.name) === expectedSheetName);
    if (!matchingSheets.length) {
      const available = worksheets.items.map((item) => item.name).join(", ") || "nenhuma";
      throw new Error(`Não encontrei a aba ${sheetName}. Abas disponíveis: ${available}.`);
    }
    if (matchingSheets.length > 1) {
      throw new Error(`Há mais de uma aba com nome equivalente a ${sheetName}. Renomeie para deixar somente uma.`);
    }
    const sheet = worksheets.getItem(matchingSheets[0].name);
    const tables = sheet.tables;
    tables.load("items/name");
    await syncExcelStage(context, `localizar a tabela na aba ${matchingSheets[0].name}`);
    if (!tables.items.length) throw new Error(`Não encontrei uma tabela na aba ${sheetName}.`);
    if (tables.items.length > 1) throw new Error(`Há mais de uma tabela na aba ${sheetName}; não vou escolher uma automaticamente.`);

    const table = tables.getItemAt(0);
    const headerRange = table.getHeaderRowRange();
    headerRange.load("values");
    const countResult = table.rows.getCount();
    await syncExcelStage(context, "ler os cabeçalhos e a quantidade de linhas da tabela");
    const existingCount = countResult.value;
    const headers = headerRange.values[0].map((value) => String(value || "").trim());
    const columnMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
    const uniqueColumn = columnMap.get(normalizeHeader(inputKeyName));
    if (uniqueColumn === undefined) throw new Error(`Não encontrei a coluna ${inputKeyName} na tabela da aba ${sheetName}.`);
    const inputKeyHeader = state.headers.find((header) => normalizeHeader(header) === normalizeHeader(inputKeyName));
    if (!inputKeyHeader) throw new Error(`O arquivo não tem a coluna ${inputKeyName}.`);
    const inputCpfHeader = base === "100km" ? state.headers.find((header) => normalizeHeader(header) === normalizeHeader("CPF")) : null;
    const inputNameHeader = base === "100km" ? state.headers.find((header) => normalizeHeader(header) === normalizeHeader("Motorista")) : null;

    const ignored100kmHeaders = new Set(["MATRICULA", "TOTAL"]);
    const mappedInputColumns = new Map();
    for (const record of state.records) {
      if (!String(record[inputKeyHeader] || "").trim()) throw new Error(`Há uma linha sem ${inputKeyName}; revise o arquivo antes de inserir.`);
      if (base === "100km" && !String(record[inputCpfHeader] || "").replace(/\D/g, "")) throw new Error("Há uma linha sem CPF; confira o motorista e o ajudante antes de inserir.");
      for (const inputHeader of Object.keys(record)) {
        if (base === "100km" && ignored100kmHeaders.has(normalizeHeader(inputHeader))) continue;
        const targetColumn = columnMap.get(normalizeHeader(inputHeader));
        if (targetColumn !== undefined) mappedInputColumns.set(normalizeHeader(inputHeader), targetColumn);
      }
    }
    const existing = new Set();
    const personColumn = base === "100km" ? columnMap.get(normalizeHeader("CPF")) : undefined;
    if (base === "100km" && personColumn === undefined) throw new Error("Para importar uma linha por pessoa, a tabela BASE precisa ter a coluna CPF.");
    if (existingCount > 0) {
      const manifestValues = table.columns.getItemAt(uniqueColumn).getDataBodyRange();
      manifestValues.load("values");
      let personValues = null;
      if (base === "100km") {
        personValues = table.columns.getItemAt(personColumn).getDataBodyRange();
        personValues.load("values");
      }
      await syncExcelStage(context, "comparar os registros já existentes");
      manifestValues.values.flat().forEach((value, index) => {
        const number = String(value || "").trim();
        if (!number) return;
        if (base === "100km") {
          existing.add(personRecordKey(number, personValues.values[index]?.[0]));
        } else {
          existing.add(number);
        }
      });
    }
    const seenInFile = new Set();
    const duplicateList = [];
    const newRecordIndexes = [];
    state.records.forEach((record, index) => {
      const number = String(record[inputKeyHeader] || "").trim();
      const identity = base === "100km"
        ? personRecordKey(number, record[inputCpfHeader], record[inputNameHeader])
        : number;
      if (existing.has(identity) || seenInFile.has(identity)) duplicateList.push(number);
      else newRecordIndexes.push(index);
      seenInFile.add(identity);
    });

    const bodyRange = existingCount ? table.getDataBodyRange() : null;
    let bodyValues = [], bodyFormulas = [];
    if (bodyRange) {
      bodyRange.load("values,formulas");
      await syncExcelStage(context, "ler os dados e as fórmulas atuais");
      bodyValues = bodyRange.values;
      bodyFormulas = bodyRange.formulas;
    }
    const formulaColumns = new Set();
    const formulaTemplateRows = new Map();
    bodyFormulas.forEach((row, rowIndex) => row.forEach((formula, col) => {
      if (typeof formula === "string" && formula.startsWith("=")) {
        formulaColumns.add(col);
        formulaTemplateRows.set(col, rowIndex);
      }
    }));
    const protected100kmColumns = base === "100km"
      ? [...ignored100kmHeaders].map((header) => columnMap.get(normalizeHeader(header))).filter((col) => col !== undefined)
      : [];
    const columnsToCheckForData = new Set([...mappedInputColumns.values(), ...protected100kmColumns]);
    const reusableRows = [];
    for (let row = 0; row < existingCount; row += 1) {
      const hasIncomingData = [...columnsToCheckForData].some((col) => {
        if (formulaColumns.has(col)) return false;
        const value = bodyValues[row]?.[col];
        return value !== null && value !== undefined && String(value).trim() !== "";
      });
      if (!hasIncomingData && String(bodyValues[row]?.[uniqueColumn] ?? "").trim() === "") reusableRows.push(row);
    }

    const candidateIndexes = mode === "new" ? newRecordIndexes : state.records.map((_, index) => index);
    const reusedCount = Math.min(reusableRows.length, candidateIndexes.length);
    const plan = {
      base, sheetName, total: state.records.length, existingCount,
      availableBlankRows: reusableRows.length,
      count: candidateIndexes.length, reused: reusedCount,
      appended: candidateIndexes.length - reusedCount,
      duplicates: [...new Set(duplicateList)], newCount: newRecordIndexes.length,
      formulas: [...formulaColumns].map((col) => headers[col]),
      skipped: state.records.length - candidateIndexes.length,
    };
    if (mode === null || candidateIndexes.length === 0) return plan;
    const previousPlan = state.plan;
    const workbookChanged = !previousPlan || plan.existingCount !== previousPlan.existingCount ||
      plan.availableBlankRows !== previousPlan.availableBlankRows ||
      plan.duplicates.join("|") !== previousPlan.duplicates.join("|") ||
      plan.formulas.join("|") !== previousPlan.formulas.join("|");
    if (workbookChanged) return { ...plan, needsReview: true };

    const inputValues = candidateIndexes.map((index) => {
      const record = state.records[index];
      const values = new Map();
      for (const [inputHeader, rawValue] of Object.entries(record)) {
        if (base === "100km" && ignored100kmHeaders.has(normalizeHeader(inputHeader))) continue;
        const targetColumn = columnMap.get(normalizeHeader(inputHeader));
        if (targetColumn === undefined || rawValue === "" || rawValue == null) continue;
        let value = rawValue;
        const normalizedHeader = normalizeHeader(inputHeader);
        if (["DATA DE TRANSFERENCIA", "DT DE TRANSFERENCIA", "DATA"].includes(normalizedHeader)) value = excelDateSerial(rawValue);
        if (normalizedHeader === "VALOR") value = excelAmount(rawValue);
        values.set(targetColumn, value);
      }
      return values;
    });
    const assignments = [];
    for (let i = 0; i < reusedCount; i += 1) assignments.push({ row: reusableRows[i], values: inputValues[i], existing: true });
    if (plan.appended > 0) {
      const blankRows = Array.from({ length: plan.appended }, () => new Array(headers.length).fill(""));
      const added = table.rows.add(-1, blankRows, true);
      added.load("index");
      await syncExcelStage(context, "adicionar linhas à tabela");
      for (let i = 0; i < plan.appended; i += 1) assignments.push({ row: added.index + i, values: inputValues[reusedCount + i], existing: false });
    }
    const formattedBody = table.getDataBodyRange();
    const findFormatTemplateRow = (targetRow) => {
      // Use the closest populated row so the inserted cells inherit the
      // workbook's font, size, number format, alignment, borders, and fills.
      for (let row = Math.min(targetRow - 1, bodyValues.length - 1); row >= 0; row -= 1) {
        if (bodyValues[row]?.some((value) => value !== null && value !== undefined && String(value).trim() !== "")) return row;
      }
      for (let row = Math.max(0, targetRow + 1); row < bodyValues.length; row += 1) {
        if (bodyValues[row]?.some((value) => value !== null && value !== undefined && String(value).trim() !== "")) return row;
      }
      return bodyValues.length ? 0 : undefined;
    };
    const formatTemplateRows = new Map();
    const alignmentSources = new Map();
    for (const assignment of assignments) {
      const templateRow = findFormatTemplateRow(assignment.row);
      formatTemplateRows.set(assignment, templateRow);
      if (templateRow === undefined || templateRow === assignment.row) continue;
      for (const col of assignment.values.keys()) {
        const source = formattedBody.getCell(templateRow, col);
        source.format.load("horizontalAlignment");
        alignmentSources.set(`${assignment.row}:${col}`, source);
      }
    }
    if (alignmentSources.size) await syncExcelStage(context, "ler o formato das linhas-modelo");
    let preservedFormulas = 0, copiedFormulas = 0;
    for (const assignment of assignments) {
      const formatTemplateRow = formatTemplateRows.get(assignment);
      if (formatTemplateRow !== undefined && formatTemplateRow !== assignment.row) {
        formattedBody.getCell(assignment.row, 0).getResizedRange(0, headers.length - 1)
          .copyFrom(formattedBody.getCell(formatTemplateRow, 0).getResizedRange(0, headers.length - 1), Excel.RangeCopyType.formats);
      }
      for (const [col, value] of assignment.values) {
        if (!formulaColumns.has(col)) {
          const target = table.getDataBodyRange().getCell(assignment.row, col);
          let cellValue = applyColumnTextStyle(value, bodyValues[formatTemplateRow]?.[col], headers[col]);
          const alignmentSource = alignmentSources.get(`${assignment.row}:${col}`);
          if (alignmentSource) target.format.horizontalAlignment = alignmentSource.format.horizontalAlignment;
          if (normalizeHeader(headers[col]) === "CPF") {
            target.numberFormat = [["@"]];
            target.format.horizontalAlignment = "Left";
          }
          if (base === "100km" && normalizeHeader(headers[col]) === "KM") {
            const numericKm = Number(String(value).replace(/\s*km\s*$/i, "").trim().replace(",", "."));
            if (Number.isFinite(numericKm)) {
              cellValue = numericKm;
              target.numberFormat = [['0" Km"']];
            }
          }
          target.values = [[cellValue]];
        }
      }
      for (const col of formulaColumns) {
        const target = table.getDataBodyRange().getCell(assignment.row, col);
        const alreadyHasFormula = assignment.existing && typeof bodyFormulas[assignment.row]?.[col] === "string" && bodyFormulas[assignment.row][col].startsWith("=");
        const templateRow = formulaTemplateRows.get(col);
        if (alreadyHasFormula) preservedFormulas += 1;
        else if (templateRow !== undefined && templateRow !== assignment.row) {
          target.copyFrom(table.getDataBodyRange().getCell(templateRow, col), Excel.RangeCopyType.formulas);
          copiedFormulas += 1;
        }
      }
    }
    await syncExcelStage(context, "preencher as linhas e copiar as fórmulas");
    return { ...plan, preservedFormulas, copiedFormulas };
  });
}

function addReviewLine(container, label, value) {
  const line = document.createElement("p");
  const strong = document.createElement("strong");
  strong.textContent = `${label}: `;
  line.append(strong, document.createTextNode(String(value)));
  container.appendChild(line);
}

function showReview(plan) {
  state.plan = plan;
  const container = $("reviewSummary");
  container.replaceChildren();
  addReviewLine(container, "Destino", plan.sheetName);
  addReviewLine(container, "Linhas selecionadas", state.records.length);
  addReviewLine(container, "Linhas que serão inseridas", plan.count);
  addReviewLine(container, "Linhas em branco que serão reutilizadas", plan.reused);
  addReviewLine(container, "Novas linhas necessárias", plan.appended);
  addReviewLine(container, "Fórmulas detectadas", plan.formulas.join(", ") || "nenhuma coluna com fórmula encontrada");
  const hasDuplicates = plan.duplicates.length > 0;
  $("duplicateWarning").classList.toggle("hidden", !hasDuplicates);
  $("duplicateText").textContent = hasDuplicates ? plan.duplicates.join(", ") : "";
  $("confirmNewOnly").classList.toggle("hidden", !hasDuplicates);
  $("confirmNewOnly").textContent = `Inserir somente novos (${plan.newCount})`;
  $("confirmAll").textContent = hasDuplicates ? "Inserir todos" : "Confirmar e inserir";
  $("review").classList.remove("hidden");
}

function hideFillReview() {
  state.fillPlan = null;
  $("fillReview").replaceChildren();
  $("fillReview").classList.add("hidden");
  $("noCteWarning").classList.add("hidden");
  $("noCteText").textContent = "";
  $("noCteList").replaceChildren();
  $("fillConfirmActions").classList.add("hidden");
  updateFillControls();
}

function normalizedRecordId(value) {
  return String(value ?? "").trim().replace(/\.0+$/, "");
}

function convertColumnValue(header, value) {
  const normalized = normalizeHeader(header);
  if (normalized === "VALOR") return excelAmount(value);
  if (["DATA DE TRANSFERENCIA", "DT DE TRANSFERENCIA", "DATA"].includes(normalized)) return excelDateSerial(value);
  return value;
}

async function inspectColumnFill(apply = false) {
  const base = $("base").value;
  ensureFileMatchesBase(base);
  const sheetName = base === "100km" ? "BASE" : (base === "recebendo" ? "BASE-RECEBENDO" : "BASE-EXPEDINDO");
  const keyName = base === "100km" ? "SM" : "MANIFESTO";
  const inputKeyHeader = state.headers.find((header) => normalizeHeader(header) === normalizeHeader(keyName));
  const selectedHeaders = selectedFillHeaders();
  if (!inputKeyHeader) throw new Error("O arquivo não contém a coluna " + keyName + ".");
  if (!selectedHeaders.length) throw new Error("Marque pelo menos uma coluna para preencher.");
  const inputHeaders = selectedHeaders.map((selected) => state.headers.find((header) => normalizeHeader(header) === normalizeHeader(selected)));
  if (inputHeaders.some((header) => !header)) throw new Error("Uma das colunas selecionadas não existe no arquivo.");
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await syncExcelStage(context, "listar as abas da pasta de trabalho");
    const expectedSheetName = normalizeHeader(sheetName);
    const matchingSheets = worksheets.items.filter((item) => normalizeHeader(item.name) === expectedSheetName);
    if (!matchingSheets.length) throw new Error("Não encontrei a aba " + sheetName + ".");
    if (matchingSheets.length > 1) throw new Error("Há mais de uma aba com nome equivalente a " + sheetName + ".");

    const sheet = worksheets.getItem(matchingSheets[0].name);
    const tables = sheet.tables;
    tables.load("items/name");
    await syncExcelStage(context, "localizar a tabela na aba " + matchingSheets[0].name);
    if (tables.items.length !== 1) throw new Error(tables.items.length ? "Há mais de uma tabela na aba " + sheetName + "." : "Não encontrei uma tabela na aba " + sheetName + ".");

    const table = tables.getItemAt(0);
    const headerRange = table.getHeaderRowRange();
    headerRange.load("values");
    const countResult = table.rows.getCount();
    await syncExcelStage(context, "ler os cabeçalhos e a quantidade de linhas da tabela");
    const headers = headerRange.values[0].map((value) => String(value || "").trim());
    const columnMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
    const keyColumn = columnMap.get(normalizeHeader(keyName));
    if (keyColumn === undefined) throw new Error("Não encontrei a coluna " + keyName + " na tabela.");
    const targetColumns = new Map();
    for (let i = 0; i < selectedHeaders.length; i += 1) {
      const targetColumn = columnMap.get(normalizeHeader(selectedHeaders[i]));
      if (targetColumn === undefined) throw new Error("A coluna " + selectedHeaders[i] + " não existe na tabela de destino.");
      targetColumns.set(inputHeaders[i], targetColumn);
    }

    const existingCount = countResult.value;
    let idValues = [], body = null;
    const columnData = new Map();
    if (existingCount > 0) {
      body = table.getDataBodyRange();
      const ids = table.columns.getItemAt(keyColumn).getDataBodyRange();
      ids.load("values");
      for (const inputHeader of inputHeaders) {
        const values = table.columns.getItemAt(targetColumns.get(inputHeader)).getDataBodyRange();
        values.load("values,formulas");
        columnData.set(inputHeader, values);
      }
      await syncExcelStage(context, "localizar os manifestos e ler as colunas selecionadas");
      idValues = ids.values.flat();
    }

    const targetRows = new Map();
    idValues.forEach((value, rowIndex) => {
      const key = normalizedRecordId(value);
      if (!key) return;
      if (!targetRows.has(key)) targetRows.set(key, []);
      targetRows.get(key).push(rowIndex);
    });
    const sourceValues = new Map();
    const sourceConflicts = new Set();
    const hasCtes = new Map();
    const missingInputValues = Object.fromEntries(inputHeaders.map((header) => [header, 0]));
    const cteSelected = inputHeaders.some((header) => normalizeHeader(header) === "CTE S");
    for (const record of state.records) {
      const key = normalizedRecordId(record[inputKeyHeader]);
      if (!key) throw new Error("Há uma linha sem " + keyName + "; revise o arquivo.");
      if (!sourceValues.has(key)) sourceValues.set(key, new Map());
      if (cteSelected) hasCtes.set(key, hasCtes.get(key) || Boolean(String(record[inputHeaders.find((header) => normalizeHeader(header) === "CTE S")] ?? "").trim()));
      for (const inputHeader of inputHeaders) {
        const rawValue = record[inputHeader];
        if (rawValue === "" || rawValue == null) { missingInputValues[inputHeader] += 1; continue; }
        const value = convertColumnValue(inputHeader, rawValue);
        if (value === "" || value == null || (typeof value === "number" && !Number.isFinite(value))) { missingInputValues[inputHeader] += 1; continue; }
        const sourceKey = key + "\u001f" + normalizeHeader(inputHeader);
        const valuesForKey = sourceValues.get(key);
        if (valuesForKey.has(inputHeader) && String(valuesForKey.get(inputHeader)) !== String(value)) sourceConflicts.add(sourceKey);
        else valuesForKey.set(inputHeader, value);
      }
    }

    const updates = [];
    let notFound = 0, ambiguous = 0, protectedByFormula = 0, alreadyFilled = 0, unchanged = 0;
    for (const [key, valuesForKey] of sourceValues) {
      const rows = targetRows.get(key) || [];
      if (!rows.length) { notFound += 1; continue; }
      if (rows.length !== 1) { ambiguous += 1; continue; }
      const rowIndex = rows[0];
      for (const inputHeader of inputHeaders) {
        const value = valuesForKey.get(inputHeader);
        if (value === undefined) continue;
        if (sourceConflicts.has(key + "\u001f" + normalizeHeader(inputHeader))) { ambiguous += 1; continue; }
        const data = columnData.get(inputHeader);
        const formula = data.formulas.flat()[rowIndex];
        if (typeof formula === "string" && formula.startsWith("=")) { protectedByFormula += 1; continue; }
        const current = data.values.flat()[rowIndex];
        const isBlank = current === null || current === undefined || String(current).trim() === "";
        if (!isBlank && String(current) === String(value)) { unchanged += 1; continue; }
        if (!isBlank && !$("overwriteValues").checked) { alreadyFilled += 1; continue; }
        updates.push({ rowIndex, key, value, current: current ?? "", inputHeader, targetColumn: targetColumns.get(inputHeader) });
      }
    }
    const keysWithOtherUpdates = new Set(updates.filter((item) => normalizeHeader(item.inputHeader) !== "CTE S").map((item) => item.key));
    const noCteManifestos = cteSelected
      ? [...hasCtes].filter(([key, found]) => !found && targetRows.has(key) && keysWithOtherUpdates.has(key)).map(([key]) => key).sort()
      : [];
    const plan = {
      sheetName, keyName, selectedHeaders, sourceCount: sourceValues.size,
      updates, notFound, ambiguous, protectedByFormula, alreadyFilled, unchanged,
      missingInputValues, noCteManifestos,
      signature: JSON.stringify({ headers, updates, notFound, ambiguous, protectedByFormula, alreadyFilled, unchanged, missingInputValues, noCteManifestos }),
    };
    if (apply) {
      if (plan.signature !== state.fillPlan?.signature) return { plan, changed: true };
      const approvedNoCte = new Set([...$("noCteList").querySelectorAll("input:checked")].map((input) => input.value));
      const skippedNoCte = plan.noCteManifestos.filter((key) => !approvedNoCte.has(key));
      const updatesToApply = plan.updates.filter((item) => !skippedNoCte.includes(item.key));
      for (const item of updatesToApply) body.getCell(item.rowIndex, item.targetColumn).values = [[item.value]];
      if (updatesToApply.length) await syncExcelStage(context, "preencher somente as colunas selecionadas");
      return { plan, appliedCount: updatesToApply.length, skippedNoCte };
    }
    return { plan };
  });
}

function showFillReview(plan) {
  state.fillPlan = plan;
  const container = $("fillReview");
  container.replaceChildren();
  addReviewLine(container, "Destino", plan.sheetName);
  addReviewLine(container, "Colunas selecionadas", plan.selectedHeaders.join(", "));
  addReviewLine(container, "Células que serão preenchidas", plan.updates.length);
  addReviewLine(container, "Células já preenchidas e preservadas", plan.alreadyFilled);
  addReviewLine(container, "Já estavam iguais", plan.unchanged);
  addReviewLine(container, "Manifestos não encontrados", plan.notFound);
  for (const [header, count] of Object.entries(plan.missingInputValues)) {
    if (count) addReviewLine(container, "Sem valor no arquivo — " + header, count);
  }
  addReviewLine(container, "Correspondências ambíguas ou com fórmula", plan.ambiguous + plan.protectedByFormula);
  container.classList.remove("hidden");
  const hasNoCtes = plan.noCteManifestos.length > 0;
  $("noCteWarning").classList.toggle("hidden", !hasNoCtes);
  $("noCteText").textContent = hasNoCtes
    ? "Verifiquei que estes manifestos também não têm CTE’S. Deseja fazer o preenchimento para cada um?"
    : "";
  const noCteList = $("noCteList");
  noCteList.replaceChildren();
  for (const number of plan.noCteManifestos) {
    const label = document.createElement("label");
    label.className = "check-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = number;
    const text = document.createElement("span");
    text.textContent = "Verifiquei que o manifesto " + number + " também não tem CTE’S. Deseja fazer o preenchimento também?";
    label.append(checkbox, text);
    noCteList.appendChild(label);
  }
  $("fillConfirmActions").classList.toggle("hidden", plan.updates.length === 0);
  updateFillControls();
}

async function prepareColumnFill() {
  if (state.importing || !state.records.length) return;
  state.importing = true;
  updateFillControls();
  $("checkFill").disabled = true;
  clearStatus();
  hideFillReview();
  try {
    const result = await inspectColumnFill();
    showFillReview(result.plan);
    setStatus("Confira a coluna e a quantidade de células antes de preencher. Nenhuma linha será adicionada.", false);
  } catch (error) {
    setStatus(error.message || "Não foi possível conferir a coluna.", true);
  } finally {
    state.importing = false;
    updateFillControls();
  }
}

async function fillExistingColumn() {
  if (state.importing || !state.fillPlan?.updates.length) return;
  state.importing = true;
  updateFillControls();
  $("confirmFill").disabled = true;
  $("checkFill").disabled = true;
  clearStatus();
  try {
    const result = await inspectColumnFill(true);
    if (result.changed) {
      showFillReview(result.plan);
      setStatus("A planilha mudou desde a conferência. Revise o novo resumo antes de confirmar.", false);
      return;
    }
    const completed = result.appliedCount;
    const selectedHeaders = state.fillPlan.selectedHeaders.join(", ");
    const skippedText = result.skippedNoCte.length
      ? " Manifestos sem CTE’S não preenchidos: " + result.skippedNoCte.join(", ") + "."
      : "";
    hideFillReview();
    setStatus(completed + " célula(s) preenchida(s) nas colunas " + selectedHeaders + ". Os demais campos e linhas foram preservados." + skippedText, false, true);
  } catch (error) {
    setStatus(error.message || "Não foi possível preencher a coluna.", true);
  } finally {
    state.importing = false;
    $("confirmFill").disabled = false;
    updateFillControls();
  }
}

async function prepareInsertion() {
  if (state.importing || !state.records.length) return;
  state.importing = true;
  $("addRows").disabled = true;
  $("insertDirect").disabled = true;
  clearStatus();
  try {
    const plan = await inspectOrInsert(null);
    showReview(plan);
    setStatus("Confira o destino, os duplicados e as linhas antes de confirmar.", false);
  } catch (error) {
    setStatus(error.message || "Não foi possível conferir a planilha.", true);
  } finally {
    state.importing = false;
    $("addRows").disabled = !state.records.length;
    $("insertDirect").disabled = !state.records.length;
    $("clearSelection").disabled = !$("inputFile").files.length;
  }
}

async function insertDirectly() {
  if (state.importing || !state.records.length) return;
  state.importing = true;
  $("addRows").disabled = true;
  $("insertDirect").disabled = true;
  $("clearSelection").disabled = true;
  hideReview();
  clearStatus();
  try {
    const initialPlan = await inspectOrInsert(null);
    if (initialPlan.count === 0) {
      setStatus("O arquivo não contém linhas para inserir.", false);
      return;
    }
    state.plan = initialPlan;
    const result = await inspectOrInsert("new");
    if (result.needsReview) {
      showReview(result);
      setStatus("A planilha mudou durante a operação. Confira o resumo antes de inserir.", false);
      return;
    }
    if (result.count === 0) {
      setStatus("Todos os registros deste arquivo já estão na tabela. Nenhuma linha foi inserida.", false);
      return;
    }
    const skippedText = result.skipped ? ` ${result.skipped} registro(s) repetido(s) foram ignorados automaticamente.` : "";
    const formulaText = result.formulas.length
      ? ` Fórmulas preservadas: ${result.preservedFormulas}; copiadas para novas linhas: ${result.copiedFormulas} (${result.formulas.join(", ")}).`
      : " Não foram detectadas fórmulas na tabela.";
    setStatus(`${result.count} linha(s) inserida(s) diretamente em ${result.sheetName}: ${result.reused} linha(s) em branco reutilizada(s) e ${result.appended} nova(s) criada(s).${skippedText}${formulaText}`, false, true);
  } catch (error) {
    setStatus(error.message || "Não foi possível adicionar as linhas.", true);
  } finally {
    state.importing = false;
    $("addRows").disabled = !state.records.length;
    $("insertDirect").disabled = !state.records.length;
    $("clearSelection").disabled = !$("inputFile").files.length;
  }
}

async function insertRows(mode) {
  if (state.importing || !state.records.length) return;
  state.importing = true;
  $("addRows").disabled = true;
  $("confirmNewOnly").disabled = true;
  $("confirmAll").disabled = true;
  clearStatus();
  try {
    const result = await inspectOrInsert(mode);
    if (result.needsReview) {
      showReview(result);
      setStatus("A planilha mudou desde a conferência. Revise o resumo atualizado antes de confirmar.", false);
      return;
    }
    if (result.count === 0) {
      hideReview();
      setStatus("Todos os identificadores deste arquivo já estão na tabela. Nenhuma linha foi inserida.", false);
      return;
    }
    hideReview();
    const skippedText = result.skipped ? ` ${result.skipped} registro(s) repetido(s) foram ignorados.` : "";
    const formulaText = result.formulas.length
      ? ` Fórmulas preservadas: ${result.preservedFormulas}; copiadas para novas linhas: ${result.copiedFormulas} (${result.formulas.join(", ")}).`
      : " Não foram detectadas fórmulas na tabela.";
    setStatus(`${result.count} linha(s) inserida(s) em ${result.sheetName}: ${result.reused} linha(s) em branco reutilizada(s) e ${result.appended} nova(s) criada(s).${skippedText}${formulaText}`, false, true);
  } catch (error) {
    setStatus(error.message || "Não foi possível adicionar as linhas.", true);
  } finally {
    state.importing = false;
    $("addRows").disabled = !state.records.length;
    $("insertDirect").disabled = !state.records.length;
    $("confirmNewOnly").disabled = false;
    $("confirmAll").disabled = false;
    $("clearSelection").disabled = !$("inputFile").files.length;
  }
}

function setStatus(message, isError, isSuccess = false) {
  const status = $("status");
  status.textContent = message;
  status.className = isError ? "error" : (isSuccess ? "success" : "");
}

function clearStatus() { setStatus("", false); }
