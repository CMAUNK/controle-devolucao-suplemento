/* global Office, Excel */
"use strict";

const state = { records: [], headers: [], fileName: "", importing: false, plan: null };
const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  $("inputFile").addEventListener("change", readSelectedFile);
  $("base").addEventListener("change", () => {
    hideReview();
    clearStatus();
    renderPreview();
  });
  $("addRows").addEventListener("click", prepareInsertion);
  $("insertDirect").addEventListener("click", insertDirectly);
  $("clearSelection").addEventListener("click", clearSelection);
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
  const file = $("inputFile").files[0];
  state.records = [];
  state.headers = [];
  state.fileName = "";
  $("addRows").disabled = true;
  $("insertDirect").disabled = true;
  $("clearSelection").disabled = !file;
  if (!file) { renderPreview(); return; }
  try {
    const parsed = parseTSV(await file.text());
    state.records = parsed.records;
    state.headers = parsed.headers;
    state.fileName = file.name;
    $("addRows").disabled = false;
    $("insertDirect").disabled = false;
    $("clearSelection").disabled = false;
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
  const ids = state.records.map((row) => row[identifier]).filter(Boolean);
  const total = state.records.reduce((sum, row) => sum + (Number(String(row[quantity] || "").replace(/[^0-9.]/g, "")) || 0), 0);
  const destination = is100km ? "BASE" : `BASE-${base.toUpperCase()}`;
  $("preview").textContent = `${state.fileName}\nDestino: ${destination}\nLinhas no arquivo: ${state.records.length}\n${is100km ? "SMs" : "Manifestos"}: ${ids.join(", ") || "não identificados"}\n${is100km ? "KM somados" : "Volumes somados"}: ${total}`;
  renderSampleRows();
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
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
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

function hideReview() {
  state.plan = null;
  $("review").classList.add("hidden");
  $("duplicateWarning").classList.add("hidden");
  $("duplicateText").textContent = "";
  $("reviewSummary").replaceChildren();
}

function clearSelection() {
  if (state.importing) return;
  $("inputFile").value = "";
  state.records = [];
  state.headers = [];
  state.fileName = "";
  hideReview();
  $("addRows").disabled = true;
  $("insertDirect").disabled = true;
  $("clearSelection").disabled = true;
  clearStatus();
  renderPreview();
}

function ensureFileMatchesBase(base) {
  if (base === "100km") {
    const required = ["SM", "DATA", "KM"];
    const missing = required.filter((header) => !state.headers.some((candidate) => normalizeHeader(candidate) === normalizeHeader(header)));
    if (missing.length) throw new Error(`Este arquivo não parece ser da planilha 100km. Não encontrei: ${missing.join(", ")}.`);
    return;
  }
  const expected = base === "recebendo" ? "DATA DE TRASFERENCIA" : "DT DE TRASFERENCIA";
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
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();
    if (!tables.items.length) throw new Error(`Não encontrei uma tabela na aba ${sheetName}.`);
    if (tables.items.length > 1) throw new Error(`Há mais de uma tabela na aba ${sheetName}; não vou escolher uma automaticamente.`);

    const table = tables.getItemAt(0);
    const headerRange = table.getHeaderRowRange();
    headerRange.load("values");
    const countResult = table.rows.getCount();
    await context.sync();
    const existingCount = countResult.value;
    const headers = headerRange.values[0].map((value) => String(value || "").trim());
    const columnMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
    const uniqueColumn = columnMap.get(normalizeHeader(inputKeyName));
    if (uniqueColumn === undefined) throw new Error(`Não encontrei a coluna ${inputKeyName} na tabela da aba ${sheetName}.`);
    const inputKeyHeader = state.headers.find((header) => normalizeHeader(header) === normalizeHeader(inputKeyName));
    if (!inputKeyHeader) throw new Error(`O arquivo não tem a coluna ${inputKeyName}.`);

    const ignored100kmHeaders = new Set(["MATRICULA", "TOTAL"]);
    const mappedInputColumns = new Map();
    for (const record of state.records) {
      if (!String(record[inputKeyHeader] || "").trim()) throw new Error(`Há uma linha sem ${inputKeyName}; revise o arquivo antes de inserir.`);
      for (const inputHeader of Object.keys(record)) {
        if (base === "100km" && ignored100kmHeaders.has(normalizeHeader(inputHeader))) continue;
        const targetColumn = columnMap.get(normalizeHeader(inputHeader));
        if (targetColumn !== undefined) mappedInputColumns.set(normalizeHeader(inputHeader), targetColumn);
      }
    }
    const existing = new Set();
    if (existingCount > 0) {
      const manifestValues = table.columns.getItemAt(uniqueColumn).getDataBodyRange();
      manifestValues.load("values");
      await context.sync();
      manifestValues.values.flat().map((value) => String(value || "").trim()).filter(Boolean).forEach((value) => existing.add(value));
    }
    const seenInFile = new Set();
    const duplicateList = [];
    const newRecordIndexes = [];
    state.records.forEach((record, index) => {
      const number = String(record[inputKeyHeader] || "").trim();
      if (existing.has(number) || seenInFile.has(number)) duplicateList.push(number);
      else newRecordIndexes.push(index);
      seenInFile.add(number);
    });

    const bodyRange = existingCount ? table.getDataBodyRange() : null;
    let bodyValues = [], bodyFormulas = [];
    if (bodyRange) {
      bodyRange.load("values,formulas");
      await context.sync();
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
        if (["DATA DE TRASFERENCIA", "DT DE TRASFERENCIA", "DATA"].includes(normalizeHeader(inputHeader))) value = excelDateSerial(rawValue);
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
      await context.sync();
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
    if (alignmentSources.size) await context.sync();
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
          const cellValue = applyColumnTextStyle(value, bodyValues[formatTemplateRow]?.[col], headers[col]);
          const alignmentSource = alignmentSources.get(`${assignment.row}:${col}`);
          if (alignmentSource) target.format.horizontalAlignment = alignmentSource.format.horizontalAlignment;
          if (normalizeHeader(headers[col]) === "CPF") {
            target.numberFormat = [["@"]];
            target.format.horizontalAlignment = "Left";
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
    await context.sync();
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
