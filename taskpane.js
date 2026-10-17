/* global Office, Excel */
"use strict";

const state = { records: [], headers: [], fileName: "", importing: false };
const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  $("inputFile").addEventListener("change", readSelectedFile);
  $("base").addEventListener("change", renderPreview);
  $("addRows").addEventListener("click", insertRows);
  $("clearSelection").addEventListener("click", clearSelection);
  $("confirmDuplicates").addEventListener("click", () => insertRows(true));
  $("cancelDuplicates").addEventListener("click", () => { hideDuplicateWarning(); clearStatus(); });
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
  const records = rows.slice(1).map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])));
  return { headers, records };
}

async function readSelectedFile() {
  clearStatus(); hideDuplicateWarning();
  const file = $("inputFile").files[0];
  state.records = []; state.headers = []; state.fileName = "";
  $("addRows").disabled = true; $("clearSelection").disabled = !file;
  if (!file) { renderPreview(); return; }
  try {
    const parsed = parseTSV(await file.text());
    state.records = parsed.records; state.headers = parsed.headers; state.fileName = file.name;
    $("addRows").disabled = false; $("clearSelection").disabled = false; renderPreview();
  } catch (error) {
    $("addRows").disabled = true; setStatus(error.message, true);
    $("preview").textContent = "Não foi possível ler o arquivo.";
  }
}

function renderPreview() {
  if (!state.records.length) { $("preview").textContent = "Selecione um arquivo para ver uma prévia."; return; }
  const base = $("base").value;
  const manifestos = state.records.map((row) => row.MANIFESTO).filter(Boolean);
  const totalVolumes = state.records.reduce((sum, row) => sum + (Number(row.VOLUMES) || 0), 0);
  $("preview").textContent = state.fileName + "\nDestino: BASE-" + base.toUpperCase() + "\nLinhas: " + state.records.length + "\nManifestos: " + (manifestos.join(", ") || "não identificados") + "\nVolumes somados: " + totalVolumes;
}

function normalizeHeader(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

function excelDateSerial(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return value;
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  return (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}

function showDuplicateWarning(manifestos) {
  $("duplicateText").textContent = manifestos.join(", ");
  $("duplicateWarning").classList.remove("hidden");
}
function hideDuplicateWarning() {
  $("duplicateWarning").classList.add("hidden"); $("duplicateText").textContent = "";
}
function clearSelection() {
  if (state.importing) return;
  $("inputFile").value = ""; state.records = []; state.headers = []; state.fileName = "";
  $("addRows").disabled = true; $("clearSelection").disabled = true;
  hideDuplicateWarning(); clearStatus(); renderPreview();
}

async function insertRows(confirmDuplicates = false) {
  if (state.importing) return;
  state.importing = true; $("addRows").disabled = true; clearStatus();
  if (confirmDuplicates) hideDuplicateWarning();
  try {
    const base = $("base").value;
    const expectedTransferHeader = base === "recebendo" ? "DATA DE TRASFERENCIA" : "DT DE TRASFERENCIA";
    if (!state.headers.some((header) => normalizeHeader(header) === normalizeHeader(expectedTransferHeader))) {
      throw new Error("O arquivo não corresponde à base " + (base === "recebendo" ? "RECEBENDO" : "EXPEDINDO") + ". Selecione o TSV dessa base.");
    }
    const sheetName = base === "recebendo" ? "BASE-RECEBENDO" : "BASE-EXPEDINDO";
    const result = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const tables = sheet.tables; tables.load("items/name"); await context.sync();
      if (!tables.items.length) throw new Error("Não encontrei uma tabela na aba " + sheetName + ".");
      if (tables.items.length > 1) throw new Error("Há mais de uma tabela na aba " + sheetName + "; não vou escolher uma automaticamente.");
      const table = tables.getItemAt(0);
      const headerRange = table.getHeaderRowRange(); headerRange.load("values");
      const existingCountResult = table.rows.getCount(); await context.sync();
      const existingCount = existingCountResult.value;
      const headers = headerRange.values[0].map((value) => String(value || "").trim());
      const columnMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
      const manifestColumn = columnMap.get("MANIFESTO");
      if (manifestColumn === undefined) throw new Error("Não encontrei a coluna MANIFESTO na tabela da aba " + sheetName + ".");
      const inputManifestHeader = state.headers.find((header) => normalizeHeader(header) === "MANIFESTO");
      if (!inputManifestHeader) throw new Error("O arquivo não tem a coluna MANIFESTO.");

      const mappedInputColumns = new Set();
      for (const record of state.records) {
        if (!String(record[inputManifestHeader] || "").trim()) throw new Error("Há uma linha sem número de manifesto; revise o arquivo antes de inserir.");
        for (const inputHeader of Object.keys(record)) {
          const targetColumn = columnMap.get(normalizeHeader(inputHeader));
          if (targetColumn !== undefined) mappedInputColumns.add(targetColumn);
        }
      }
      const inputValues = state.records.map((record) => {
        const values = new Map();
        for (const [inputHeader, rawValue] of Object.entries(record)) {
          const targetColumn = columnMap.get(normalizeHeader(inputHeader));
          if (targetColumn === undefined || rawValue === "" || rawValue == null) continue;
          let value = rawValue;
          const normalized = normalizeHeader(inputHeader);
          if (normalized === normalizeHeader("DATA DE TRASFERENCIA") || normalized === normalizeHeader("DT DE TRASFERENCIA")) value = excelDateSerial(rawValue);
          values.set(targetColumn, value);
        }
        return values;
      });

      const existing = new Set();
      if (existingCount > 0) {
        const manifestValues = table.columns.getItemAt(manifestColumn).getDataBodyRange();
        manifestValues.load("values"); await context.sync();
        manifestValues.values.flat().map((value) => String(value || "").trim()).filter(Boolean).forEach((value) => existing.add(value));
      }
      const duplicates = new Set(), duplicateManifestos = [];
      for (const record of state.records) {
        const number = String(record[inputManifestHeader] || "").trim();
        if (existing.has(number) || duplicates.has(number)) duplicateManifestos.push(number);
        duplicates.add(number);
      }
      const duplicateList = [...new Set(duplicateManifestos)];
      if (duplicateList.length && !confirmDuplicates) return { needsConfirmation: duplicateList };

      const bodyRange = existingCount ? table.getDataBodyRange() : null;
      let bodyValues = [], bodyFormulas = [];
      if (bodyRange) {
        bodyRange.load("values,formulas"); await context.sync();
        bodyValues = bodyRange.values; bodyFormulas = bodyRange.formulas;
      }
      const formulaColumns = new Set(), formulaTemplateRows = new Map();
      for (let row = 0; row < bodyFormulas.length; row += 1) {
        for (let col = 0; col < headers.length; col += 1) {
          const formula = bodyFormulas[row]?.[col];
          if (typeof formula === "string" && formula.startsWith("=")) {
            formulaColumns.add(col); formulaTemplateRows.set(col, row);
          }
        }
      }
      const reusableRows = [];
      for (let row = 0; row < existingCount; row += 1) {
        const hasIncomingData = [...mappedInputColumns].some((col) => {
          if (formulaColumns.has(col)) return false;
          const value = bodyValues[row]?.[col];
          return value !== null && value !== undefined && String(value).trim() !== "";
        });
        const manifestValue = bodyValues[row]?.[manifestColumn];
        if (!hasIncomingData && String(manifestValue ?? "").trim() === "") reusableRows.push(row);
      }
      const assignments = [];
      const reusedCount = Math.min(reusableRows.length, inputValues.length);
      for (let i = 0; i < reusedCount; i += 1) assignments.push({ row: reusableRows[i], values: inputValues[i], existing: true });
      const appendCount = inputValues.length - reusedCount;
      if (appendCount > 0) {
        const blankRows = Array.from({ length: appendCount }, () => new Array(headers.length).fill(""));
        const added = table.rows.add(-1, blankRows, true); added.load("index"); await context.sync();
        for (let i = 0; i < appendCount; i += 1) assignments.push({ row: added.index + i, values: inputValues[reusedCount + i], existing: false });
      }
      for (const assignment of assignments) {
        for (const [col, value] of assignment.values) {
          if (!formulaColumns.has(col)) table.getDataBodyRange().getCell(assignment.row, col).values = [[value]];
        }
        for (const col of formulaColumns) {
          const target = table.getDataBodyRange().getCell(assignment.row, col);
          const alreadyHasFormula = assignment.existing && typeof bodyFormulas[assignment.row]?.[col] === "string" && bodyFormulas[assignment.row][col].startsWith("=");
          const templateRow = formulaTemplateRows.get(col);
          if (!alreadyHasFormula && templateRow !== undefined && templateRow !== assignment.row) {
            target.copyFrom(table.getDataBodyRange().getCell(templateRow, col), Excel.RangeCopyType.formulas);
          }
        }
      }
      await context.sync();
      return { count: inputValues.length, reused: reusedCount };
    });
    if (result.needsConfirmation) { showDuplicateWarning(result.needsConfirmation); return; }
    const baseName = base === "recebendo" ? "BASE-RECEBENDO" : "BASE-EXPEDINDO";
    const reusedText = result.reused ? " " + result.reused + " linha(s) vazia(s) existente(s) foram reutilizadas; fórmulas preservadas." : " Fórmulas existentes da tabela foram mantidas e copiadas quando disponíveis.";
    setStatus(result.count + " linha(s) preenchida(s) na " + baseName + "." + reusedText, false, true);
  } catch (error) {
    setStatus(error.message || "Não foi possível adicionar as linhas.", true);
  } finally {
    state.importing = false; $("addRows").disabled = !state.records.length;
    $("clearSelection").disabled = !$("inputFile").files.length;
  }
}

function setStatus(message, isError, isSuccess = false) {
  const status = $("status"); status.textContent = message;
  status.className = isError ? "error" : (isSuccess ? "success" : "");
}
function clearStatus() { setStatus("", false); }
