/* global Office, Excel */
"use strict";

const state = { records: [], headers: [], fileName: "" };
const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  $("inputFile").addEventListener("change", readSelectedFile);
  $("base").addEventListener("change", renderPreview);
  $("addRows").addEventListener("click", insertRows);
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
  const file = $("inputFile").files[0];
  state.records = [];
  state.headers = [];
  if (!file) { renderPreview(); return; }
  try {
    const parsed = parseTSV(await file.text());
    state.records = parsed.records;
    state.headers = parsed.headers;
    state.fileName = file.name;
    $("addRows").disabled = false;
    renderPreview();
  } catch (error) {
    $("addRows").disabled = true;
    setStatus(error.message, true);
    $("preview").textContent = "Não foi possível ler o arquivo.";
  }
}

function renderPreview() {
  if (!state.records.length) return;
  const base = $("base").value;
  const manifestoKey = "MANIFESTO";
  const volumesKey = "VOLUMES";
  const manifestos = state.records.map((row) => row[manifestoKey]).filter(Boolean);
  const totalVolumes = state.records.reduce((sum, row) => sum + (Number(row[volumesKey]) || 0), 0);
  $("preview").textContent = `${state.fileName}\nDestino: BASE-${base.toUpperCase()}\nLinhas: ${state.records.length}\nManifestos: ${manifestos.join(", ") || "não identificados"}\nVolumes somados: ${totalVolumes}`;
}

function normalizeHeader(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
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
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return value;
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  return (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}

async function insertRows() {
  $("addRows").disabled = true;
  clearStatus();
  try {
    const base = $("base").value;
    const sheetName = base === "recebendo" ? "BASE-RECEBENDO" : "BASE-EXPEDINDO";
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const tables = sheet.tables;
      tables.load("items/name");
      await context.sync();
      if (!tables.items.length) throw new Error(`Não encontrei uma tabela na aba ${sheetName}.`);
      if (tables.items.length > 1) throw new Error(`Há mais de uma tabela na aba ${sheetName}; não vou escolher uma automaticamente.`);

      const table = tables.getItemAt(0);
      const headerRange = table.getHeaderRowRange();
      headerRange.load("values");
      const fullRange = table.getRange();
      fullRange.load("rowIndex");
      const existingCountResult = table.rows.getCount();
      await context.sync();
      const existingCount = existingCountResult.value;

      const headers = headerRange.values[0].map((value) => String(value || "").trim());
      const columnMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
      const manifestColumn = columnMap.get("MANIFESTO");
      if (manifestColumn === undefined) throw new Error(`Não encontrei a coluna MANIFESTO na tabela da aba ${sheetName}.`);
      const inputManifestHeader = state.headers.find((header) => normalizeHeader(header) === "MANIFESTO");
      if (!inputManifestHeader) throw new Error("O arquivo não tem a coluna MANIFESTO.");

      const formulas = [];
      const data = [];
      const agingColumn = base === "recebendo" ? columnMap.get("AGING") : undefined;
      const transferHeader = base === "recebendo" ? "DATA DE TRASFERENCIA" : "DT DE TRASFERENCIA";
      const transferColumn = base === "recebendo" ? columnMap.get(normalizeHeader(transferHeader)) : undefined;
      for (const record of state.records) {
        const blankRow = new Array(headers.length).fill("");
        data.push(blankRow);
        const manifest = String(record[inputManifestHeader] || "").trim();
        if (!manifest) throw new Error("Há uma linha sem número de manifesto; revise o arquivo antes de inserir.");
        for (const [inputHeader, rawValue] of Object.entries(record)) {
          const targetColumn = columnMap.get(normalizeHeader(inputHeader));
          if (targetColumn === undefined || rawValue === "" || rawValue == null) continue;
          if (base === "recebendo" && targetColumn === agingColumn) continue;
          let value = rawValue;
          if (normalizeHeader(inputHeader) === normalizeHeader(transferHeader)) value = excelDateSerial(rawValue);
          blankRow[targetColumn] = value;
        }
        if (base === "recebendo" && agingColumn !== undefined && transferColumn !== undefined) {
          formulas.push({ offset: data.length - 1, column: agingColumn, transferColumn });
        }
      }

      const duplicates = new Set();
      const existing = new Set();
      if (existingCount > 0) {
        const manifestValues = table.columns.getItemAt(manifestColumn).getDataBodyRange();
        manifestValues.load("values");
        await context.sync();
        manifestValues.values.flat().map((value) => String(value || "").trim()).filter(Boolean)
          .forEach((value) => existing.add(value));
      }
      for (const record of state.records) {
        const number = String(record[inputManifestHeader] || "").trim();
        if (existing.has(number) || duplicates.has(number)) {
          throw new Error(`O manifesto ${number} já aparece na tabela ou está repetido no arquivo. Nenhuma linha foi inserida.`);
        }
        duplicates.add(number);
      }

      const firstDataRow = fullRange.rowIndex + 1 + existingCount;
      table.rows.add(-1, data, true);
      await context.sync();
      for (const item of formulas) {
        const excelRow = firstDataRow + item.offset + 1;
        const transferCell = `${columnLetter(item.transferColumn)}${excelRow}`;
        const agingCell = sheet.getCell(firstDataRow + item.offset, item.column);
        agingCell.formulas = [[`=IF(${transferCell}="","",IF(TODAY()-${transferCell}>=6,"D5+","D"&MAX(0,TODAY()-${transferCell})))`]];
      }
      await context.sync();
    });
    setStatus(`${state.records.length} linha(s) adicionada(s) à ${base === "recebendo" ? "BASE-RECEBENDO" : "BASE-EXPEDINDO"}.`, false, true);
  } catch (error) {
    setStatus(error.message || "Não foi possível adicionar as linhas.", true);
  } finally {
    $("addRows").disabled = !state.records.length;
  }
}

function setStatus(message, isError, isSuccess = false) {
  const status = $("status");
  status.textContent = message;
  status.className = isError ? "error" : (isSuccess ? "success" : "");
}

function clearStatus() { setStatus("", false); }
