import fs from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import process from "node:process";
import xlsx from "xlsx";

const DEFAULT_DB_CONNECTIONS = "C:\\CierreBancarioQA\\src\\config\\dbConnections.json";
const DEFAULT_SQLPLUS = "C:\\instantclient_21_13\\sqlplus.exe";

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) {
      continue;
    }
    const key = token.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith("--")) {
      args[key] = "true";
      continue;
    }
    args[key] = next;
    i += 1;
  }
  return args;
}

function fail(message) {
  console.error(message);
  process.exit(1);
}

function ensureFileExists(filePath, label) {
  if (!filePath || !fs.existsSync(filePath)) {
    fail(`${label} no existe: ${filePath || "(vacío)"}`);
  }
}

function normalizeCedula(value) {
  return String(value ?? "").replace(/\D/g, "").trim();
}

function buildOutputPath(inputPath, instance) {
  const parsed = path.parse(inputPath);
  return path.join(
    parsed.dir,
    `${parsed.name}_no_encontradas_${instance.toUpperCase()}${parsed.ext || ".xlsx"}`
  );
}

function loadConnections(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  const parsed = JSON.parse(raw);
  if (!Array.isArray(parsed)) {
    fail("El archivo dbConnections.json no tiene el formato esperado.");
  }
  return parsed;
}

function getConnectionString(connections, instance) {
  const target = connections.find(
    (item) => String(item?.name || "").toUpperCase() === instance.toUpperCase()
  );
  if (!target?.connectString) {
    fail(`No encontré la instancia ${instance} en dbConnections.json.`);
  }
  return target.connectString;
}

function collectCedulasFromWorkbook(inputPath) {
  const workbook = xlsx.readFile(inputPath, { raw: false, cellText: true });
  const rows = [];
  const seen = new Set();

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const matrix = xlsx.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      blankrows: false,
      defval: ""
    });
    if (!matrix.length) {
      continue;
    }

    const headers = matrix[0].map((value) => String(value ?? "").trim().toLowerCase());
    const cedulaIndex = headers.findIndex((header) => header === "cedula");
    if (cedulaIndex < 0) {
      continue;
    }

    for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
      const rawCedula = matrix[rowIndex]?.[cedulaIndex] ?? "";
      const cedula = String(rawCedula).trim();
      const cedulaNormalizada = normalizeCedula(rawCedula);
      if (!cedulaNormalizada) {
        continue;
      }

      const key = `${sheetName}|${cedula}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);

      rows.push({
        hoja: sheetName,
        fila_excel: rowIndex + 1,
        cedula,
        cedula_normalizada: cedulaNormalizada
      });
    }
  }

  return rows;
}

function fetchCedulasFromOracle({ sqlplusPath, username, password, connectionString }) {
  const sql = `
set heading off
set feedback off
set pagesize 0
set linesize 32767
set trimspool on
select regexp_replace(trim(NUMERO_IDENTIFICACION), '[^0-9]', '')
from CLIENTES_B2000
where CODIGO_TIPO_IDENTIFICACION = 1;
exit
`;

  return new Promise((resolve, reject) => {
    const connectionArg = `${username}/${password}@${connectionString}`;
    const child = spawn(sqlplusPath, ["-S", connectionArg], {
      windowsHide: true
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    child.on("error", (error) => {
      reject(error);
    });

    child.on("close", (code) => {
      if (code !== 0) {
        reject(
          new Error(
            `sqlplus finalizó con código ${code}.${stderr ? ` Detalle: ${stderr.trim()}` : ""}`
          )
        );
        return;
      }

      const oracleErrors = stdout
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => /^ORA-\d+|^SP2-\d+/i.test(line));

      if (oracleErrors.length) {
        reject(new Error(`Oracle devolvió error: ${oracleErrors.join(" | ")}`));
        return;
      }

      const values = new Set(
        stdout
          .split(/\r?\n/)
          .map((line) => normalizeCedula(line))
          .filter(Boolean)
      );

      resolve(values);
    });

    child.stdin.write(sql);
    child.stdin.end();
  });
}

function writeWorkbook({ inputPath, outputPath, instance, missingRows, inputRows, dbCedulas }) {
  const workbook = xlsx.utils.book_new();

  const missingSheet = missingRows.map((item) => ({
    Hoja: item.hoja,
    FilaExcel: item.fila_excel,
    Cedula: item.cedula,
    CedulaNormalizada: item.cedula_normalizada
  }));

  const summarySheet = [
    { Metric: "Archivo origen", Value: inputPath },
    { Metric: "Instancia", Value: instance.toUpperCase() },
    { Metric: "Total filas con Cedula", Value: inputRows.length },
    { Metric: "Cedulas no encontradas", Value: missingRows.length },
    { Metric: "Cedulas encontradas", Value: inputRows.length - missingRows.length },
    { Metric: "Total identificaciones en DB tipo 1", Value: dbCedulas.size }
  ];

  const noExistenWorksheet = xlsx.utils.json_to_sheet(missingSheet);
  const resumenWorksheet = xlsx.utils.json_to_sheet(summarySheet);

  xlsx.utils.book_append_sheet(workbook, noExistenWorksheet, "No_Existen");
  xlsx.utils.book_append_sheet(workbook, resumenWorksheet, "Resumen");
  xlsx.writeFile(workbook, outputPath);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const inputPath = args.input;
  const instance = args.instance;
  const username = args.username || "system";
  const password = args.password || "system";
  const dbConnectionsPath = args["db-connections"] || DEFAULT_DB_CONNECTIONS;
  const sqlplusPath = args.sqlplus || DEFAULT_SQLPLUS;
  const outputPath = args.output || buildOutputPath(inputPath, instance || "QA1");

  if (!inputPath) {
    fail("Debes indicar --input con la ruta del Excel.");
  }
  if (!instance) {
    fail("Debes indicar --instance con la instancia Oracle, por ejemplo QA1.");
  }

  ensureFileExists(inputPath, "El archivo Excel");
  ensureFileExists(dbConnectionsPath, "El archivo dbConnections.json");
  ensureFileExists(sqlplusPath, "sqlplus");

  console.log(`Leyendo conexiones desde: ${dbConnectionsPath}`);
  const connections = loadConnections(dbConnectionsPath);
  const connectionString = getConnectionString(connections, instance);

  console.log(`Leyendo Excel: ${inputPath}`);
  const inputRows = collectCedulasFromWorkbook(inputPath);
  if (!inputRows.length) {
    fail("No encontré ninguna columna 'Cedula' en el Excel o no tenía valores.");
  }
  console.log(`Cédulas leídas del Excel: ${inputRows.length}`);

  console.log(`Consultando Oracle en instancia ${instance.toUpperCase()}...`);
  const dbCedulas = await fetchCedulasFromOracle({
    sqlplusPath,
    username,
    password,
    connectionString
  });
  console.log(`Cédulas recuperadas desde Oracle: ${dbCedulas.size}`);

  const missingRows = inputRows.filter((item) => !dbCedulas.has(item.cedula_normalizada));
  console.log(`Cédulas no encontradas: ${missingRows.length}`);

  writeWorkbook({
    inputPath,
    outputPath,
    instance,
    missingRows,
    inputRows,
    dbCedulas
  });

  console.log(`Excel generado: ${outputPath}`);
}

main().catch((error) => {
  console.error(error?.message || String(error));
  process.exit(1);
});
