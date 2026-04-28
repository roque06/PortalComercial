import * as XLSXNamespace from "xlsx";
import { execFileSync } from "node:child_process";
import path from "node:path";

const XLSX: any = (XLSXNamespace as any).default ?? XLSXNamespace;

export type RegistroExcel = {
  identificacion: string;
  tipoCuenta: string;
  relacionado: boolean;
  cedRelacionado: string;
  clienteNuevo?: boolean;
  tasaExepcion?: string;
  valorTasa?: string;
};

export type LeerRegistrosExcelOpts = {
  soloClienteNuevoSI?: boolean;
};

function normKey(s: any) {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function onlyDigits(v: string) {
  return String(v ?? "").replace(/\D/g, "");
}

export function leerRegistrosDesdeExcel(
  ruta: string,
  nombreHoja?: string,
  opts: LeerRegistrosExcelOpts = {}
): RegistroExcel[] {
  const workbook = XLSX.readFile(ruta);
  const hojaObjetivo = nombreHoja?.trim() || workbook.SheetNames[0];
  const hoja = workbook.Sheets[hojaObjetivo];
  if (!hoja) {
    throw new Error(
      `No se encontro la hoja '${hojaObjetivo}' en '${ruta}'. Hojas disponibles: ${workbook.SheetNames.join(", ")}`
    );
  }
  const datos = XLSX.utils.sheet_to_json(hoja, { defval: "" }) as any[];

  console.log("Headers detectados:", Object.keys(datos?.[0] ?? {}));

  const registros = datos
    .map((fila: any) => {
      const f: Record<string, string> = {};
      for (const [k, v] of Object.entries(fila)) {
        f[normKey(k)] = String(v ?? "").trim();
      }

      const identificacion = f["identificacion"] || f["cedula"] || "";
      const tipoCuenta = f["tipo cuenta"] || f["tipocuenta"] || f["tipo_cuenta"] || "";
      const relacionadoRaw = f["relacionado"] || "";
      const cedRelacionado =
        f["ced_relacionado"] || f["ced relacionado"] || f["cedrelacionado"] || "";
      const clienteNuevoRaw =
        f["cliente nuevo"] || f["cliente_nuevo"] || f["clientenuevo"] || "";
      const tasaExepcion =
        f["tasa exepcion"] ||
        f["tasa excepcion"] ||
        f["tasa_exepcion"] ||
        f["tasa_excepcion"] ||
        "";
      const valorTasa =
        f["valor tasa"] ||
        f["valor_tasa"] ||
        f["valortasa"] ||
        "";

      const relacionado = /^(si|sí|true|1)$/i.test(String(relacionadoRaw).trim());
      const clienteNuevo = /^(si|sí|true|1)$/i.test(String(clienteNuevoRaw).trim());

      return {
        identificacion,
        tipoCuenta,
        relacionado,
        cedRelacionado,
        clienteNuevo,
        tasaExepcion,
        valorTasa,
      };
    })
    .filter((r) => r.identificacion && r.tipoCuenta);

  if (!opts.soloClienteNuevoSI) return registros;

  const filtrados = registros.filter((r) => Boolean(r.clienteNuevo));
  if (registros.length > 0 && filtrados.length === 0) {
    throw new Error(
      `Se solicito filtrar por Cliente_Nuevo=SI en hoja '${hojaObjetivo}', pero no se encontro ninguna fila valida.`
    );
  }
  return filtrados;
}

export type MarcarCedulaProcesadaOpts = {
  hojaClientesNuevos?: string;
  hojaCedulasNuevas?: string;
  colorHex?: string; // ARGB: FFF4B183
};

function runExcelMarkScript(
  rutaAbs: string,
  hojaClientesNuevos: string,
  hojaCedulasNuevas: string,
  colorIndex: string,
  extraEnv: Record<string, string>
) {
  const ps = String.raw`
$ErrorActionPreference = 'Stop'
$ruta = $env:EXCEL_RUTA
$hojaClientes = $env:HOJA_CLIENTES
$hojaCedulas = $env:HOJA_CEDULAS
$colorIndex = [int]$env:COLOR_INDEX

function Normalize-Key([string]$s) {
  if ($null -eq $s) { return '' }
  return (($s.ToLower().Replace('_',' ')) -replace '\s+', ' ').Trim()
}
function Digits([string]$s) {
  if ($null -eq $s) { return '' }
  return ($s -replace '\D', '')
}
function Norm11([string]$s) {
  $d = Digits $s
  if ([string]::IsNullOrWhiteSpace($d)) { return '' }
  return $d.PadLeft(11, '0')
}
function Find-Col($ws, [string[]]$keys) {
  $usedCols = [Math]::Max(1, $ws.UsedRange.Columns.Count)
  for ($c = 1; $c -le $usedCols; $c++) {
    $k = Normalize-Key ([string]$ws.Cells.Item(1, $c).Text)
    if ($keys -contains $k) { return $c }
  }
  return 0
}
function Get-Sheet($wb, [string]$name) {
  $n = [string]$name
  foreach ($ws in $wb.Worksheets) {
    if ([string]$ws.Name -ieq $n) { return $ws }
  }
  if ($n.EndsWith('s')) {
    $alt = $n.Substring(0, $n.Length - 1)
    foreach ($ws in $wb.Worksheets) {
      if ([string]$ws.Name -ieq $alt) { return $ws }
    }
  } else {
    $alt = $n + 's'
    foreach ($ws in $wb.Worksheets) {
      if ([string]$ws.Name -ieq $alt) { return $ws }
    }
  }
  return $null
}

$targets = New-Object 'System.Collections.Generic.HashSet[string]'
$singleTarget = [string]$env:TARGET_DIGITS
if (-not [string]::IsNullOrWhiteSpace($singleTarget)) {
  $targets.Add((Norm11 $singleTarget)) | Out-Null
}

$targetsJson = [string]$env:TARGET_DIGITS_JSON
if (-not [string]::IsNullOrWhiteSpace($targetsJson)) {
  $parsed = ConvertFrom-Json $targetsJson
  foreach ($item in $parsed) {
    $norm = Norm11 ([string]$item)
    if (-not [string]::IsNullOrWhiteSpace($norm)) {
      $targets.Add($norm) | Out-Null
    }
  }
}

$excel = $null
$wb = $null
$updated = 0
$painted = 0

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($ruta)

  $wsClientes = Get-Sheet $wb $hojaClientes
  if ($null -ne $wsClientes) {
    $colId = Find-Col $wsClientes @('identificacion','cedula')
    $colCli = Find-Col $wsClientes @('cliente nuevo','clientenuevo')
    if ($colId -gt 0 -and $colCli -gt 0) {
      $lastRow = [Math]::Max(2, $wsClientes.UsedRange.Rows.Count + $wsClientes.UsedRange.Row - 1)
      for ($r = 2; $r -le $lastRow; $r++) {
        $val = [string]$wsClientes.Cells.Item($r, $colId).Text
        $norm = Norm11 $val
        if (-not [string]::IsNullOrWhiteSpace($norm) -and $targets.Contains($norm)) {
          $wsClientes.Cells.Item($r, $colCli).Value2 = 'NO'
          $updated++
        }
      }
    }
  }

  $wsCed = Get-Sheet $wb $hojaCedulas
  if ($null -ne $wsCed) {
    $colCed = Find-Col $wsCed @('cedula original','cedula','identificacion')
    if ($colCed -gt 0) {
      $lastRow2 = [Math]::Max(2, $wsCed.UsedRange.Rows.Count + $wsCed.UsedRange.Row - 1)
      for ($r = 2; $r -le $lastRow2; $r++) {
        $val2 = [string]$wsCed.Cells.Item($r, $colCed).Text
        $norm2 = Norm11 $val2
        if (-not [string]::IsNullOrWhiteSpace($norm2) -and $targets.Contains($norm2)) {
          $wsCed.Cells.Item($r, $colCed).Interior.ColorIndex = $colorIndex
          $painted++
        }
      }
    }
  }

  $wb.Save()
} finally {
  if ($null -ne $wb) { $wb.Close($true) | Out-Null }
  if ($null -ne $excel) { $excel.Quit() | Out-Null }
  if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
  if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

$obj = @{ updated_cliente_nuevo = $updated; painted_cedula = $painted; targets = $targets.Count } | ConvertTo-Json -Compress
Write-Output $obj
`;

  const out = execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps], {
    encoding: "utf-8",
    env: {
      ...process.env,
      EXCEL_RUTA: rutaAbs,
      HOJA_CLIENTES: hojaClientesNuevos,
      HOJA_CEDULAS: hojaCedulasNuevas,
      COLOR_INDEX: colorIndex,
      ...extraEnv,
    },
  });

  const parsed = JSON.parse(String(out || "{}").trim() || "{}");
  return {
    updatedClienteNuevo: Number(parsed.updated_cliente_nuevo || 0),
    paintedCedula: Number(parsed.painted_cedula || 0),
    targets: Number(parsed.targets || 0),
  };
}

export function marcarCedulaProcesadaEnExcel(
  ruta: string,
  identificacion: string,
  opts: MarcarCedulaProcesadaOpts = {}
) {
  const rutaAbs = path.resolve(process.cwd(), ruta);
  const targetDigits = onlyDigits(identificacion);
  if (!targetDigits) {
    return { updatedClienteNuevo: 0, paintedCedula: 0 };
  }

  const hojaClientesNuevos = opts.hojaClientesNuevos ?? "Clientes_Nuevos";
  const hojaCedulasNuevas = opts.hojaCedulasNuevas ?? "Ced_Clientes_Nuevos";
  // Excel ColorIndex 6 = amarillo (resaltado visible)
  const colorIndex = "6";
  return runExcelMarkScript(rutaAbs, hojaClientesNuevos, hojaCedulasNuevas, colorIndex, {
    TARGET_DIGITS: targetDigits,
  });
}

export function marcarCedulasProcesadasEnExcel(
  ruta: string,
  identificaciones: string[],
  opts: MarcarCedulaProcesadaOpts = {}
) {
  const targets = Array.from(
    new Set(
      (identificaciones ?? [])
        .map((id) => onlyDigits(id))
        .filter(Boolean)
    )
  );
  if (!targets.length) {
    return { updatedClienteNuevo: 0, paintedCedula: 0, targets: 0 };
  }

  const rutaAbs = path.resolve(process.cwd(), ruta);
  const hojaClientesNuevos = opts.hojaClientesNuevos ?? "Clientes_Nuevos";
  const hojaCedulasNuevas = opts.hojaCedulasNuevas ?? "Ced_Clientes_Nuevos";
  const colorIndex = "6";

  return runExcelMarkScript(rutaAbs, hojaClientesNuevos, hojaCedulasNuevas, colorIndex, {
    TARGET_DIGITS_JSON: JSON.stringify(targets),
  });
}
