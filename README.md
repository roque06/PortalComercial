# PortalComercial (Playwright + TypeScript)

Proyecto de automatización E2E con Playwright. Se reorganizó para separar **código editable** de **artefactos generados** y reducir el contexto que consume Claude Code.

## Estructura

```
.
├─ tests/
│  ├─ e2e/                     # Specs principales
│  │  └─ ce-e/                 # Cuenta efectivo / certificados
│  └─ helpers/                 # Helpers compartidos (UI, Bizagi, Excel, etc.)
├─ scripts/                    # Runners/GUI y utilidades (p.ej. setupLogin)
├─ data/                       # Insumos de prueba (Excel, documentos)
├─ artifacts/                  # Salidas locales (no versionar)
│  ├─ evidencias_tmp/          # Screenshots PNG (evidencia)
│  ├─ evidencias/              # PDFs generados desde PNGs
│  ├─ outputs/                 # Logs del runner
│  ├─ playwright-report/       # Reporte HTML
│  ├─ test-results/            # Traces/screenshots/videos de Playwright
│  └─ profiles/                # Perfiles de navegador (muy pesado)
└─ playwright.config.js
```

## Tests principales

- `tests/e2e/ce-e/ce-ex.spec.ts`
- `tests/e2e/ce-e/ce-new.spec.ts`
- `tests/e2e/ce-e/Certificado-ex.spec.ts`

## Evidencias y reportes

- PNGs: `artifacts/evidencias_tmp/` (configurable con `PW_EVIDENCIAS_DIR`)
- PDFs: `artifacts/evidencias/` (configurable con `PW_EVIDENCIAS_PDF_DIR`)
- HTML report: `artifacts/playwright-report/`
- Playwright output: `artifacts/test-results/`

## Claude Code (contexto mínimo)

El archivo `.claudeignore` excluye carpetas pesadas (por ejemplo `node_modules/`, `artifacts/`, `test-results/`, perfiles, imágenes/PDFs/logs).

Sugerencia: abre Claude Code apuntando a una subcarpeta útil:

- Solo tests: `tests/`
- Solo specs: `tests/e2e/`
- Helpers: `tests/helpers/`

## Runner

Ver `scripts/README_RUNNER.md`.
