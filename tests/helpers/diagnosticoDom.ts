import type { Page } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';

const DIAG_DOM_ENABLED = process.env.DIAG_DOM !== '0';
const DIAGNOSTICOS_DIR = path.join(process.cwd(), 'artifacts', 'diagnosticos_dom');

/**
 * Sanitiza un nombre de archivo para evitar caracteres inválidos en Windows/Unix
 */
function sanitizarNombreArchivo(nombre: string): string {
  return String(nombre || 'sin_nombre')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .toLowerCase();
}

/**
 * Obtiene información sobre un elemento visible o el primero no visible
 */
async function obtenerDetalleElemento(page: Page, selector: string) {
  try {
    const elementos = await page.locator(selector).all().catch(() => []);
    const detalles = [];

    for (let i = 0; i < Math.min(elementos.length, 5); i++) {
      const el = elementos[i];
      const visible = await el.isVisible().catch(() => false);
      const box = await el.boundingBox().catch(() => null);
      const tagName = await el.evaluate(e => e.tagName).catch(() => '');
      const textContent = await el.evaluate(e => (e.textContent || '').substring(0, 100)).catch(() => '');
      const innerText = await el.evaluate(e => (e.innerText || '').substring(0, 100)).catch(() => '');
      const value = await el.inputValue().catch(() => '');
      const id = await el.getAttribute('id').catch(() => '');
      const className = await el.getAttribute('class').catch(() => '');
      const name = await el.getAttribute('name').catch(() => '');
      const type = await el.getAttribute('type').catch(() => '');
      const role = await el.getAttribute('role').catch(() => '');
      const ariaLabel = await el.getAttribute('aria-label').catch(() => '');
      const ariaControls = await el.getAttribute('aria-controls').catch(() => '');
      const ariaExpanded = await el.getAttribute('aria-expanded').catch(() => '');
      const ariaSelected = await el.getAttribute('aria-selected').catch(() => '');
      const disabled = await el.getAttribute('disabled').catch(() => '');
      const readonly = await el.getAttribute('readonly').catch(() => '');

      // Obtener data-* principales
      const dataAttrs = await el.evaluate((e) => {
        const attrs: Record<string, string> = {};
        Array.from(e.attributes || []).forEach((attr: any) => {
          if (attr.name.startsWith('data-')) {
            attrs[attr.name] = (attr.value || '').substring(0, 50);
          }
        });
        return attrs;
      }).catch(() => ({}));

      // Obtener display/visibility/opacity/pointerEvents
      const computedStyle = await el.evaluate(e => {
        const style = window.getComputedStyle(e);
        return {
          display: style.display,
          visibility: style.visibility,
          opacity: style.opacity,
          pointerEvents: style.pointerEvents,
        };
      }).catch(() => ({}));

      const outerHTML = await el.evaluate(e => e.outerHTML.substring(0, 200)).catch(() => '');

      // Redactar passwords
      let redactedValue = value;
      if (type === 'password' || className.includes('password')) {
        redactedValue = '[PASSWORD_REDACTED]';
      }

      detalles.push({
        indice: i,
        visible,
        selector,
        tagName,
        textContent,
        innerText,
        value: redactedValue,
        id: id || undefined,
        class: className || undefined,
        name: name || undefined,
        type: type || undefined,
        role: role || undefined,
        ariaLabel: ariaLabel || undefined,
        ariaControls: ariaControls || undefined,
        ariaExpanded: ariaExpanded || undefined,
        ariaSelected: ariaSelected || undefined,
        disabled: disabled || undefined,
        readonly: readonly || undefined,
        dataAttrs: Object.keys(dataAttrs).length > 0 ? dataAttrs : undefined,
        computedStyle,
        boundingBox: box,
        outerHTML,
      });
    }

    return {
      selector,
      totalCount: elementos.length,
      visibleCount: elementos.filter((_, i) => i < 5).length,
      elementos: detalles,
    };
  } catch (error) {
    return {
      selector,
      error: String(error),
    };
  }
}

/**
 * Captura diagnóstico completo del DOM
 */
export async function capturarDiagnosticoDom(
  page: Page,
  nombre: string,
  selectors: string[] = [],
  options?: {
    includeFullHtml?: boolean;
    maxHtmlLength?: number;
    motivo?: string;
  }
): Promise<void> {
  if (!DIAG_DOM_ENABLED) {
    return;
  }

  const nombreSanitizado = sanitizarNombreArchivo(nombre);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const baseFileName = `${timestamp}__${nombreSanitizado}`;

  try {
    // Asegurar directorio existe
    fs.mkdirSync(DIAGNOSTICOS_DIR, { recursive: true });

    console.log(`[DiagDOM] capturando nombre=${nombre} motivo=${options?.motivo || 'sin motivo'}`);

    // 1. Capturar screenshot
    const screenshotPath = path.join(DIAGNOSTICOS_DIR, `${baseFileName}.png`);
    try {
      await page.screenshot({ path: screenshotPath, fullPage: true });
      console.log(`[DiagDOM] screenshot=${screenshotPath}`);
    } catch (e) {
      console.log(`[DiagDOM][WARN] No se pudo capturar screenshot: ${String(e)}`);
    }

    // 2. Capturar texto visible
    const visibleTextPath = path.join(DIAGNOSTICOS_DIR, `${baseFileName}.visible_text.txt`);
    try {
      const visibleText = await page.evaluate(() => document.body.innerText).catch(() => '');
      fs.writeFileSync(visibleTextPath, visibleText, 'utf-8');
      console.log(`[DiagDOM] visibleText=${visibleTextPath}`);
    } catch (e) {
      console.log(`[DiagDOM][WARN] No se pudo capturar visibleText: ${String(e)}`);
    }

    // 3. Capturar DOM completo
    const domPath = path.join(DIAGNOSTICOS_DIR, `${baseFileName}.dom.html`);
    try {
      let html = await page.evaluate(() => document.body.outerHTML).catch(() => '');
      const maxLength = options?.maxHtmlLength ?? 1_500_000;

      if (html.length > maxLength) {
        const truncatedNote = `\n\n<!-- NOTA: HTML TRUNCADO. Original: ${html.length} bytes. Máximo: ${maxLength} bytes. -->\n`;
        html = html.substring(0, maxLength) + truncatedNote;
      }

      fs.writeFileSync(domPath, html, 'utf-8');
      console.log(`[DiagDOM] dom=${domPath}`);
    } catch (e) {
      console.log(`[DiagDOM][WARN] No se pudo capturar dom: ${String(e)}`);
    }

    // 4. Capturar elementos específicos
    const selectoresABuscar = selectors.length > 0 ? selectors : getDefaultSelectors();
    const elementosPath = path.join(DIAGNOSTICOS_DIR, `${baseFileName}.elements.json`);
    try {
      const elementosInfo = [];

      for (const selector of selectoresABuscar) {
        const info = await obtenerDetalleElemento(page, selector);
        elementosInfo.push(info);
      }

      fs.writeFileSync(elementosPath, JSON.stringify(elementosInfo, null, 2), 'utf-8');
      console.log(`[DiagDOM] elements=${elementosPath}`);
    } catch (e) {
      console.log(`[DiagDOM][WARN] No se pudo capturar elements: ${String(e)}`);
    }

    // 5. Capturar metadata
    const metaPath = path.join(DIAGNOSTICOS_DIR, `${baseFileName}.meta.json`);
    try {
      const url = page.url();
      const title = await page.title().catch(() => '');
      const viewport = page.viewportSize();

      const iframesCount = await page.locator('iframe').count().catch(() => 0);
      const dialogsCount = await page.locator('[role="dialog"], .modal, .p-dialog').count().catch(() => 0);

      const meta = {
        url,
        title,
        timestamp,
        motivo: options?.motivo || 'sin motivo',
        viewport,
        selectoresRecibidos: selectoresABuscar.length,
        selectoresUsados: selectoresABuscar,
        iframesCount,
        dialogsCount,
      };

      fs.writeFileSync(metaPath, JSON.stringify(meta, null, 2), 'utf-8');
      console.log(`[DiagDOM] meta=${metaPath}`);
    } catch (e) {
      console.log(`[DiagDOM][WARN] No se pudo capturar meta: ${String(e)}`);
    }

  } catch (e) {
    // El helper nunca debe romper el flujo
    console.log(`[DiagDOM][ERROR] Fallo capturando diagnóstico: ${String(e)}`);
  }
}

/**
 * Helper específico para Otras Coincidencias Bizagi
 */
export async function capturarDiagnosticoOtrasCoincidenciasBizagi(
  page: Page,
  nombre: string
): Promise<void> {
  const selectores = [
    'text=Otras Coincidencias',
    'text=Lexis Nexis',
    'text=Acción',
    'text=% Coinc.',
    'text=Falso Positivo',
    'input.ui-select-data.ui-selectmenu-value',
    'input[role="combobox"]',
    'div.ui-selectmenu',
    'td.ui-bizagi-grid-align-center',
    'ul[role="listbox"]',
    'li[role="option"]',
    '.ui-selectmenu-menu',
    '.ui-menu',
    '.ui-bizagi-grid',
    '.ui-bizagi-grid-row',
    '.ui-bizagi-grid-cell',
  ];

  await capturarDiagnosticoDom(page, nombre, selectores, {
    motivo: 'diagnostico_otras_coincidencias_bizagi',
  });
}

/**
 * Helper específico para PostProducto
 */
export async function capturarDiagnosticoPostProducto(
  page: Page,
  nombre: string
): Promise<void> {
  const selectores = [
    'text=Nivel de estudio',
    'text=Reside actualmente',
    'text=Correspondencia',
    'text=Tipo de envío',
    'text=Reintentar buscar lista',
    'text=Es requerido',
    'input[role="combobox"]',
    '.p-dropdown',
    '.p-dropdown-label',
    'button',
    '[role="dialog"]',
    '.p-error',
  ];

  await capturarDiagnosticoDom(page, nombre, selectores, {
    motivo: 'diagnostico_post_producto',
  });
}

/**
 * Helper específico para Taller
 */
export async function capturarDiagnosticoTaller(
  page: Page,
  nombre: string
): Promise<void> {
  const selectores = [
    'text=Propósito',
    'text=Origen de fondos',
    'text=Instrumento',
    'text=Tipo de correspondencia',
    'text=Es requerido',
    'input[role="combobox"]',
    '.p-dropdown',
    '.p-dropdown-label',
    'button',
  ];

  await capturarDiagnosticoDom(page, nombre, selectores, {
    motivo: 'diagnostico_taller',
  });
}

/**
 * Helper específico para Nueva Solicitud
 */
export async function capturarDiagnosticoNuevaSolicitud(
  page: Page,
  nombre: string
): Promise<void> {
  const selectores = [
    'text=Solicitud Nueva',
    'text=Información básica',
    'text=Verificar solicitante',
    'text=Oficial',
    'text=Promotor',
    'text=Número de identificación',
    'text=Depurar',
    'input',
    'button',
    '[role="combobox"]',
    '.p-dropdown',
  ];

  await capturarDiagnosticoDom(page, nombre, selectores, {
    motivo: 'diagnostico_nueva_solicitud',
  });
}

/**
 * Selectors estándar por defecto
 */
function getDefaultSelectors(): string[] {
  return [
    'button',
    'input',
    'select',
    'textarea',
    '[role="combobox"]',
    '[role="listbox"]',
    '[role="option"]',
    '.p-dropdown',
    '.p-dropdown-label',
    '.ui-selectmenu',
    '.ui-select-data',
    '.ui-selectmenu-value',
    '.ui-bizagi-grid',
    '.ui-bizagi-grid-row',
    '.ui-bizagi-grid-cell',
    '.ui-bizagi-grid-align-center',
    '.modal',
    '[role="dialog"]',
    '.p-dialog',
    '.toast',
    '.p-toast',
    '.alert',
    '.p-message',
    '.p-error',
    'text=Otras Coincidencias',
    'text=Lexis Nexis',
    'text=Falso Positivo',
    'text=Reintentar buscar lista',
    'text=Actualizando solicitud',
    'text=Es requerido',
  ];
}

/**
 * Watchdog para detectar pantallas enganchadas
 */
export async function esperarCambioODiagnosticar(
  page: Page,
  nombre: string,
  accion: () => Promise<void>,
  opciones?: {
    timeoutMs?: number;
    descripcion?: string;
    urlDebeCambiar?: boolean;
    textoEsperado?: RegExp;
    selectorsDiag?: string[];
  }
): Promise<void> {
  if (!DIAG_DOM_ENABLED) {
    await accion();
    return;
  }

  const timeoutMs = opciones?.timeoutMs ?? 8000;
  const descripcion = opciones?.descripcion ?? 'sin descripción';
  const urlDebeCambiar = opciones?.urlDebeCambiar ?? true;
  const textoEsperado = opciones?.textoEsperado;
  const selectorsDiag = opciones?.selectorsDiag;

  // Captura estado inicial
  const urlAntes = page.url();
  const bodyTextAntes = await page.evaluate(() =>
    (document.body.innerText || '').toLowerCase().replace(/\s+/g, ' ')
  ).catch(() => '');

  try {
    // Ejecuta la acción
    await accion();

    // Espera el tiempo especificado
    await page.waitForTimeout(timeoutMs);

    // Verifica cambios
    const urlDespues = page.url();
    const bodyTextDespues = await page.evaluate(() =>
      (document.body.innerText || '').toLowerCase().replace(/\s+/g, ' ')
    ).catch(() => '');

    const urlCambio = urlAntes !== urlDespues;
    const textoCambio = bodyTextAntes !== bodyTextDespues;
    const textoPresenteCheck = textoEsperado ? textoEsperado.test(bodyTextDespues) : true;

    // Determina si está enganchado
    const enganchado = (urlDebeCambiar && !urlCambio && !textoCambio) ||
                       (!textoPresenteCheck);

    if (enganchado) {
      console.log(`[Watchdog][STUCK] nombre=${nombre} descripcion=${descripcion}`);
      await capturarDiagnosticoDom(page, `stuck_${nombre}`, selectorsDiag, {
        motivo: `watchdog_enganchado_${descripcion}`,
      });
    }
  } catch (error) {
    // En caso de timeout o error, capturar diagnóstico
    console.log(`[Watchdog][ERROR] nombre=${nombre} error=${String(error)}`);
    await capturarDiagnosticoDom(page, `error_${nombre}`, selectorsDiag, {
      motivo: `watchdog_error_${descripcion}`,
    });
  }
}
