import { test, type Locator, type Page } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import {
  seleccionarDefaultSiVacio,
  llenarInputMask,
  llenarInputNumber,
  asegurarIdentificacionHabilitada,
  clickReintentarListaSiVisible,
  llenarCampoPorLabel,
  seleccionarDropdownFiltrableConReintentar,
  capturarCuentaComoPNG,
  unirPNGsEnUnPDF,
  llenarFechaSiVisibleYVacia,
  validarCorreoPredeterminadoYCorregir,
} from '../../helpers/uiHelpers';
import {
  leerRegistrosDesdeExcel,
  cancelarCasoEnBizagiDesdePortal,
  extraerCasoActivoMpn,
  seleccionarInstrumentoRobusto,
  seleccionarDropdownPorCampo,
  esperarYClickReintentarPaisIdentificacion,
  validarApnfdYSeleccionarNoSiVacio,
  validarGestionDocumentalSiRequerido,
  cargarDocumentoEnGestionDocumental,
  abrirBpmSiVerificacionConoceCliente,
  esperarPortalListoTrasLogin,
  agregarRelacionadoSiAplica,
  seleccionarProductoCuentaEfectivoExistente,
  type RegistroExcel,
} from '../../helpers/productos/cuenta-efectivo/existente/helpers';
import {
  cerrarModalFinalizadaRapido as cerrarModalFinalizadaRapidoShared,
  cerrarModalSolicitudFinalizada as cerrarModalSolicitudFinalizadaShared,
  clickBotonFinalizarRobusto as clickBotonFinalizarRobustoShared,
  confirmarFinalizacionSiHayPopup as confirmarFinalizacionSiHayPopupShared,
  estadoPendienteFinalizarVisible as estadoPendienteFinalizarVisibleShared,
  estadoSolicitudFinalizadaVisible as estadoSolicitudFinalizadaVisibleShared,
  finalizarSolicitudSiPendiente as finalizarSolicitudSiPendienteShared,
  irAPasoFinalDesdeWizard as irAPasoFinalDesdeWizardShared,
  modalSolicitudFinalizadaLocator as modalSolicitudFinalizadaLocatorShared,
} from '../../helpers/portal/finalizacion';
import {
  confirmarSeleccionProductoRapida as confirmarSeleccionProductoRapidaShared,
  detectarProductoAgregadoEnUI as detectarProductoAgregadoEnUIShared,
  esValorDropdownVacio as esValorDropdownVacioShared,
  escapeRegexText as escapeRegexTextShared,
  extraerCodigoProducto as extraerCodigoProductoShared,
  extraerNombreProducto as extraerNombreProductoShared,
  leerValorDropdownEnScope as leerValorDropdownEnScopeShared,
  localizarSeccionProductos as localizarSeccionProductosShared,
  modalProductoConfigVisible as modalProductoConfigVisibleShared,
  seleccionarCategoriaEnSeccionProductos as seleccionarCategoriaEnSeccionProductosShared,
} from '../../helpers/portal/productosShared';
import {
  seleccionarDropdownEnScopePorTexto as seleccionarDropdownEnScopePorTextoShared,
  seleccionarProductoPorDropdownSecundario as seleccionarProductoPorDropdownSecundarioShared,
} from '../../helpers/portal/productoSelectionShared';
import { detectarProductoSeleccionadoEnUI as detectarProductoSeleccionadoEnUIShared } from '../../helpers/portal/productoDetection';
import { asegurarSeccionProductosVisible } from '../../helpers/portal/productosStepNavigation';
import { launchPortalSession } from '../../../src/infra/playwright/browserSession';
import { abrirArchivoAlFinal, limpiarEvidenciasTemporales } from '../../../src/services/common/evidenciasService';
import { closeBrowserSession, runRegistros } from '../../../src/services/common/registroRunner';

const capturas: string[] = [];

const getBotonContinuar = (page: Page) =>
  page.locator('button:visible').filter({ hasText: /^Continuar\b/i }).last();

const escapeRegexText = (value: string) => escapeRegexTextShared(value);
const extraerCodigoProducto = (tipoCuenta: string) => {
  return extraerCodigoProductoShared(tipoCuenta);
};
const extraerNombreProducto = (tipoCuenta: string) => extraerNombreProductoShared(tipoCuenta);
const esValorDropdownVacio = (value: string) => esValorDropdownVacioShared(value);
const LIST_PROBE_TIMEOUT = 2500;
const LIST_RETRY_CLICK_TIMEOUT = 1200;
const LIST_RETRY_WAIT_MS = 600;
const LIST_LABEL_TIMEOUT = 5000;
const LIST_PANEL_TIMEOUT = 5000;
const LIST_PANEL_QUICK_TIMEOUT = 3200;
const LIST_SCOPE_RETRY_WAIT_MS = 300;
const PRODUCTO_MAX_INTENTOS = 4;
const PRODUCTO_WAIT_CONFIRM_MS = 1800;
const PRODUCTO_RETRY_WAIT_MS = 240;
const FINALIZACION_MAX_INTENTOS = 6;
const FINALIZACION_TIMEOUT_MS = 120000;
const BIZAGI_PASSWORD = "H0la1234.";

function esErrorCasoEnCursoNoGestionable(errorMsg: string) {
  const txt = String(errorMsg ?? "");
  const timeout = /Timeout\s+\d+ms exceeded/i.test(txt);
  const tablaCasos = /hasText:\s*\/Casos\/i/i.test(txt);
  const filaMpn = /tr:has-text\("MPN-\d+/i.test(txt);
  return timeout && tablaCasos && filaMpn;
}

async function modalProductoConfigVisible(page: Page) {
  return modalProductoConfigVisibleShared(page, /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i);
}

async function localizarSeccionProductos(page: Page) {
  return localizarSeccionProductosShared(page);
}

async function seleccionarCategoriaEnSeccionProductos(page: Page, seccionProductos: Locator) {
  return seleccionarCategoriaEnSeccionProductosShared(page, seccionProductos, {
    panelQuickTimeout: LIST_PANEL_QUICK_TIMEOUT,
    panelTimeout: LIST_PANEL_TIMEOUT,
    scopeRetryWaitMs: LIST_SCOPE_RETRY_WAIT_MS,
  }, {
    onBeforeOpen: async (currentPage) => {
      await cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false);
    },
  });
}

async function leerValorDropdownEnScope(
  scope: Locator,
  labelText: string | RegExp,
  options?: { timeoutMs?: number }
) {
  return leerValorDropdownEnScopeShared(scope, labelText, options);
}

async function detectarProductoAgregadoEnUI(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string
) {
  return detectarProductoAgregadoEnUIShared(page, seccionProductos, tipoCuenta, { useGlobalFallback: true });
}

async function modalProductoConfigVisibleLegacy(page: Page) {
  const modalProducto = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
    .first();
  return modalProducto.isVisible().catch(() => false);
}

async function localizarSeccionProductosLegacy(page: Page) {
  const candidatos: Locator[] = [
    page
      .locator('xpath=(//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Categoria de producto")] and .//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Producto")]][1])')
      .first(),
    page
      .locator('xpath=(//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú¿?","AEIOUaeiou"),"Que quiere el cliente")] and .//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Producto")] and .//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Relacionados")]][1])')
      .first(),
    page
      .locator('xpath=(//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Productos")] and .//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Categoria de producto")]][1])')
      .first(),
  ];

  for (const candidato of candidatos) {
    const visible = await candidato.isVisible().catch(() => false);
    if (!visible) continue;
    const dropdownsVisibles = await candidato
      .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
      .count()
      .catch(() => 0);
    if (dropdownsVisibles >= 2) return candidato;
  }

  return candidatos[0];
}

async function seleccionarCategoriaEnSeccionProductosLegacy(page: Page, seccionProductos: Locator) {
  const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
  const count = await dropdowns.count().catch(() => 0);
  if (count < 1) {
    throw new Error("[CRITICO] No se encontro dropdown de 'Categoria de producto' en la seccion de Productos.");
  }

  const categoriaDropdown = dropdowns.first();
  const labelCategoria = categoriaDropdown.locator('.p-dropdown-label, [data-pc-section="label"]').first();
  let valor = ((await labelCategoria.textContent().catch(() => "")) || "").trim();
  if (!esValorDropdownVacio(valor)) return;

  for (let intento = 1; intento <= 4; intento++) {
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    await categoriaDropdown.scrollIntoViewIfNeeded().catch(() => { });
    await categoriaDropdown.click({ force: true }).catch(() => { });

    const combobox = categoriaDropdown.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
    let panel: Locator | null = null;
    if (panelId) {
      const byId = page.locator(`#${panelId}`);
      const visible = await byId.waitFor({ state: "visible", timeout: LIST_PANEL_QUICK_TIMEOUT })
        .then(() => true)
        .catch(() => false);
      if (visible) panel = byId;
    }
    if (!panel) {
      const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
      const visible = await fallback.waitFor({ state: "visible", timeout: LIST_PANEL_QUICK_TIMEOUT })
        .then(() => true)
        .catch(() => false);
      if (visible) panel = fallback;
    }

    if (panel) {
      const btnRetryPanel = panel
        .locator('button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button.p-button-warning')
        .first();
      if (await btnRetryPanel.isVisible().catch(() => false)) {
        await btnRetryPanel.click({ force: true }).catch(() => { });
        await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
      }

      const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
      const listo = await items.first().waitFor({ state: "visible", timeout: LIST_PANEL_TIMEOUT })
        .then(() => true)
        .catch(() => false);
      const countItems = await items.count().catch(() => 0);
      if (listo && countItems > 0) {
        const itemCategoria = items.filter({ hasText: /Cuentas de Efectivo/i }).first();
        if (await itemCategoria.isVisible().catch(() => false)) {
          await itemCategoria.click({ force: true }).catch(() => { });
        } else {
          await items.nth(0).click({ force: true }).catch(() => { });
        }
      }
    }

    await page.waitForTimeout(120);
    valor = ((await labelCategoria.textContent().catch(() => "")) || "").trim();
    if (!esValorDropdownVacio(valor)) return;
    await page.keyboard.press("Escape").catch(() => { });
    await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
  }

  throw new Error("[CRITICO] No se pudo seleccionar 'Categoria de producto' en la seccion de Productos.");
}

async function leerValorDropdownEnScopeLegacy(
  scope: Locator,
  labelText: string | RegExp,
  options?: { timeoutMs?: number }
) {
  const timeoutMs = options?.timeoutMs ?? 5000;
  const pattern = labelText instanceof RegExp
    ? labelText
    : new RegExp(String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

  let label = scope.locator("label").filter({ hasText: pattern }).first();
  if ((await label.count().catch(() => 0)) === 0) {
    label = scope.getByText(pattern).first();
  }
  await label.waitFor({ state: "visible", timeout: timeoutMs });

  let field = label
    .locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]')
    .first();
  if (!(await field.isVisible().catch(() => false))) {
    field = label
      .locator('xpath=following::*[self::div[contains(@class,"p-dropdown")] or @data-pc-name="dropdown"][1]')
      .first();
  }
  await field.waitFor({ state: "visible", timeout: timeoutMs });
  return ((await field.locator('.p-dropdown-label, [data-pc-section="label"]').first().textContent().catch(() => "")) || "").trim();
}

async function detectarProductoAgregadoEnUILegacy(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string
) {
  const tipoCuentaSafe = String(tipoCuenta ?? "").trim();
  if (!tipoCuentaSafe) return false;

  const tipoCuentaRegex = new RegExp(escapeRegexText(tipoCuentaSafe).replace(/\s+/g, "\\s+"), "i");
  const nombreProducto = extraerNombreProducto(tipoCuentaSafe);
  const nombreRegex = nombreProducto
    ? new RegExp(escapeRegexText(nombreProducto).replace(/\s+/g, "\\s+"), "i")
    : null;
  const codigoProducto = extraerCodigoProducto(tipoCuentaSafe);
  const codigoRegex = codigoProducto ? new RegExp(`\\b${escapeRegexText(codigoProducto)}\\b`, "i") : null;

  const msgSinProductos = page.getByText(/No se agregaron productos en simulaci[oó]n/i).first();
  const sinProductosVisible = await msgSinProductos.isVisible().catch(() => false);
  if (sinProductosVisible) return false;

  const selectorTextoProductoAgregado =
    'xpath=.//*[' +
    'not(ancestor::*[contains(@class,"p-dropdown") or contains(@class,"p-dropdown-panel") or @role="listbox" or @role="option"])' +
    ' and not(descendant::*[contains(@class,"p-dropdown") or contains(@class,"p-dropdown-panel") or @role="listbox" or @role="option"])' +
    ' and (self::span or self::td or self::p or self::small or self::strong or self::a or self::li or self::div)' +
    ']';

  const textosFueraDeDropdown = seccionProductos.locator(selectorTextoProductoAgregado);
  const textosFueraDeDropdownGlobal = page.locator(selectorTextoProductoAgregado);

  const visibleTextoCompleto = await textosFueraDeDropdown.filter({ hasText: tipoCuentaRegex }).first().isVisible().catch(() => false);
  if (visibleTextoCompleto) return true;
  const visibleTextoCompletoGlobal = await textosFueraDeDropdownGlobal.filter({ hasText: tipoCuentaRegex }).first().isVisible().catch(() => false);
  if (visibleTextoCompletoGlobal) return true;

  if (nombreRegex) {
    const visibleNombre = await textosFueraDeDropdown.filter({ hasText: nombreRegex }).first().isVisible().catch(() => false);
    if (visibleNombre) return true;
    const visibleNombreGlobal = await textosFueraDeDropdownGlobal.filter({ hasText: nombreRegex }).first().isVisible().catch(() => false);
    if (visibleNombreGlobal) return true;
  }

  if (codigoRegex) {
    const visibleCodigo = await textosFueraDeDropdown.filter({ hasText: codigoRegex }).first().isVisible().catch(() => false);
    if (visibleCodigo) return true;
    const visibleCodigoGlobal = await textosFueraDeDropdownGlobal.filter({ hasText: codigoRegex }).first().isVisible().catch(() => false);
    if (visibleCodigoGlobal) return true;
  }

  return false;
}

async function cerrarModalCancelarProcesoSiVisible(page: Page) {
  const modalCancelarProceso = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Cancelar proceso de solicitud/i })
    .first();
  const visible = await modalCancelarProceso.isVisible().catch(() => false);
  if (!visible) return false;

  const btnMantenerProceso = modalCancelarProceso.getByRole('button', { name: /^Cancelar$/i }).first();
  if (await btnMantenerProceso.isVisible().catch(() => false)) {
    await btnMantenerProceso.click({ force: true }).catch(() => { });
    await modalCancelarProceso.waitFor({ state: 'hidden', timeout: 6000 }).catch(() => { });
    return true;
  }
  return false;
}

async function esperarFinActualizandoSolicitud(page: Page, timeoutMs = 120000) {
  const inicio = Date.now();
  const txtActualizando = page.getByText(/Actualizando solicitud/i).first();
  const txtDepurando = page.getByText(/Depurando solicitante|Consultando datos del solicitante/i).first();
    const overlays = page.locator(
        '.p-blockui:visible, [data-pc-name="blockui"]:visible, .p-progressspinner:visible, .p-progress-spinner:visible, [role="progressbar"]:visible'
    );

  while (Date.now() - inicio < timeoutMs) {
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    const actualizandoVisible = await txtActualizando.isVisible().catch(() => false);
    const depurandoVisible = await txtDepurando.isVisible().catch(() => false);
    const overlayCount = await overlays.count().catch(() => 0);

    if (!actualizandoVisible && !depurandoVisible && overlayCount === 0) {
      return true;
    }

    await page.waitForTimeout(350);
  }

  return false;
}

async function resolverNoPoseeCorreoSiFalta(page: Page) {
  const msgFaltaCorreo = page.getByText(/Debe agregar al menos un Correo Electr(?:o|\u00f3)nico/i).first();

  const msgVisible = await msgFaltaCorreo.isVisible().catch(() => false);
  if (!msgVisible) return false;

  console.log("[Correo] Detectado 'Debe agregar al menos un Correo Electronico'. Activando 'No posee correo electronico'.");

  const inputSwitch = page
    .locator('xpath=//*[contains(normalize-space(.),"No posee correo electr")]/following::input[@type="checkbox"][1]')
    .first();
  const switchVisible = await inputSwitch.isVisible().catch(() => false);

  if (switchVisible) {
    const checked = await inputSwitch.isChecked().catch(() => false);
    if (!checked) {
      const checkedByApi = await inputSwitch
        .check({ force: true, timeout: 1200 })
        .then(() => true)
        .catch(() => false);
      if (!checkedByApi) {
        await inputSwitch.evaluate((el) => {
          const input = el as HTMLInputElement;
          input.checked = true;
          input.dispatchEvent(new Event("input", { bubbles: true }));
          input.dispatchEvent(new Event("change", { bubbles: true }));
          input.dispatchEvent(new Event("blur", { bubbles: true }));
        }).catch(() => { });
      }
      await page.waitForTimeout(120);
    }
  } else {
    console.log("[Correo] No se encontro el switch 'No posee correo electronico'.");
  }

  const btnContinuar = getBotonContinuar(page);
  const continuarVisible = await btnContinuar.isVisible().catch(() => false);
  const continuarEnabled = continuarVisible ? await btnContinuar.isEnabled().catch(() => false) : false;
  if (continuarVisible && continuarEnabled) {
    await btnContinuar.click({ force: true, timeout: 1800 }).catch(() => { });
    await esperarFinActualizandoSolicitud(page, 8000).catch(() => false);
    await page.waitForTimeout(250);
  }

  const msgSigueVisible = await msgFaltaCorreo.isVisible().catch(() => false);
  return !msgSigueVisible;
}

async function resolverCorreoPredeterminadoYContinuarSiVisible(page: Page) {
  const msgCorreoPred = page.getByText(/Debe agregar un Correo Electr(?:o|\u00f3)nico como predeterminado/i).first();
  const visible = await msgCorreoPred.isVisible().catch(() => false);
  if (!visible) return false;

  console.log("[Correo] Detectado mensaje de correo predeterminado. Corrigiendo y avanzando con Continuar.");

  for (let intento = 1; intento <= 2; intento++) {
    await validarCorreoPredeterminadoYCorregir(page).catch(() => false);
    await page.waitForTimeout(180);

    const btnContinuar = getBotonContinuar(page);
    const continuarVisible = await btnContinuar.isVisible().catch(() => false);
    const continuarEnabled = continuarVisible ? await btnContinuar.isEnabled().catch(() => false) : false;
    if (continuarVisible && continuarEnabled) {
      await btnContinuar.click({ force: true, timeout: 1800 }).catch(() => { });
      await esperarFinActualizandoSolicitud(page, 7000).catch(() => false);
      await page.waitForTimeout(220);
    }

    const sigueVisible = await msgCorreoPred.isVisible().catch(() => false);
    if (!sigueVisible) return true;
  }

  return false;
}

async function asegurarTiempoEnVivienda(page: Page, valor = "0") {
  const candidatos = [
    page.locator("label").filter({ hasText: /Tiempo en vivienda/i }).first().locator('xpath=following::input[1]').first(),
    page.locator('xpath=//*[contains(normalize-space(.),"Tiempo en vivienda")]/following::input[1]').first(),
  ];

  for (const input of candidatos) {
    const visible = await input.isVisible().catch(() => false);
    if (!visible) continue;
    const editable = await input.isEditable().catch(() => false);
    if (!editable) continue;

    const actual = (await input.inputValue().catch(() => "")).trim();
    if (actual) return false;

    await input.scrollIntoViewIfNeeded().catch(() => { });
    await input.click({ force: true }).catch(() => { });
    await page.keyboard.press("Control+A").catch(() => { });
    await page.keyboard.press("Delete").catch(() => { });
    await input.fill(valor).catch(async () => {
      await input.pressSequentially(valor, { delay: 20 }).catch(() => { });
    });
    await input.dispatchEvent("input").catch(() => { });
    await input.dispatchEvent("change").catch(() => { });
    await input.blur().catch(() => { });
    await page.waitForTimeout(150);

    let after = (await input.inputValue().catch(() => "")).trim();
    if (!after) {
      await input.evaluate((el: any, v: string) => {
        const inputEl = el instanceof HTMLInputElement ? el : null;
        if (!inputEl) return;
        inputEl.value = String(v);
        inputEl.dispatchEvent(new Event("input", { bubbles: true }));
        inputEl.dispatchEvent(new Event("change", { bubbles: true }));
        inputEl.dispatchEvent(new Event("blur", { bubbles: true }));
      }, valor).catch(() => { });
      after = (await input.inputValue().catch(() => "")).trim();
    }

    if (after) {
      console.log("[Validacion] 'Tiempo en vivienda' estaba vacio. Se completo con 0.");
      return true;
    }
  }

  return false;
}

async function estaEnPantallaPreProductos(page: Page) {
  const marcadores: Locator[] = [
    page.getByText(/Nivel de estudio/i).first(),
    page.getByText(/Informaci(?:o|\u00f3)n FATCA/i).first(),
    page.getByText(/Referencia personal/i).first(),
    page.getByText(/\u00bfPosee otras identificaciones o nacionalidades\?/i).first(),
    page.getByRole('button', { name: /A(?:\u00f1|n)adir direcci(?:o|\u00f3)n/i }).first(),
  ];

  for (const m of marcadores) {
    if (await m.isVisible().catch(() => false)) {
      return true;
    }
  }
  return false;
}

function modalSolicitudFinalizadaLocatorLegacy(page: Page) {
  return page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i })
    .last();
}

async function cerrarModalSolicitudFinalizadaLegacy(page: Page, options?: { timeoutMs?: number; waitForAppearMs?: number }) {
  const timeoutMs = options?.timeoutMs ?? 60000;
  const waitForAppearMs = options?.waitForAppearMs ?? Math.min(timeoutMs, 7000);
  const modalFinalizada = modalSolicitudFinalizadaLocator(page);
  const visibleAhora = await modalFinalizada.isVisible().catch(() => false);
  if (!visibleAhora) {
    const aparecioRapido = await modalFinalizada
      .waitFor({ state: 'visible', timeout: waitForAppearMs })
      .then(() => true)
      .catch(() => false);
    if (!aparecioRapido) return false;
  }
  console.log("[Finalizacion] Modal 'Solicitud finalizada' visible. Intentando cerrar.");

  for (let intento = 1; intento <= 5; intento++) {
    const botonesCerrar = [
      modalFinalizada.getByRole('button', { name: /^Cerrar$/i }).first(),
      modalFinalizada.getByRole('button', { name: /^Aceptar$/i }).first(),
      modalFinalizada.locator('button[data-pc-name="acceptbutton"], button.p-confirm-dialog-accept').first(),
      modalFinalizada.locator('button:has-text("Cerrar"), button:has-text("Aceptar")').first(),
      modalFinalizada.locator('.p-dialog-footer button:visible').last(),
    ];

    let hizoClick = false;
    for (const btn of botonesCerrar) {
      const visible = await btn.isVisible().catch(() => false);
      if (!visible) continue;
      await btn.scrollIntoViewIfNeeded().catch(() => { });
      const clicked = await btn.click({ force: true }).then(() => true).catch(() => false);
      if (clicked) {
        hizoClick = true;
        break;
      }
    }

    if (!hizoClick) {
      const clickTextoCerrar = await modalFinalizada.evaluate((el) => {
        const botones = Array.from(el.querySelectorAll('button'));
        const btnCerrar = botones.find((b) => /cerrar/i.test((b.textContent || '').trim()));
        if (!btnCerrar) return false;
        (btnCerrar as HTMLButtonElement).click();
        return true;
      }).catch(() => false);
      if (clickTextoCerrar) {
        hizoClick = true;
      }
    }

    if (!hizoClick) {
      const btnCerrarHeader = modalFinalizada
        .locator('button[aria-label="Close"], button[aria-label="Cerrar"], button.p-dialog-header-close, [data-pc-section="closebutton"]')
        .first();
      if (await btnCerrarHeader.isVisible().catch(() => false)) {
        await btnCerrarHeader.click({ force: true }).catch(() => { });
      } else {
        await page.keyboard.press('Escape').catch(() => { });
      }
    }

    await page.waitForTimeout(250);
    const sigueVisible = await modalFinalizada
      .waitFor({ state: 'hidden', timeout: 1500 })
      .then(() => false)
      .catch(async () => modalFinalizada.isVisible().catch(() => false));
    if (!sigueVisible) return true;
  }

  const intentoClickDirecto = await modalFinalizada
    .evaluate((el) => {
      const botones = Array.from(el.querySelectorAll('button'));
      const objetivo = botones.find((b) => /(cerrar|aceptar)/i.test((b.textContent || '').trim()));
      if (!objetivo) return false;
      (objetivo as HTMLButtonElement).click();
      return true;
    })
    .catch(() => false);

  if (intentoClickDirecto) {
    await page.waitForTimeout(400);
  }

  return !(await modalFinalizada.isVisible().catch(() => false));
}

async function estadoSolicitudFinalizadaVisibleLegacy(page: Page) {
  const candidatos = [
    page.getByText(/Solicitud finalizada/i),
    page.locator(
      'xpath=//*[not(ancestor::*[contains(@class,"p-dialog") or @role="dialog"])][contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚ","abcdefghijklmnopqrstuvwxyzáéíóú"),"solicitud finalizada")]'
    ),
  ];
  for (const locator of candidatos) {
    const total = await locator.count().catch(() => 0);
    if (!total) continue;
    for (let i = 0; i < total; i++) {
      const vis = await locator.nth(i).isVisible().catch(() => false);
      if (vis) return true;
    }
  }
  return false;
}

async function estadoPendienteFinalizarVisibleLegacy(page: Page) {
  const candidatos = [
    page.getByText(/Pendiente finalizar/i),
    page.locator(
      'xpath=//*[not(ancestor::*[contains(@class,"p-dialog") or @role="dialog"])][contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚ","abcdefghijklmnopqrstuvwxyzáéíóú"),"pendiente finalizar")]'
    ),
  ];
  for (const locator of candidatos) {
    const total = await locator.count().catch(() => 0);
    if (!total) continue;
    for (let i = 0; i < total; i++) {
      const vis = await locator.nth(i).isVisible().catch(() => false);
      if (vis) return true;
    }
  }
  return false;
}

async function cerrarModalFinalizadaRapidoLegacy(page: Page) {
  const modal = modalSolicitudFinalizadaLocator(page);
  const visible = await modal.isVisible().catch(() => false);
  if (!visible) return true;

  const btnCerrar = modal.getByRole('button', { name: /^Cerrar$/i }).first();
  const btnAceptar = modal.getByRole('button', { name: /^Aceptar$/i }).first();
  const btnAny = modal.locator('button:has-text("Cerrar"), button:has-text("Aceptar")').first();

  const candidatos = [btnCerrar, btnAceptar, btnAny];
  for (const btn of candidatos) {
    const btnVisible = await btn.isVisible().catch(() => false);
    if (!btnVisible) continue;

    const clicked = await btn.click({ force: true, timeout: 1200 }).then(() => true).catch(() => false);
    if (!clicked) {
      const handle = await btn.elementHandle().catch(() => null);
      if (handle) {
        await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => { });
      }
    }

    const hidden = await modal.waitFor({ state: 'hidden', timeout: 1800 }).then(() => true).catch(() => false);
    if (hidden) return true;
  }

  await page.keyboard.press("Escape").catch(() => { });
  await page.waitForTimeout(250);
  return !(await modal.isVisible().catch(() => false));
}

async function confirmarFinalizacionSiHayPopupLegacy(page: Page) {
  const dialogosVisibles = page.locator('.p-dialog:visible, [role="dialog"]:visible');
  const total = await dialogosVisibles.count().catch(() => 0);
  if (!total) return false;

  for (let i = total - 1; i >= 0; i--) {
    const dialogo = dialogosVisibles.nth(i);
    const esFinalizada = await dialogo
      .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i })
      .isVisible()
      .catch(() => false);
    if (esFinalizada) continue;

    const botones = [
      dialogo.locator('button.p-confirm-dialog-accept, button[data-pc-name="acceptbutton"]').first(),
      dialogo.getByRole('button', { name: /^(Aceptar|Si|S[i\u00ed]|Finalizar)$/i }).first(),
      dialogo.locator('button:has-text("Aceptar"), button:has-text("Si"), button:has-text("S\u00ed"), button:has-text("Finalizar")').first(),
    ];
    for (const btn of botones) {
      const visible = await btn.isVisible().catch(() => false);
      if (!visible) continue;
      const enabled = await btn.isEnabled().catch(() => false);
      if (!enabled) continue;
      await btn.click({ force: true }).catch(() => { });
      await page.waitForTimeout(220);
      return true;
    }
  }
  return false;
}

async function clickBotonFinalizarRobustoLegacy(page: Page) {
  const candidatos = [
    page.getByRole('button', { name: /^Finalizar$/i }).first(),
    page.locator('button:has-text("Finalizar")').first(),
    page.locator('button.p-button:visible').filter({ hasText: /^Finalizar$/i }).first(),
  ];

  for (const btn of candidatos) {
    const visible = await btn.isVisible().catch(() => false);
    if (!visible) continue;
    const enabled = await btn.isEnabled().catch(() => false);
    if (!enabled) continue;
    await btn.scrollIntoViewIfNeeded().catch(() => { });
    const clicked = await btn.click({ force: true, timeout: 2800 }).then(() => true).catch(() => false);
    if (clicked) return true;
  }

  return false;
}

async function irAPasoFinalDesdeWizardLegacy(page: Page) {
  const candidatoPaso4 = page.locator(
    'li.p-steps-item:has-text("4"), [data-pc-section="menuitem"]:has-text("4"), .p-steps-item:has(.p-steps-number:has-text("4"))'
  ).first();
  const visible = await candidatoPaso4.isVisible().catch(() => false);
  if (!visible) return false;

  await candidatoPaso4.click({ force: true }).catch(() => { });
  await esperarFinActualizandoSolicitud(page, 12000).catch(() => false);
  await page.waitForTimeout(220);
  return true;
}

async function avanzarAPasoFinalSiHaceFaltaLegacy(page: Page) {
  const btnFinalizar = page.getByRole('button', { name: /^Finalizar$/i }).first();
  for (let i = 1; i <= 3; i++) {
    const finalizarVisible = await btnFinalizar.isVisible().catch(() => false);
    const finalizarEnabled = finalizarVisible ? await btnFinalizar.isEnabled().catch(() => false) : false;
    if (finalizarVisible && finalizarEnabled) return true;

    await irAPasoFinalDesdeWizard(page).catch(() => false);
    const finalizaTrasPaso4 = await btnFinalizar.isVisible().catch(() => false);
    if (finalizaTrasPaso4) return true;

    await asegurarTiempoEnVivienda(page, "0").catch(() => false);
    await resolverNoPoseeCorreoSiFalta(page).catch(() => false);
    await resolverCorreoPredeterminadoYContinuarSiVisible(page).catch(() => false);

    const btnContinuar = getBotonContinuar(page);
    const continuarVisible = await btnContinuar.isVisible().catch(() => false);
    const continuarEnabled = continuarVisible ? await btnContinuar.isEnabled().catch(() => false) : false;
    if (!continuarVisible || !continuarEnabled) break;

    await btnContinuar.scrollIntoViewIfNeeded().catch(() => { });
    await btnContinuar.click({ force: true, timeout: 2200 }).catch(() => { });
    await esperarFinActualizandoSolicitud(page, 18000).catch(() => false);
    await confirmarFinalizacionSiHayPopup(page).catch(() => false);
    await page.waitForTimeout(220);
  }

  return await btnFinalizar.isVisible().catch(() => false);
}

async function finalizarSolicitudSiPendienteLegacy(
  page: Page,
  registro: RegistroExcel,
  options?: { timeoutMs?: number; maxIntentos?: number }
) {
  const timeoutMs = options?.timeoutMs ?? FINALIZACION_TIMEOUT_MS;
  const maxIntentos = options?.maxIntentos ?? FINALIZACION_MAX_INTENTOS;
  const inicio = Date.now();

  const modalFinalizadaVisibleInicial = await modalSolicitudFinalizadaLocator(page).isVisible().catch(() => false);
  if (modalFinalizadaVisibleInicial) {
    const cerroInicial = await cerrarModalFinalizadaRapido(page).catch(() => false);
    if (cerroInicial) return true;
  }

  const yaFinalizada = await estadoSolicitudFinalizadaVisible(page);
  if (yaFinalizada) return true;

  for (let intento = 1; intento <= maxIntentos; intento++) {
    if (Date.now() - inicio >= timeoutMs) break;
    console.log(`[Finalizacion][${registro.identificacion}] Intento ${intento}/${maxIntentos} para finalizar solicitud.`);

    const cerroModalPrevio = await cerrarModalSolicitudFinalizada(page, {
      timeoutMs: 5000,
      waitForAppearMs: 900,
    }).catch(() => false);
    if (cerroModalPrevio) return true;

    const modalFinalizadaVisible = await modalSolicitudFinalizadaLocator(page).isVisible().catch(() => false);
    if (modalFinalizadaVisible) {
      const cerro = await cerrarModalFinalizadaRapido(page).catch(() => false);
      if (!cerro) {
        console.log(`[Finalizacion][${registro.identificacion}] Modal 'Solicitud finalizada' detectado pero no cerro al primer intento. Se considera finalizada.`);
      }
      return true;
    }

    await avanzarAPasoFinalSiHaceFalta(page).catch(() => false);

    const btnFinalizar = page.getByRole('button', { name: /^Finalizar$/i }).first();
    const finalizarVisible = await btnFinalizar.isVisible().catch(() => false);
    const finalizarEnabled = finalizarVisible ? await btnFinalizar.isEnabled().catch(() => false) : false;
    const pendiente = await estadoPendienteFinalizarVisible(page).catch(() => false);
    console.log(
      `[Finalizacion][${registro.identificacion}] estado: finalizarVisible=${finalizarVisible} finalizarEnabled=${finalizarEnabled} pendiente=${pendiente}`
    );
    if (finalizarVisible && finalizarEnabled) {
      await clickBotonFinalizarRobusto(page).catch(() => false);
      await page.waitForTimeout(300);
      await confirmarFinalizacionSiHayPopup(page).catch(() => false);

      const cerroModalTrasFinalizar = await cerrarModalSolicitudFinalizada(page, {
        timeoutMs: 12000,
        waitForAppearMs: 9000,
      }).catch(() => false);
      if (cerroModalTrasFinalizar) return true;

      const adjuntoDocumentos = await validarGestionDocumentalSiRequerido(page).catch(() => false);
      if (adjuntoDocumentos) {
        await page.waitForTimeout(650);
        if (await btnFinalizar.isVisible().catch(() => false)) {
          await clickBotonFinalizarRobusto(page).catch(() => false);
          await page.waitForTimeout(300);
          await confirmarFinalizacionSiHayPopup(page).catch(() => false);
          const cerroModalReintento = await cerrarModalSolicitudFinalizada(page, {
            timeoutMs: 12000,
            waitForAppearMs: 9000,
          }).catch(() => false);
          if (cerroModalReintento) return true;
        }
      }
    }

    await esperarFinActualizandoSolicitud(page, 18000).catch(() => false);
    await confirmarFinalizacionSiHayPopup(page).catch(() => false);
    const cerroModalPostEspera = await cerrarModalSolicitudFinalizada(page, {
      timeoutMs: 10000,
      waitForAppearMs: 7000,
    }).catch(() => false);
    if (cerroModalPostEspera) return true;
    await cerrarModalFinalizadaRapido(page).catch(() => false);

    const finalizada = await estadoSolicitudFinalizadaVisible(page);
    if (finalizada) return true;

    const modalFinal = await modalSolicitudFinalizadaLocator(page).isVisible().catch(() => false);
    if (modalFinal) {
      const cerro = await cerrarModalSolicitudFinalizada(page, { timeoutMs: 5000 }).catch(() => false);
      if (cerro) return true;
    }

    const pendientePost = await estadoPendienteFinalizarVisible(page).catch(() => false);
    if (pendientePost && !finalizarVisible) {
      await page.waitForTimeout(800);
    }
  }

  const finalizadaAlFinal = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
  const modalFinalizadaAlFinal = await modalSolicitudFinalizadaLocator(page).isVisible().catch(() => false);
  if (finalizadaAlFinal && modalFinalizadaAlFinal) {
    console.log(`[Finalizacion][${registro.identificacion}] Estado finalizado detectado, pero modal sigue abierto. Se delega cierre al cleanup para continuar el loop.`);
    return true;
  }

  // Ventana corta adicional: a veces el modal/estado final aparece segundos despues del ultimo intento.
  const modalAparecioTarde = await modalSolicitudFinalizadaLocator(page)
    .waitFor({ state: "visible", timeout: 3500 })
    .then(() => true)
    .catch(() => false);
  if (modalAparecioTarde) {
    await cerrarModalFinalizadaRapido(page).catch(() => false);
    return true;
  }
  const finalizadaTardia = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
  if (finalizadaTardia) return true;

  return false;
}

function modalSolicitudFinalizadaLocator(page: Page) {
  return modalSolicitudFinalizadaLocatorShared(page);
}

async function cerrarModalSolicitudFinalizada(page: Page, options?: { timeoutMs?: number; waitForAppearMs?: number }) {
  return cerrarModalSolicitudFinalizadaShared(page, options);
}

async function estadoSolicitudFinalizadaVisible(page: Page) {
  return estadoSolicitudFinalizadaVisibleShared(page);
}

async function estadoPendienteFinalizarVisible(page: Page) {
  return estadoPendienteFinalizarVisibleShared(page);
}

async function cerrarModalFinalizadaRapido(page: Page) {
  return cerrarModalFinalizadaRapidoShared(page);
}

async function confirmarFinalizacionSiHayPopup(page: Page) {
  return confirmarFinalizacionSiHayPopupShared(page);
}

async function clickBotonFinalizarRobusto(page: Page) {
  return clickBotonFinalizarRobustoShared(page);
}

async function irAPasoFinalDesdeWizard(page: Page) {
  return irAPasoFinalDesdeWizardShared(page, {
    stepNumber: 4,
    esperarFinActualizandoSolicitud,
  });
}

async function avanzarAPasoFinalSiHaceFalta(page: Page) {
  const btnFinalizar = page.getByRole('button', { name: /^Finalizar$/i }).first();
  for (let i = 1; i <= 3; i++) {
    const finalizarVisible = await btnFinalizar.isVisible().catch(() => false);
    const finalizarEnabled = finalizarVisible ? await btnFinalizar.isEnabled().catch(() => false) : false;
    if (finalizarVisible && finalizarEnabled) return true;

    await irAPasoFinalDesdeWizard(page).catch(() => false);
    const finalizaTrasPaso4 = await btnFinalizar.isVisible().catch(() => false);
    if (finalizaTrasPaso4) return true;

    await asegurarTiempoEnVivienda(page, "0").catch(() => false);
    await resolverNoPoseeCorreoSiFalta(page).catch(() => false);
    await resolverCorreoPredeterminadoYContinuarSiVisible(page).catch(() => false);

    const btnContinuar = getBotonContinuar(page);
    const continuarVisible = await btnContinuar.isVisible().catch(() => false);
    const continuarEnabled = continuarVisible ? await btnContinuar.isEnabled().catch(() => false) : false;
    if (!continuarVisible || !continuarEnabled) break;

    await btnContinuar.scrollIntoViewIfNeeded().catch(() => { });
    await btnContinuar.click({ force: true, timeout: 2200 }).catch(() => { });
    await esperarFinActualizandoSolicitud(page, 18000).catch(() => false);
    await confirmarFinalizacionSiHayPopup(page).catch(() => false);
    await page.waitForTimeout(220);
  }

  return await btnFinalizar.isVisible().catch(() => false);
}

async function finalizarSolicitudSiPendiente(
  page: Page,
  registro: RegistroExcel,
  options?: { timeoutMs?: number; maxIntentos?: number }
) {
  return finalizarSolicitudSiPendienteShared(page, registro, {
    timeoutMs: options?.timeoutMs ?? FINALIZACION_TIMEOUT_MS,
    maxIntentos: options?.maxIntentos ?? FINALIZACION_MAX_INTENTOS,
    stepNumber: 4,
    getBotonContinuar,
    esperarFinActualizandoSolicitud,
    beforeContinuar: async (currentPage) => {
      await asegurarTiempoEnVivienda(currentPage, "0").catch(() => false);
      await resolverNoPoseeCorreoSiFalta(currentPage).catch(() => false);
      await resolverCorreoPredeterminadoYContinuarSiVisible(currentPage).catch(() => false);
    },
    afterFinalizarClick: async (currentPage) => {
      const adjuntoDocumentos = await validarGestionDocumentalSiRequerido(currentPage).catch(() => false);
      if (!adjuntoDocumentos) return;

      await currentPage.waitForTimeout(650);
      const btnFinalizar = currentPage.getByRole('button', { name: /^Finalizar$/i }).first();
      if (await btnFinalizar.isVisible().catch(() => false)) {
        await clickBotonFinalizarRobustoShared(currentPage).catch(() => false);
        await currentPage.waitForTimeout(300);
        await confirmarFinalizacionSiHayPopupShared(currentPage).catch(() => false);
      }
    },
  });
}

async function confirmarSeleccionProductoRapida(
  page: Page,
  seccionProductos: Locator,
  tipoCuentaRegex: RegExp,
  codigoRegex: RegExp | null,
  waitMs = PRODUCTO_WAIT_CONFIRM_MS
) {
  return confirmarSeleccionProductoRapidaShared(
    page,
    seccionProductos,
    tipoCuentaRegex,
    codigoRegex,
    /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i,
    waitMs
  );
}

async function seleccionarDropdownEnScopePorTexto(
  page: Page,
  scope: Locator,
  labelText: string | RegExp,
  textoOpcion: RegExp,
  fallbackIndex = -1
) {
  return seleccionarDropdownEnScopePorTextoShared(
    page,
    scope,
    labelText,
    textoOpcion,
    fallbackIndex,
    {
      listLabelTimeout: LIST_LABEL_TIMEOUT,
      listPanelTimeout: LIST_PANEL_TIMEOUT,
      listPanelQuickTimeout: LIST_PANEL_QUICK_TIMEOUT,
      listScopeRetryWaitMs: LIST_SCOPE_RETRY_WAIT_MS,
    },
    {
      onBeforeAttempt: async (currentPage) => {
        await clickReintentarListaSiVisible(currentPage, labelText, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
      },
    }
  );
}

async function seleccionarProductoPorDropdownSecundario(
  page: Page,
  seccionProductos: Locator,
  tipoCuentaRegex: RegExp,
  codigoRegex: RegExp | null
) {
  return seleccionarProductoPorDropdownSecundarioShared(
    page,
    seccionProductos,
    tipoCuentaRegex,
    codigoRegex,
    {
      listLabelTimeout: LIST_LABEL_TIMEOUT,
      listPanelTimeout: LIST_PANEL_TIMEOUT,
      listPanelQuickTimeout: LIST_PANEL_QUICK_TIMEOUT,
      listScopeRetryWaitMs: LIST_SCOPE_RETRY_WAIT_MS,
    }
  );
}

async function detectarProductoSeleccionadoEnUI(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string
) {
  return detectarProductoSeleccionadoEnUIShared(page, seccionProductos, tipoCuenta, {
    escapeRegexText,
    extraerCodigoProducto,
    useGlobalFallback: true,
    modalRegex: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i,
  });
}

async function seleccionarDropdownEnScopePorTextoLegacy(
  page: Page,
  scope: Locator,
  labelText: string | RegExp,
  textoOpcion: RegExp,
  fallbackIndex = -1
) {
  const pattern = labelText instanceof RegExp
    ? labelText
    : new RegExp(String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

  const clickReintentarScope = async () => {
    const btnRetry = scope
      .locator(
        'button:has-text("Reintentar buscar lista"), ' +
        'button:has-text("Reintentar"), ' +
        'button[aria-label*="Reintentar"], ' +
        'button.p-button-warning'
      )
      .first();
    const visible = await btnRetry.isVisible().catch(() => false);
    if (!visible) return false;
    await btnRetry.scrollIntoViewIfNeeded().catch(() => { });
    await btnRetry.click({ force: true }).catch(() => { });
    await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
    return true;
  };

  for (let intento = 1; intento <= 3; intento++) {
    await clickReintentarScope().catch(() => false);
    await clickReintentarListaSiVisible(page, labelText, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);

    let label = scope.locator("label").filter({ hasText: pattern }).first();
    if ((await label.count().catch(() => 0)) === 0) {
      label = scope.getByText(pattern).first();
    }
    const labelVisible = await label.waitFor({ state: "visible", timeout: LIST_LABEL_TIMEOUT }).then(() => true).catch(() => false);
    if (!labelVisible) continue;

    let field = label
      .locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]')
      .first();
    if (!(await field.isVisible().catch(() => false))) {
      field = label
        .locator('xpath=following::*[self::div[contains(@class,"p-dropdown")] or @data-pc-name="dropdown"][1]')
        .first();
    }
    const fieldVisible = await field.waitFor({ state: "visible", timeout: LIST_LABEL_TIMEOUT }).then(() => true).catch(() => false);
    if (!fieldVisible) continue;

    await field.scrollIntoViewIfNeeded().catch(() => { });
    await field.click({ force: true }).catch(() => { });

    const combobox = field.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
    let panel: Locator | null = null;
    if (panelId) {
      const byId = page.locator(`#${panelId}`);
      const byIdVisible = await byId.waitFor({ state: "visible", timeout: Math.min(LIST_PANEL_TIMEOUT, LIST_PANEL_QUICK_TIMEOUT) })
        .then(() => true)
        .catch(() => false);
      if (byIdVisible) panel = byId;
    }
    if (!panel) {
      const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
      const fallbackVisible = await fallback.waitFor({ state: "visible", timeout: Math.min(LIST_PANEL_TIMEOUT, LIST_PANEL_QUICK_TIMEOUT) })
        .then(() => true)
        .catch(() => false);
      if (fallbackVisible) panel = fallback;
    }
    if (!panel) {
      await page.keyboard.press("Escape").catch(() => { });
      await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
      continue;
    }

    const itemByText = panel
      .locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]')
      .filter({ hasText: textoOpcion })
      .first();
    const existeItemTexto = await itemByText.isVisible().catch(() => false);
    if (existeItemTexto) {
      await itemByText.scrollIntoViewIfNeeded().catch(() => { });
      await itemByText.click({ force: true });
      return;
    }

    const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
    const count = await items.count().catch(() => 0);
    if (fallbackIndex >= 0 && count > fallbackIndex) {
      await items.nth(fallbackIndex).click({ force: true }).catch(() => { });
      return;
    }

    await page.mouse.click(10, 10).catch(() => { });
    await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
  }

  throw new Error(`No se pudo seleccionar opcion '${String(textoOpcion)}' en '${String(labelText)}'.`);
}

async function seleccionarProductoPorDropdownSecundarioLegacy(
  page: Page,
  seccionProductos: Locator,
  tipoCuentaRegex: RegExp,
  codigoRegex: RegExp | null
) {
  const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
  const count = await dropdowns.count().catch(() => 0);
  console.log(`[ProductoFallback] dropdowns visibles en seccion=${count}`);
  if (count < 2) return false;

  // Usualmente: dropdown[0]=Categoria, dropdown[1]=Producto.
  const productoDropdown = dropdowns.nth(1);
  await productoDropdown.scrollIntoViewIfNeeded().catch(() => { });
  await productoDropdown.click({ force: true }).catch(() => { });

  const combobox = productoDropdown.locator('[role="combobox"]').first();
  const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
  let panel: Locator | null = null;
  if (panelId) {
    const byId = page.locator(`#${panelId}`);
    const byIdVisible = await byId.waitFor({ state: "visible", timeout: Math.min(LIST_PANEL_TIMEOUT, LIST_PANEL_QUICK_TIMEOUT) })
      .then(() => true)
      .catch(() => false);
    if (byIdVisible) panel = byId;
  }
  if (!panel) {
    const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
    const fallbackVisible = await fallback.waitFor({ state: "visible", timeout: Math.min(LIST_PANEL_TIMEOUT, LIST_PANEL_QUICK_TIMEOUT) })
      .then(() => true)
      .catch(() => false);
    if (fallbackVisible) panel = fallback;
  }
  if (!panel) return false;

  const btnRetryPanel = panel
    .locator('button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button.p-button-warning')
    .first();
  if (await btnRetryPanel.isVisible().catch(() => false)) {
    await btnRetryPanel.click({ force: true }).catch(() => { });
    await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
  }

  const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
  const hasItems = await items.first().waitFor({ state: "visible", timeout: LIST_PANEL_TIMEOUT })
    .then(() => true)
    .catch(() => false);
  const countItems = await items.count().catch(() => 0);
  console.log(`[ProductoFallback] panel visible, items=${countItems}`);
  if (!hasItems) return false;

  const byTipo = items.filter({ hasText: tipoCuentaRegex }).first();
  if (await byTipo.isVisible().catch(() => false)) {
    await byTipo.scrollIntoViewIfNeeded().catch(() => { });
    const clickedTipo = await byTipo
      .click({ force: true, timeout: 2500 })
      .then(() => true)
      .catch(() => false);
    if (!clickedTipo) {
      const handle = await byTipo.elementHandle().catch(() => null);
      if (handle) {
        await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => { });
      }
    }
    return true;
  }

  if (codigoRegex) {
    const byCodigo = items.filter({ hasText: codigoRegex }).first();
    if (await byCodigo.isVisible().catch(() => false)) {
      await byCodigo.scrollIntoViewIfNeeded().catch(() => { });
      const clickedCodigo = await byCodigo
        .click({ force: true, timeout: 2500 })
        .then(() => true)
        .catch(() => false);
      if (!clickedCodigo) {
        const handle = await byCodigo.elementHandle().catch(() => null);
        if (handle) {
          await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => { });
        }
      }
      return true;
    }
  }

  return false;
}

async function confirmarSeleccionProductoRapidaLegacy(
  page: Page,
  seccionProductos: Locator,
  tipoCuentaRegex: RegExp,
  codigoRegex: RegExp | null,
  waitMs = PRODUCTO_WAIT_CONFIRM_MS
) {
  const inicio = Date.now();
  while (Date.now() - inicio < waitMs) {
    if (await modalProductoConfigVisible(page)) return true;

    const dropdownProducto = seccionProductos
      .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
      .nth(1);
    const valorProducto = ((await dropdownProducto
      .locator('.p-dropdown-label, [data-pc-section="label"]')
      .first()
      .textContent()
      .catch(() => "")) || "").trim();

    if (!esValorDropdownVacio(valorProducto)) {
      const coincideTexto = tipoCuentaRegex.test(valorProducto);
      const coincideCodigo = codigoRegex ? codigoRegex.test(valorProducto) : true;
      if (coincideTexto || coincideCodigo) return true;
    }

    await page.waitForTimeout(110);
  }
  return false;
}

async function seleccionarProductoEnSeccionProductos(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string
) {
  return seleccionarProductoCuentaEfectivoExistente(page, seccionProductos, tipoCuenta, {
    confirmarSeleccionProductoRapida,
    detectarProductoSeleccionadoEnUI,
    esValorDropdownVacio,
    escapeRegexText,
    extraerCodigoProducto,
    leerValorDropdownEnScope,
    seleccionarDropdownEnScopePorTexto,
    seleccionarProductoPorDropdownSecundario,
    seleccionarDropdownFiltrableConReintentar: async (currentPage, textoFiltro) =>
      seleccionarDropdownFiltrableConReintentar(
        currentPage,
        'Producto',
        { texto: textoFiltro },
        { maxIntentos: 2, esperaMs: 240, timeoutCampoMs: 9000, timeoutPanelMs: 4500, usarFiltro: true }
      ),
    beforeAttempt: async (currentPage) => {
      if (await modalProductoConfigVisible(currentPage)) return;
      await cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false);
      await clickReintentarListaSiVisible(currentPage, 'Producto', LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
    },
    maxIntentos: PRODUCTO_MAX_INTENTOS,
    probeTimeoutMs: LIST_PROBE_TIMEOUT,
    retryWaitMs: PRODUCTO_RETRY_WAIT_MS,
  });
}

async function etapaFlujoRegistro(page: Page, registro: RegistroExcel) {
  const maxIntentosCasoActivo = 2;

  for (let intento = 1; intento <= maxIntentosCasoActivo; intento++) {
    await page.evaluate(() => {
      window.moveTo(0, 0);
      window.resizeTo(window.screen.availWidth, window.screen.availHeight);
    });

    await page.goto('https://srvqacgowb01.local.bsc.com:5000/requests/create/multiproduct', {
      waitUntil: 'domcontentloaded',
    });
    const estadoSesion = await esperarPortalListoTrasLogin(page, { timeoutMs: 240000 });
    if (estadoSesion.loginDetectado) {
      console.log('Sesion reanudada despues de login manual.');
    } else {
      console.log('Sesion cargada, sin MFA.');
    }

    const estaEnSolicitudMultiproducto = /\/requests\/create\/multiproduct/i.test(page.url());
    if (!estaEnSolicitudMultiproducto) {
      const btnSolicitudMultiproducto = page
        .locator('li[aria-label="Solicitud multiproducto"] a[data-pc-section="action"]:visible')
        .first();

      const puedeAbrirSolicitudMultiproducto = await btnSolicitudMultiproducto
        .waitFor({ state: 'visible', timeout: 4000 })
        .then(() => true)
        .catch(() => false);

      if (puedeAbrirSolicitudMultiproducto) {
        await btnSolicitudMultiproducto.click({ force: true });
        await page.waitForURL(/\/requests\/create\/multiproduct/i, { timeout: 60000 });
        await page.waitForLoadState('domcontentloaded');
      }
    }

    const modalCancelarProceso = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Cancelar proceso de solicitud/i })
      .first();
    const modalCancelarVisible = await modalCancelarProceso.isVisible().catch(() => false);
    if (modalCancelarVisible) {
      const btnMantenerProceso = modalCancelarProceso.getByRole('button', { name: /^Cancelar$/i }).first();
      if (await btnMantenerProceso.isVisible().catch(() => false)) {
        await btnMantenerProceso.click({ force: true });
        await modalCancelarProceso.waitFor({ state: 'hidden', timeout: 10000 }).catch(() => { });
      }
    }

    await clickReintentarListaSiVisible(page, 'Oficial');
    await clickReintentarListaSiVisible(page, 'Promotor');
    await clickReintentarListaSiVisible(page, 'Sexo');
    await clickReintentarListaSiVisible(page, /Pa[ií]s de nacimiento/i);
    await clickReintentarListaSiVisible(page, 'Estado civil');

    await asegurarIdentificacionHabilitada(page, { timeoutCampoMs: 60000 });
    const inputNumeroRefuerzo = page
      .locator('label')
      .filter({ hasText: /N[uú]mero de identificaci[oó]n/i })
      .first()
      .locator('xpath=following::input[1]')
      .first();
    await inputNumeroRefuerzo.waitFor({ state: 'visible', timeout: 15000 });
    let habilitadoNumero = await inputNumeroRefuerzo.isEnabled({ timeout: 5000 }).catch(() => false);
    if (!habilitadoNumero) {
      await esperarYClickReintentarPaisIdentificacion(page, 12000);
      const clicPais = await clickReintentarListaSiVisible(page, 'País de identificación', 2000).catch(() => false);
      if (!clicPais) {
        await clickReintentarListaSiVisible(page, 'Tipo de identificación', 2000).catch(() => false);
      }
      habilitadoNumero = await inputNumeroRefuerzo.isEnabled({ timeout: 5000 }).catch(() => false);
      if (!habilitadoNumero) {
        throw new Error("Input 'Número de identificación' sigue deshabilitado tras refuerzo.");
      }
    }

    const inputNumero = page
      .locator('label')
      .filter({ hasText: /N[uú]mero de identificaci[oó]n/i })
      .first()
      .locator('xpath=following::input[1]')
      .first();
    await inputNumero.waitFor({ state: 'visible', timeout: 15000 });
    await inputNumero.click({ force: true });
    await inputNumero.fill(registro.identificacion);

    await page.keyboard.press('Enter');
    const spinnerConsulta = page.getByText(/Consultando datos del solicitante/i);
    await spinnerConsulta.waitFor({ state: 'visible', timeout: 6000 }).catch(() => { });
    await spinnerConsulta.waitFor({ state: 'hidden', timeout: 120000 }).catch(() => { });
    await page.waitForTimeout(500);

    const inputFechaExp = page
      .locator('xpath=//*[contains(normalize-space(.),"Fecha de expir")]/following::input[contains(@class,"p-inputtext")][1]')
      .first();
    const fechaExpVisible = await inputFechaExp.isVisible().catch(() => false);
    if (fechaExpVisible) {
      let valorFechaExp = (await inputFechaExp.inputValue().catch(() => '')).trim();
      if (!valorFechaExp) {
        for (let i = 0; i < 5; i++) {
          await inputFechaExp.scrollIntoViewIfNeeded().catch(() => { });
          await inputFechaExp.click({ force: true });
          await page.keyboard.press('Control+A').catch(() => { });
          await page.keyboard.press('Delete').catch(() => { });
          await inputFechaExp.pressSequentially('16-08-2030', { delay: 35 }).catch(() => { });
          await inputFechaExp.dispatchEvent('input').catch(() => { });
          await inputFechaExp.dispatchEvent('change').catch(() => { });
          await inputFechaExp.blur().catch(() => { });
          await page.waitForTimeout(250);

          valorFechaExp = (await inputFechaExp.inputValue().catch(() => '')).trim();
          if (valorFechaExp) break;

          await inputFechaExp.evaluate((el) => {
            const input = el instanceof HTMLInputElement ? el : null;
            if (!input) return;
            input.value = '16-08-2030';
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            input.dispatchEvent(new Event('blur', { bubbles: true }));
          }).catch(() => { });
          await page.waitForTimeout(250);
          valorFechaExp = (await inputFechaExp.inputValue().catch(() => '')).trim();
          if (valorFechaExp) break;
        }
      }
    }

    await page.getByRole('button', { name: 'Depurar' }).click({ noWaitAfter: true });
    const spinnerDepurar = page.getByText(/Consultando datos del solicitante/i).first();
    await spinnerDepurar.waitFor({ state: 'visible', timeout: 8000 }).catch(() => { });
    await spinnerDepurar.waitFor({ state: 'hidden', timeout: 120000 }).catch(() => { });

    const mpnCasoActivo = await extraerCasoActivoMpn(page);
    if (mpnCasoActivo) {
      console.log(`[CasoActivo] Detectado ${mpnCasoActivo}. Cancelando en Bizagi...`);
      await cancelarCasoEnBizagiDesdePortal(page, mpnCasoActivo, { password: BIZAGI_PASSWORD });
      if (intento >= maxIntentosCasoActivo) {
        throw new Error(`[CRITICO] Se canceló ${mpnCasoActivo} pero el caso activo persiste para '${registro.identificacion}'.`);
      }
      console.log(`[CasoActivo] ${mpnCasoActivo} cancelado. Reintentando con la misma cédula (${registro.identificacion})...`);
      await page.waitForTimeout(1500);
      continue;
    }

    const msgFechaMinima = page.getByText(/La fecha m[ií]nima permitida es/i).first();
    const hayMsgFechaMinima = await msgFechaMinima.isVisible().catch(() => false);
    if (hayMsgFechaMinima) {
      const inputFechaExp2 = page
        .locator('xpath=//*[contains(normalize-space(.),"Fecha de expir")]/following::input[contains(@class,"p-inputtext")][1]')
        .first();
      await inputFechaExp2.scrollIntoViewIfNeeded().catch(() => { });
      await inputFechaExp2.click({ force: true }).catch(() => { });
      await page.keyboard.press('Control+A').catch(() => { });
      await page.keyboard.press('Delete').catch(() => { });
      await inputFechaExp2.fill('16-08-2030').catch(() => { });
      await inputFechaExp2.dispatchEvent('input').catch(() => { });
      await inputFechaExp2.dispatchEvent('change').catch(() => { });
      await inputFechaExp2.blur().catch(() => { });

      await page.getByRole('button', { name: 'Depurar' }).click({ noWaitAfter: true });
      await spinnerDepurar.waitFor({ state: 'visible', timeout: 8000 }).catch(() => { });
      await spinnerDepurar.waitFor({ state: 'hidden', timeout: 120000 }).catch(() => { });

      const mpnCasoActivo2 = await extraerCasoActivoMpn(page);
      if (mpnCasoActivo2) {
        console.log(`[CasoActivo] Detectado ${mpnCasoActivo2} en segundo depurar. Cancelando en Bizagi...`);
        await cancelarCasoEnBizagiDesdePortal(page, mpnCasoActivo2, { password: BIZAGI_PASSWORD });
        if (intento >= maxIntentosCasoActivo) {
          throw new Error(`[CRITICO] Se canceló ${mpnCasoActivo2} pero el caso activo persiste para '${registro.identificacion}'.`);
        }
        console.log(`[CasoActivo] ${mpnCasoActivo2} cancelado. Reintentando con la misma cédula (${registro.identificacion})...`);
        await page.waitForTimeout(1500);
        continue;
      }
    }

    await page.waitForTimeout(3000);
    await page.waitForTimeout(5000);
    await asegurarTiempoEnVivienda(page, "0").catch(() => false);
    await getBotonContinuar(page).click();
    await esperarFinActualizandoSolicitud(page, 25000).catch(() => false);
    await resolverNoPoseeCorreoSiFalta(page).catch(() => false);
    await page.waitForTimeout(4000);
    return;
  }

  throw new Error(`[CRITICO] No se pudo completar etapa inicial para '${registro.identificacion}' tras reintentos.`);
}

async function etapaValidacionesPrevias(page: Page) {
  let hizoValidaciones = false;
  const hizoTiempoExclusion = await llenarFechaSiVisibleYVacia(page, "Tiempo de exclusion", "16-08-2030", { debug: true });
  if (hizoTiempoExclusion) hizoValidaciones = true;

  const hizoCorreoPredeterminado = await validarCorreoPredeterminadoYCorregir(page);
  if (hizoCorreoPredeterminado) {
    hizoValidaciones = true;
    await resolverCorreoPredeterminadoYContinuarSiVisible(page).catch(() => false);
  }

  const hizoApnfd = await validarApnfdYSeleccionarNoSiVacio(page);
  if (hizoApnfd) hizoValidaciones = true;

  const hizoTiempoVivienda = await asegurarTiempoEnVivienda(page, "0").catch(() => false);
  if (hizoTiempoVivienda) hizoValidaciones = true;

  const labelNivelEstudio = page.locator('text="Nivel de estudio"').first();
  const visibleNivelEstudio = await labelNivelEstudio.isVisible().catch(() => false);
  if (visibleNivelEstudio) {
    await clickReintentarListaSiVisible(page, "Nivel de estudio");
    await seleccionarDefaultSiVacio(page, "Nivel de estudio", 0).catch(() => { });
    hizoValidaciones = true;
  }

  const labelTipoVivienda = page.locator('text="Tipo de vivienda"').first();
  const visibleTipoVivienda = await labelTipoVivienda.isVisible().catch(() => false);
  if (visibleTipoVivienda) {
    await clickReintentarListaSiVisible(page, "Tipo de vivienda");
    await seleccionarDefaultSiVacio(page, "Tipo de vivienda", 0).catch(() => { });
    hizoValidaciones = true;
  }

  const seccionReferencia = page
    .locator('xpath=//fieldset[.//*[contains(normalize-space(.),"Referencia personal")]]')
    .first();
  const visibleReferencia = await seccionReferencia.isVisible().catch(() => false);
  if (visibleReferencia) {
    const sinResultadosRef = seccionReferencia.getByText(/No se han encontrado resultados/i).first();
    const noTieneResultados = await sinResultadosRef.isVisible().catch(() => false);
    if (noTieneResultados) {
      const btnAddReferencia = seccionReferencia
        .getByRole('button', { name: /A?adir referencia|Anadir referencia/i })
        .first();
      await btnAddReferencia.waitFor({ state: 'visible', timeout: 10000 });
      await btnAddReferencia.click();
      await llenarCampoPorLabel(page, "Nombres y apellidos", "Juan Perez");
      await llenarInputMask(page, "Teléfono", "8092551233");
      await llenarCampoPorLabel(page, "Relación", "Amigo");
      await page.getByRole('button', { name: 'Aceptar' }).click();
      const btnGestionRef = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
      await btnGestionRef.waitFor({ state: 'visible', timeout: 10000 });
      await btnGestionRef.click();
      await cargarDocumentoEnGestionDocumental(page, path.resolve(process.cwd(), 'data', 'Documentos', 'evidencia.pdf'));
      hizoValidaciones = true;
    }
  }

  if (hizoValidaciones) {
    // No avanzar aqui: el paso de productos controla cuando pulsar Continuar.
    await page.waitForTimeout(300);
  }
}

async function etapaSeccionProductos(page: Page, registro: RegistroExcel) {
  let { seccionProductos } = await asegurarSeccionProductosVisible(page, {
    maxIntentos: 10,
    waitInicialMs: 90000,
    waitLoopMs: 8000,
    scrollPauseMs: 150,
    step2WaitMs: 14000,
    locateSection: localizarSeccionProductos,
    waitForUpdating: esperarFinActualizandoSolicitud,
    closeCancelModal: async (currentPage) => cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false),
    beforeStart: async (currentPage) => resolverCorreoPredeterminadoYContinuarSiVisible(currentPage).catch(() => false),
    isPreProductos: estaEnPantallaPreProductos,
    onPreProductos: async (currentPage) => {
      await asegurarTiempoEnVivienda(currentPage, "0").catch(() => false);
      await resolverNoPoseeCorreoSiFalta(currentPage).catch(() => false);
      const btnContinuar = getBotonContinuar(currentPage);
      const continuarVisible = await btnContinuar.isVisible().catch(() => false);
      const continuarEnabled = continuarVisible
        ? await btnContinuar.isEnabled().catch(() => false)
        : false;
      if (continuarVisible && continuarEnabled) {
        await btnContinuar.click({ force: true }).catch(() => { });
        await esperarFinActualizandoSolicitud(currentPage, 18000).catch(() => false);
        await resolverNoPoseeCorreoSiFalta(currentPage).catch(() => false);
        await currentPage.waitForTimeout(280);
      }
    },
    onFailure: async (currentPage, context) => {
      const urlActual = currentPage.url();
      const enPaso3SinProducto = await context.alertSinProductosPaso.isVisible().catch(() => false);
      const dropdownsVisibles = await currentPage
        .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
        .count()
        .catch(() => 0);
      const enPantallaPrevia = await estaEnPantallaPreProductos(currentPage).catch(() => false);
      const actualizandoVisible = await currentPage.getByText(/Actualizando solicitud/i).first().isVisible().catch(() => false);
      console.log(`[NAV-PRODUCTO] categoriaVisible=false url=${urlActual} paso3SinProducto=${enPaso3SinProducto} dropdownsVisibles=${dropdownsVisibles} prePantalla=${enPantallaPrevia} actualizando=${actualizandoVisible}`);
      return null;
    },
  });

  await seleccionarCategoriaEnSeccionProductos(page, seccionProductos);
  await seccionProductos.waitFor({ state: "visible", timeout: 10000 }).catch(() => { });
  await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta);

  const msgSinProductos = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
  const confirmarProductoAgregado = async () => {
    if (await modalProductoConfigVisible(page)) {
      return false;
    }
    const enUI = await detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta);
    if (enUI) return true;
    const sinProductos = await msgSinProductos.isVisible().catch(() => false);
    if (sinProductos) return false;
    await page.waitForTimeout(900);
    if (await modalProductoConfigVisible(page)) {
      return false;
    }
    return detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta);
  };

  const clickAgregarProductoDesdeBalance = async () => {
    const modalProducto = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
      .first();
    const modalVisible = await modalProducto.isVisible().catch(() => false);
    if (modalVisible) {
      const candidatosModal = [
        modalProducto.locator('xpath=//*[contains(normalize-space(.),"Balance promedio")]/following::button[contains(@class,"p-button")][1]').first(),
        modalProducto.locator('.p-inputgroup > .p-button:visible, .p-inputgroup > button:visible').first(),
        modalProducto.locator('button:has(.pi-plus):visible, button:has(i.pi-plus):visible').first(),
      ];
      for (const btn of candidatosModal) {
        const visible = await btn.isVisible().catch(() => false);
        if (!visible) continue;
        const txt = ((await btn.innerText().catch(() => "")) || "").trim();
        if (/^cancelar$|^aceptar$|^continuar$|^guardar$|^finalizar$/i.test(txt)) continue;
        await btn.scrollIntoViewIfNeeded().catch(() => { });
        await btn.click({ force: true }).catch(() => { });
        return true;
      }
      // Hay productos donde el modal no requiere boton de agregar intermedio.
      return true;
    }

    const bloqueBalance = page
      .locator('xpath=//*[contains(normalize-space(.),"Balance promedio")]/ancestor::*[self::div or self::fieldset][1]')
      .first();
    const bloqueVisible = await bloqueBalance.isVisible().catch(() => false);
    if (!bloqueVisible) return false;

    const candidatos = [
      bloqueBalance.locator('.p-inputgroup > .p-button:visible, .p-inputgroup > button:visible').first(),
      bloqueBalance.locator('button:has(.pi-plus):visible, button:has(i.pi-plus):visible').first(),
      bloqueBalance.locator('xpath=.//*[contains(normalize-space(.),"Balance promedio")]/following::button[contains(@class,"p-button")][1]').first(),
    ];
    for (const btn of candidatos) {
      const visible = await btn.isVisible().catch(() => false);
      if (!visible) continue;
      const txt = ((await btn.innerText().catch(() => "")) || "").trim();
      if (/^cancelar$|^aceptar$|^continuar$|^guardar$|^finalizar$/i.test(txt)) continue;
      await btn.scrollIntoViewIfNeeded().catch(() => { });
      await btn.click({ force: true }).catch(() => { });
      return true;
    }
    return false;
  };

  const llenarBalancePromedioEnContexto = async () => {
    const valorObjetivo = "5000";
    console.log("[BalanceModal][ce-ex] Iniciando llenado de 'Balance promedio'.");
    const modalProducto = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
      .first();
    const modalVisible = await modalProducto.isVisible().catch(() => false);

    const encontrarInputBalanceModal = async () => {
      const candidatos = [
        modalProducto.locator('xpath=.//*[normalize-space(.)="Balance promedio"]/ancestor::div[contains(@class,"space-y") or contains(@class,"col")][1]//input[not(@type="hidden")][1]').first(),
        modalProducto.locator('xpath=.//*[contains(normalize-space(.),"Balance promedio")]/ancestor::div[contains(@class,"space-y") or contains(@class,"col")][1]//input[not(@type="hidden")][1]').first(),
        modalProducto.locator('label').filter({ hasText: /Balance promedio/i }).first().locator('xpath=following::input[1]').first(),
        modalProducto.locator('xpath=.//*[contains(normalize-space(.),"Balance promedio")]/following::input[1]').first(),
      ];
      for (const input of candidatos) {
        const visible = await input.isVisible().catch(() => false);
        if (!visible) continue;
        const placeholder = ((await input.getAttribute("placeholder").catch(() => "")) || "").trim();
        if (/calcular tasa/i.test(placeholder)) continue;
        return input;
      }

      const inputsVisibles = modalProducto.locator('input:visible');
      const total = await inputsVisibles.count().catch(() => 0);
      for (let i = 0; i < total; i++) {
        const input = inputsVisibles.nth(i);
        const editable = await input.isEditable().catch(() => false);
        if (!editable) continue;
        const placeholder = ((await input.getAttribute("placeholder").catch(() => "")) || "").trim();
        if (/calcular tasa/i.test(placeholder)) continue;
        return input;
      }
      return null;
    };

    const leerValorBalanceModal = async () => {
      const input = await encontrarInputBalanceModal();
      if (!input) return "";
      return (await input.inputValue().catch(() => "")).trim();
    };

    const llenarInputDirecto = async (input: any) => {
      const visible = await input.isVisible().catch(() => false);
      if (!visible) return false;
      const editable = await input.isEditable().catch(() => false);
      if (!editable) return false;
      await input.scrollIntoViewIfNeeded().catch(() => { });
      await input.click({ force: true }).catch(() => { });
      await input.press("Control+A").catch(() => { });
      await input.press("Delete").catch(() => { });
      await input.fill(valorObjetivo).catch(async () => {
        await input.pressSequentially(valorObjetivo, { delay: 20 }).catch(() => { });
      });
      await input.dispatchEvent("input").catch(() => { });
      await input.dispatchEvent("change").catch(() => { });
      await input.blur().catch(() => { });
      await page.waitForTimeout(150);
      let valorFinal = (await input.inputValue().catch(() => "")).trim();
      if (!valorFinal) {
        await input.evaluate((el: HTMLInputElement, v: string) => {
          const inputEl = el instanceof HTMLInputElement ? el : null;
          if (!inputEl) return;
          inputEl.value = String(v);
          inputEl.dispatchEvent(new Event("input", { bubbles: true }));
          inputEl.dispatchEvent(new Event("change", { bubbles: true }));
          inputEl.dispatchEvent(new Event("blur", { bubbles: true }));
        }, valorObjetivo).catch(() => { });
        await page.waitForTimeout(150);
        valorFinal = (await input.inputValue().catch(() => "")).trim();
      }
      return !!valorFinal;
    };

    if (modalVisible) {
      const inputBalance = await encontrarInputBalanceModal();
      if (!inputBalance) {
        console.log("[BalanceModal][ce-ex] No se encontro input editable para 'Balance promedio'.");
      }
      if (inputBalance && await llenarInputDirecto(inputBalance)) {
        console.log("[BalanceModal][ce-ex] Balance promedio llenado en input identificado del modal.");
        return true;
      }

      const candidatos = [
        modalProducto.locator('input.p-inputnumber-input:visible').first(),
        modalProducto.locator('input.p-inputtext:visible').first(),
      ];
      for (const input of candidatos) {
        if (await llenarInputDirecto(input)) {
          const valorModal = await leerValorBalanceModal();
          if (valorModal) {
            console.log("[BalanceModal][ce-ex] Balance promedio llenado por fallback de inputs visibles.");
            return true;
          }
        }
      }
    }

    await llenarInputNumber(page, "Balance promedio", valorObjetivo).catch(() => { });
    if (modalVisible) {
      const valorModal = await leerValorBalanceModal();
      if (!valorModal) {
        console.log("[BalanceModal][ce-ex] El valor de 'Balance promedio' sigue vacio tras intentos de llenado.");
      }
      if (valorModal) return true;
      return false;
    }

    const valorGlobal = (await page
      .locator('label')
      .filter({ hasText: /Balance promedio/i })
      .first()
      .locator('xpath=following::input[1]')
      .first()
      .inputValue()
      .catch(() => "")).trim();
    return !!valorGlobal;
  };

  const clickCalcularTasaEnContexto = async () => {
    const modalProducto = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
      .first();
    const modalVisible = await modalProducto.isVisible().catch(() => false);
    const root = modalVisible ? modalProducto : page;

    const candidatos = [
      root.locator('button:has([class*="ph--calculator"])').first(),
      root.locator('button:has([class*="pi-calculator"])').first(),
      root.locator('button:has([class*="calculator"])').first(),
      root.locator('xpath=.//*[contains(normalize-space(.),"Tasa")]/ancestor::*[self::div or self::fieldset][1]//button[contains(@class,"p-button")][1]').first(),
      root.locator('xpath=.//input[contains(@placeholder,"Calcular tasa")]/following::button[contains(@class,"p-button")][1]').first(),
    ];
    for (const btn of candidatos) {
      const visible = await btn.isVisible().catch(() => false);
      if (!visible) continue;
      const txt = ((await btn.innerText().catch(() => "")) || "").trim();
      if (/^cancelar$|^aceptar$|^continuar$|^guardar$|^finalizar$/i.test(txt)) continue;
      await btn.scrollIntoViewIfNeeded().catch(() => { });
      await btn.click({ force: true }).catch(() => { });
      return true;
    }
    return false;
  };

  let productoAgregado = false;
  for (let intentoProducto = 1; intentoProducto <= 3; intentoProducto++) {
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    const yaAgregadoAntes = await confirmarProductoAgregado();
    if (yaAgregadoAntes) {
      productoAgregado = true;
      break;
    }
    console.log(`[BalanceModal][ce-ex] Intento de llenado ${intentoProducto}/3.`);
    const balanceLleno = await llenarBalancePromedioEnContexto();
    if (!balanceLleno) {
      if (intentoProducto < 3) {
        await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
        await page.waitForTimeout(600);
        continue;
      }
      throw new Error("[CRITICO] No se pudo llenar 'Balance promedio' en el modal de producto.");
    }
    const clicCalcularTasa = await clickCalcularTasaEnContexto();
    if (clicCalcularTasa) {
      await page.waitForTimeout(2000);
    }
    const clicAgregar = await clickAgregarProductoDesdeBalance();
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    if (!clicAgregar) {
      if (intentoProducto < 3) {
        await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
        await page.waitForTimeout(600);
        continue;
      }
      throw new Error(`[CRITICO] No se pudo encontrar boton para agregar producto en 'Balance promedio'.`);
    }
    await page.waitForTimeout(1500);
    const modalProducto = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
      .first();
    const btnAceptarProducto = (await modalProducto.isVisible().catch(() => false))
      ? modalProducto.getByRole('button', { name: /^Aceptar$/i }).first()
      : page.getByRole('button', { name: /^Aceptar$/i }).first();
    if (await btnAceptarProducto.isVisible().catch(() => false)) {
      await btnAceptarProducto.click({ force: true }).catch(() => { });
    }
    await page.waitForTimeout(1200);
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

    const confirmoAhora = await confirmarProductoAgregado();
    if (confirmoAhora) {
      productoAgregado = true;
      break;
    }

    if (intentoProducto < 3) {
      await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
      await page.waitForTimeout(600);
    }
  }

  if (!productoAgregado) {
    const sinProductosFinal = await msgSinProductos.isVisible().catch(() => false);
    if (sinProductosFinal) {
      throw new Error(`[CRITICO] No se agrego el producto '${registro.tipoCuenta}' en la simulacion.`);
    }
    throw new Error(`[CRITICO] No se detecta el producto agregado en pantalla: '${registro.tipoCuenta}'.`);
  }
}
async function etapaRelacionadosYAsociacion(page: Page, registro: RegistroExcel) {
  await agregarRelacionadoSiAplica(page, registro);
  const verificacionBpm = await abrirBpmSiVerificacionConoceCliente(page);
  if (verificacionBpm?.mpn) {
    console.log(`[Verificacion] Solicitud para BPM: ${verificacionBpm.mpn}`);
  }

  const msgSinProductosDespuesContinuar = page.getByText(/No se agregaron productos en simulaci[oó]n/i).first();
  const labelPropositoTaller = page.getByText(/Prop[oó]sito/i).first();
  const seccionProductos = await localizarSeccionProductos(page);
  const sinProductosAntesContinuar = await msgSinProductosDespuesContinuar.isVisible().catch(() => false);
  const productoEnUIPrevio = await detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta).catch(() => false);
  if (sinProductosAntesContinuar || !productoEnUIPrevio) {
    throw new Error(`[CRITICO] Bloqueado 'Continuar': aun no hay productos agregados para '${registro.tipoCuenta}'.`);
  }

  await esperarFinActualizandoSolicitud(page, 15000).catch(() => false);
  await getBotonContinuar(page).click();
  const llegoATaller = await labelPropositoTaller
    .waitFor({ state: 'visible', timeout: 45000 })
    .then(() => true)
    .catch(() => false);
  if (!llegoATaller) {
    const sinProductosTrasContinuar = await msgSinProductosDespuesContinuar.isVisible().catch(() => false);
    const productoEnUITrasContinuar = await detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta).catch(() => false);
    if (sinProductosTrasContinuar || !productoEnUITrasContinuar) {
      throw new Error(`[CRITICO] Se pulso Continuar sin productos agregados para '${registro.tipoCuenta}'.`);
    }
    throw new Error("No se pudo abrir Taller de productos despues de Continuar.");
  }

  return labelPropositoTaller;
}

async function etapaTallerProductos(
  page: Page,
  registro: RegistroExcel,
  labelPropositoTaller: Locator,
  capturasRef: string[]
) {
  await labelPropositoTaller.waitFor({ state: "visible", timeout: 60000 });
  await clickReintentarListaSiVisible(page, /Prop[oó]sito/i);
  await seleccionarDropdownPorCampo(page, /Prop[oó]sito/i, 1);
  await llenarInputNumber(page, "Monto inicial", "5000");
  await seleccionarInstrumentoRobusto(page);
  await llenarInputNumber(page, /Rango m[aá]ximo/i, "1000000");
  await llenarInputNumber(page, /Dep[oó]sito en efectivo/i, "5000");

  const inputRangoMaximo = page
    .locator("label")
    .filter({ hasText: /Rango m[aá]ximo/i })
    .first()
    .locator('xpath=following::input[1]')
    .first();
  const inputDepositoEfectivo = page
    .locator("label")
    .filter({ hasText: /Dep[oó]sito en efectivo/i })
    .first()
    .locator('xpath=following::input[1]')
    .first();

  let valorRango = (await inputRangoMaximo.inputValue().catch(() => "")).trim();
  let valorDeposito = (await inputDepositoEfectivo.inputValue().catch(() => "")).trim();
  if (!valorRango) {
    await llenarInputNumber(page, /Rango m[aá]ximo/i, "1000000");
    valorRango = (await inputRangoMaximo.inputValue().catch(() => "")).trim();
  }
  if (!valorDeposito) {
    await llenarInputNumber(page, /Dep[oó]sito en efectivo/i, "5000");
    valorDeposito = (await inputDepositoEfectivo.inputValue().catch(() => "")).trim();
  }
  if (!valorRango || !valorDeposito) {
    throw new Error(`Campos requeridos vacios en Taller. Rango='${valorRango}' Deposito='${valorDeposito}'`);
  }

  const btnCrearCuenta = page.getByRole('button', { name: /Crear cuenta/i }).first();
  await btnCrearCuenta.waitFor({ state: 'visible', timeout: 15000 });
  const crearEnabled = await btnCrearCuenta.isEnabled().catch(() => false);
  if (!crearEnabled) {
    throw new Error("El boton 'Crear cuenta' esta deshabilitado.");
  }
  await btnCrearCuenta.click();

  const creandoTxt = page.getByText(/Creando cuenta en efectivo/i).first();
  const toastCuentaOk = page.getByText(/Informaci[oó]n guardada|Cuenta .* generad/i).first();
  const modalFinalizadaPreview = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Solicitud finalizada/i })
    .first();

  const vioCreando = await creandoTxt.waitFor({ state: 'visible', timeout: 15000 })
    .then(() => true)
    .catch(() => false);
  if (vioCreando) {
    await creandoTxt.waitFor({ state: 'hidden', timeout: 180000 });
  }

  const confirmoCreacion = await Promise.race([
    toastCuentaOk.waitFor({ state: 'visible', timeout: 90000 }).then(() => true).catch(() => false),
    modalFinalizadaPreview.waitFor({ state: 'visible', timeout: 90000 }).then(() => true).catch(() => false),
  ]);
  if (!confirmoCreacion) {
    throw new Error("[CRITICO] No se confirmo la creacion de la cuenta en efectivo.");
  }

  // Evidencia inmediata tras crear cuenta (antes de continuar/finalizar el flujo).
  const evidenciaPostCrear = await capturarCuentaComoPNG(
    page,
    registro.identificacion,
    `${registro.tipoCuenta}__post-crear`
  );
  capturasRef.push(evidenciaPostCrear);

  const finalizada = await finalizarSolicitudSiPendiente(page, registro, {
    timeoutMs: FINALIZACION_TIMEOUT_MS,
    maxIntentos: FINALIZACION_MAX_INTENTOS,
  });
  if (!finalizada) {
    const cerroModalUltimaOportunidad = await cerrarModalSolicitudFinalizada(page, {
      timeoutMs: 15000,
      waitForAppearMs: 12000,
    }).catch(() => false);
    if (cerroModalUltimaOportunidad) {
      return;
    }
    console.log(`[Finalizacion][${registro.identificacion}] No se pudo confirmar finalizacion tras ${FINALIZACION_MAX_INTENTOS} intentos.`);
    throw new Error(`[NO_FINALIZADA] No se pudo confirmar finalizacion para '${registro.identificacion}'.`);
  }

  const modalFinalizadaSigueVisible = await page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i })
    .last()
    .isVisible()
    .catch(() => false);
  if (modalFinalizadaSigueVisible) {
    console.log(`[WARN][${registro.identificacion}] Se omite screenshot '__2' porque el modal 'Solicitud finalizada' aun esta visible.`);
  } else {
    const evidenciaPng2 = await capturarCuentaComoPNG(page, registro.identificacion, `${registro.tipoCuenta}__2`);
    capturasRef.push(evidenciaPng2);
  }
}

async function etapaEvidencias(page: Page, registro: RegistroExcel, capturasRef: string[]) {
  await cerrarModalSolicitudFinalizada(page, { timeoutMs: 1200 }).catch(() => false);
  const badgePendiente = await page.getByText(/Pendiente finalizar/i).first().isVisible().catch(() => false);
  if (badgePendiente) {
    console.log(`[WARN][${registro.identificacion}] Se omite screenshot final porque la solicitud sigue en estado 'Pendiente finalizar'.`);
    return;
  }
  const modalFinalizadaVisible = await page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i })
    .last()
    .isVisible()
    .catch(() => false);
  if (modalFinalizadaVisible) {
    console.log(`[WARN][${registro.identificacion}] Se omite screenshot final porque el modal 'Solicitud finalizada' sigue visible.`);
    return;
  }
  const evidenciaPng = await capturarCuentaComoPNG(page, registro.identificacion, registro.tipoCuenta);
  capturasRef.push(evidenciaPng);
}

async function prepararSiguienteRegistro(page: Page) {
  if (page.isClosed()) return;

  const cerroFinalizada = await cerrarModalSolicitudFinalizada(page, { timeoutMs: 1200 }).catch(() => false);
  if (!cerroFinalizada) {
    const modalFinalizada = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i })
      .last();
    if (await modalFinalizada.isVisible().catch(() => false)) {
      const btnCerrar = modalFinalizada.getByRole('button', { name: /^Cerrar$/i }).first();
      if (await btnCerrar.isVisible().catch(() => false)) {
        await btnCerrar.click({ force: true }).catch(() => { });
        await modalFinalizada.waitFor({ state: 'hidden', timeout: 5000 }).catch(() => { });
      }
    }
  }
  await page.keyboard.press('Escape').catch(() => { });
  await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

  const modalVisible = page.locator('.p-dialog:visible, [role="dialog"]:visible').last();
  if (await modalVisible.isVisible().catch(() => false)) {
    const btnCerrarGenerico = modalVisible
      .getByRole('button', { name: /Cerrar|Aceptar|Cancelar/i })
      .first();
    if (await btnCerrarGenerico.isVisible().catch(() => false)) {
      await btnCerrarGenerico.click({ force: true }).catch(() => { });
    }
  }

  await page.goto('https://srvqacgowb01.local.bsc.com:5000/requests/create/multiproduct', {
    waitUntil: 'domcontentloaded',
  }).catch(() => { });
  await esperarPortalListoTrasLogin(page, { timeoutMs: 120000 }).catch(() => ({ loginDetectado: false }));
}



test('Cuenta Efectivo Cliente existente - desde Excel', async () => {
  const session = await launchPortalSession({ defaultContinueOnError: true });
  const { page } = session;
  let failedRegistros: Array<{ identificacion: string; error: string }> = [];

  try {
    const ids = leerRegistrosDesdeExcel('data/ids.xlsx');
    console.log("Identificaciones encontradas:", ids);
    limpiarEvidenciasTemporales();

    const runState = await runRegistros({
      page,
      registros: ids,
      browserSession: session,
      onRegistro: async (registro) => {
        // ===== INICIO: Flujo del registro =====
        await etapaFlujoRegistro(page, registro);

        // ===== INICIO: Validaciones previas =====
        await etapaValidacionesPrevias(page);

        // ===== INICIO: Seccion de productos =====
        await etapaSeccionProductos(page, registro);

        // ===== INICIO: Relacionados y asociacion =====
        const labelPropositoTaller = await etapaRelacionadosYAsociacion(page, registro);

        // ===== INICIO: Taller de productos =====
        await etapaTallerProductos(page, registro, labelPropositoTaller, capturas);

        // ===== INICIO: Evidencias =====
        await etapaEvidencias(page, registro, capturas);
        console.log(`[OK][${registro.identificacion}] Registro completado.`);
      },
      onRegistroError: async (registro, errorMsg, state) => {
        const esSkipCasoEnCurso = esErrorCasoEnCursoNoGestionable(errorMsg);
        if (esSkipCasoEnCurso) {
          const razon = "Caso en curso detectado en Bizagi; no fue posible cargar la fila MPN para cancelacion automatica.";
          state.skippedRegistros.push({ identificacion: registro.identificacion, razon });
          state.failedRegistros.pop();
          console.log(`[SKIP][${registro.identificacion}] ${razon}`);
          return 'continue';
        }

        const esNoFinalizada = /\[NO_FINALIZADA\]/.test(errorMsg);
        const pageCerrada = page.isClosed();
        if (session.continueOnRegistroError && !esNoFinalizada && !pageCerrada) {
          console.log(`[Loop][${registro.identificacion}] Continuando con el siguiente registro por PW_CONTINUE_ON_ERROR=1.`);
          return 'continue';
        }

        if (pageCerrada) {
          throw new Error(`[CRITICO] El navegador/pagina se cerro inesperadamente durante '${registro.identificacion}'.`);
        }
        return 'throw';
      },
      onRegistroFinally: async () => {
        if (!page.isClosed()) {
          await prepararSiguienteRegistro(page).catch(() => {});
        }
      },
    });
    failedRegistros = runState.failedRegistros;
    const { skippedRegistros } = runState;

    if (capturas.length) {
      const pdfPath = await unirPNGsEnUnPDF(capturas, "Evidencias_Cuentas.pdf");
      const abierto = await abrirArchivoAlFinal(pdfPath).catch(() => false);
      if (!abierto) {
        console.log(`[Evidencias] No se pudo abrir automaticamente el PDF: ${pdfPath}`);
      }
    }

    if (failedRegistros.length) {
      console.log("Registros con error:", failedRegistros);
    }
    if (skippedRegistros.length) {
      console.log("Registros omitidos (caso en curso):", skippedRegistros);
    }
  } finally {
    await closeBrowserSession(session, failedRegistros);
  }
});
