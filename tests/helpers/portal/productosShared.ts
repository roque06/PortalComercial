import type { Locator, Page } from '@playwright/test';

export interface ProductoUiTimeouts {
  panelQuickTimeout: number;
  panelTimeout: number;
  scopeRetryWaitMs: number;
}

export function escapeRegexText(value: string) {
  return String(value ?? '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function extraerCodigoProducto(tipoCuenta: string) {
  const exact = String(tipoCuenta ?? '').match(/^\s*(\d{3})\s*-/);
  if (exact?.[1]) return exact[1];
  const generic = String(tipoCuenta ?? '').match(/\b(\d{3})\b/);
  return generic?.[1] ?? '';
}

export function extraerNombreProducto(tipoCuenta: string) {
  return String(tipoCuenta ?? '').replace(/^\s*\d{3}\s*-\s*/, '').trim();
}

export function esValorDropdownVacio(value: string) {
  const txt = String(value ?? '').trim();
  return !txt || /^seleccione/i.test(txt) || /reintentar/i.test(txt);
}

export async function localizarSeccionProductos(page: Page, opts?: { useBroadFallback?: boolean }) {
  const candidatos: Locator[] = [
    page
      .locator('xpath=(//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÃÃ‰ÃÃ“ÃšÃ¡Ã©Ã­Ã³Ãº","AEIOUaeiou"),"Categoria de producto")] and .//*[contains(translate(normalize-space(.),"ÃÃ‰ÃÃ“ÃšÃ¡Ã©Ã­Ã³Ãº","AEIOUaeiou"),"Producto")]][1])')
      .first(),
    page
      .locator('xpath=(//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÃÃ‰ÃÃ“ÃšÃ¡Ã©Ã­Ã³ÃºÂ¿?","AEIOUaeiou"),"Que quiere el cliente")] and .//*[contains(translate(normalize-space(.),"ÃÃ‰ÃÃ“ÃšÃ¡Ã©Ã­Ã³Ãº","AEIOUaeiou"),"Producto")] and .//*[contains(translate(normalize-space(.),"ÃÃ‰ÃÃ“ÃšÃ¡Ã©Ã­Ã³Ãº","AEIOUaeiou"),"Relacionados")]][1])')
      .first(),
    page
      .locator('xpath=(//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÃÃ‰ÃÃ“ÃšÃ¡Ã©Ã­Ã³Ãº","AEIOUaeiou"),"Productos")] and .//*[contains(translate(normalize-space(.),"ÃÃ‰ÃÃ“ÃšÃ¡Ã©Ã­Ã³Ãº","AEIOUaeiou"),"Categoria de producto")]][1])')
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

  if (opts?.useBroadFallback) {
    const fallbackCandidatos = page.locator(
      'fieldset:visible, form > div:visible, .p-panel-content:visible, [data-pc-section="content"]:visible'
    );
    const total = await fallbackCandidatos.count().catch(() => 0);
    for (let i = 0; i < total; i++) {
      const el = fallbackCandidatos.nth(i);
      const txt = (await el.textContent().catch(() => '')) || '';
      if (!/product/i.test(txt) && !/categor/i.test(txt)) continue;
      const drops = await el.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').count().catch(() => 0);
      if (drops >= 2) return el;
    }

    const primerDropdown = page.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').first();
    const ancestro = primerDropdown.locator('xpath=ancestor::fieldset[1] | xpath=ancestor::form[1] | xpath=ancestor::div[@class][1]').first();
    if (await ancestro.isVisible().catch(() => false)) {
      return ancestro;
    }
  }

  return candidatos[0];
}

export async function seleccionarCategoriaEnSeccionProductos(
  page: Page,
  seccionProductos: Locator,
  timeouts: ProductoUiTimeouts,
  opts?: { categoriaRegex?: RegExp; onBeforeOpen?: (page: Page) => Promise<unknown> }
) {
  const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
  const count = await dropdowns.count().catch(() => 0);
  if (count < 1) {
    throw new Error("[CRITICO] No se encontro dropdown de 'Categoria de producto' en la seccion de Productos.");
  }

  const categoriaDropdown = dropdowns.first();
  const labelCategoria = categoriaDropdown.locator('.p-dropdown-label, [data-pc-section="label"]').first();
  let valor = ((await labelCategoria.textContent().catch(() => "")) || "").trim();
  if (!esValorDropdownVacio(valor)) return;

  const categoriaRegex = opts?.categoriaRegex ?? /Cuentas de Efectivo/i;

  for (let intento = 1; intento <= 4; intento++) {
    if (opts?.onBeforeOpen) {
      await opts.onBeforeOpen(page).catch(() => false);
    }

    await categoriaDropdown.scrollIntoViewIfNeeded().catch(() => { });
    await categoriaDropdown.click({ force: true }).catch(() => { });

    const combobox = categoriaDropdown.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
    let panel: Locator | null = null;
    if (panelId) {
      const byId = page.locator(`#${panelId}`);
      const visible = await byId.waitFor({ state: "visible", timeout: timeouts.panelQuickTimeout })
        .then(() => true)
        .catch(() => false);
      if (visible) panel = byId;
    }
    if (!panel) {
      const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
      const visible = await fallback.waitFor({ state: "visible", timeout: timeouts.panelQuickTimeout })
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
        await page.waitForTimeout(timeouts.scopeRetryWaitMs);
      }

      const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
      const listo = await items.first().waitFor({ state: "visible", timeout: timeouts.panelTimeout })
        .then(() => true)
        .catch(() => false);
      const countItems = await items.count().catch(() => 0);
      if (listo && countItems > 0) {
        // Log categorías disponibles para diagnóstico
        for (let ci = 0; ci < Math.min(countItems, 10); ci++) {
          const ctxt = ((await items.nth(ci).innerText().catch(() => '')) || '').trim();
          console.log(`[Categoria][diag] item[${ci}]='${ctxt}'`);
        }
        console.log(`[Categoria][diag] buscando categoriaRegex=${categoriaRegex}`);
        const itemCategoria = items.filter({ hasText: categoriaRegex }).first();
        if (await itemCategoria.isVisible().catch(() => false)) {
          console.log(`[Categoria][diag] match encontrado, clickeando`);
          await itemCategoria.click({ force: true }).catch(() => { });
        } else {
          console.log(`[Categoria][diag] sin match, clickeando primer item (fallback)`);
          await items.nth(0).click({ force: true }).catch(() => { });
        }
      }
    }

    await page.waitForTimeout(120);
    valor = ((await labelCategoria.textContent().catch(() => "")) || "").trim();
    if (!esValorDropdownVacio(valor)) return;

    await page.keyboard.press("Escape").catch(() => { });
    await page.waitForTimeout(timeouts.scopeRetryWaitMs);
  }

  throw new Error("[CRITICO] No se pudo seleccionar 'Categoria de producto' en la seccion de Productos.");
}

export async function leerValorDropdownEnScope(
  scope: Locator,
  labelText: string | RegExp,
  options?: { timeoutMs?: number }
) {
  const timeoutMs = options?.timeoutMs ?? 5000;
  const pattern = labelText instanceof RegExp
    ? labelText
    : new RegExp(escapeRegexText(String(labelText)), "i");

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

export async function detectarProductoAgregadoEnUI(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string,
  opts?: { useGlobalFallback?: boolean }
) {
  const tipoCuentaSafe = String(tipoCuenta ?? "").trim();
  if (!tipoCuentaSafe) return false;

  const tipoCuentaRegex = new RegExp(escapeRegexText(tipoCuentaSafe).replace(/\s+/g, "\\s+"), "i");
  const nombreProducto = extraerNombreProducto(tipoCuentaSafe);
  const nombreRegex = nombreProducto ? new RegExp(escapeRegexText(nombreProducto).replace(/\s+/g, "\\s+"), "i") : null;
  const codigoProducto = extraerCodigoProducto(tipoCuentaSafe);
  const codigoRegex = codigoProducto ? new RegExp(`\\b${escapeRegexText(codigoProducto)}\\b`, "i") : null;

  const msgSinProductos = page.getByText(/No se agregaron productos en simulaci[oÃ³]n/i).first();
  if (await msgSinProductos.isVisible().catch(() => false)) return false;

  const visibleTextoCompleto = await seccionProductos.getByText(tipoCuentaRegex).first().isVisible().catch(() => false);
  if (visibleTextoCompleto) return true;

  if (nombreRegex) {
    const visibleNombre = await seccionProductos.getByText(nombreRegex).first().isVisible().catch(() => false);
    if (visibleNombre) return true;
  }

  if (codigoRegex) {
    const visibleCodigo = await seccionProductos.getByText(codigoRegex).first().isVisible().catch(() => false);
    if (visibleCodigo) return true;
  }

  if (!opts?.useGlobalFallback) {
    const textoSeccion = ((await seccionProductos.innerText().catch(() => '')) || '').trim();
    return Boolean(
      textoSeccion &&
      (tipoCuentaRegex.test(textoSeccion) || (nombreRegex && nombreRegex.test(textoSeccion)) || (codigoRegex && codigoRegex.test(textoSeccion)))
    );
  }

  const selectorTextoProductoAgregado =
    'xpath=.//*[' +
    'not(ancestor::*[contains(@class,"p-dropdown") or contains(@class,"p-dropdown-panel") or @role="listbox" or @role="option"])' +
    ' and not(descendant::*[contains(@class,"p-dropdown") or contains(@class,"p-dropdown-panel") or @role="listbox" or @role="option"])' +
    ' and (self::span or self::td or self::p or self::small or self::strong or self::a or self::li or self::div)' +
    ']';

  const textosFueraDeDropdown = seccionProductos.locator(selectorTextoProductoAgregado);
  const textosFueraDeDropdownGlobal = page.locator(selectorTextoProductoAgregado);

  if (await textosFueraDeDropdown.filter({ hasText: tipoCuentaRegex }).first().isVisible().catch(() => false)) return true;
  if (await textosFueraDeDropdownGlobal.filter({ hasText: tipoCuentaRegex }).first().isVisible().catch(() => false)) return true;

  if (nombreRegex) {
    if (await textosFueraDeDropdown.filter({ hasText: nombreRegex }).first().isVisible().catch(() => false)) return true;
    if (await textosFueraDeDropdownGlobal.filter({ hasText: nombreRegex }).first().isVisible().catch(() => false)) return true;
  }

  if (codigoRegex) {
    if (await textosFueraDeDropdown.filter({ hasText: codigoRegex }).first().isVisible().catch(() => false)) return true;
    if (await textosFueraDeDropdownGlobal.filter({ hasText: codigoRegex }).first().isVisible().catch(() => false)) return true;
  }

  return false;
}

export async function modalProductoConfigVisible(page: Page, regex: RegExp) {
  const modalProducto = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: regex })
    .first();
  return modalProducto.isVisible().catch(() => false);
}

export async function confirmarSeleccionProductoRapida(
  page: Page,
  seccionProductos: Locator,
  tipoCuentaRegex: RegExp,
  codigoRegex: RegExp | null,
  modalRegex: RegExp,
  waitMs: number
) {
  const inicio = Date.now();
  while (Date.now() - inicio < waitMs) {
    if (await modalProductoConfigVisible(page, modalRegex)) return true;

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
      if (coincideTexto || coincideCodigo) {
        if (await modalProductoConfigVisible(page, modalRegex)) return true;
      }
    }

    await page.waitForTimeout(110);
  }
  return false;
}
