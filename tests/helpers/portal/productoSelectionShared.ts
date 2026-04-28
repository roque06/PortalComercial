import type { Locator, Page } from '@playwright/test';

export interface ProductoSelectionTimeouts {
  listLabelTimeout: number;
  listPanelTimeout: number;
  listPanelQuickTimeout: number;
  listScopeRetryWaitMs: number;
}

async function cerrarModalCancelarProcesoSiVisible(page: Page) {
  const modalCancelarProceso = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /(Cancelar|Salir) del proceso de solicitud/i })
    .first();
  const visible = await modalCancelarProceso.isVisible().catch(() => false);
  if (!visible) return false;

  const btnMantenerProceso = modalCancelarProceso.getByRole('button', { name: /^Cancelar$/i }).first();
  const btnCerrar = modalCancelarProceso
    .locator('.p-dialog-header-close, .p-dialog-titlebar-close, [aria-label*="close" i], [aria-label*="cerr" i]')
    .first();

  if (await btnMantenerProceso.isVisible().catch(() => false)) {
    await btnMantenerProceso.click({ force: true }).catch(() => {});
  } else if (await btnCerrar.isVisible().catch(() => false)) {
    await btnCerrar.click({ force: true }).catch(() => {});
  } else {
    await page.keyboard.press('Escape').catch(() => {});
  }

  await modalCancelarProceso.waitFor({ state: 'hidden', timeout: 4000 }).catch(() => {});
  return true;
}

async function resolvePanel(page: Page, field: Locator, timeouts: ProductoSelectionTimeouts) {
  const combobox = field.locator('[role="combobox"]').first();
  const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
  let panel: Locator | null = null;
  if (panelId) {
    const byId = page.locator(`#${panelId}`);
    const byIdVisible = await byId.waitFor({
      state: "visible",
      timeout: Math.min(timeouts.listPanelTimeout, timeouts.listPanelQuickTimeout),
    }).then(() => true).catch(() => false);
    if (byIdVisible) panel = byId;
  }
  if (!panel) {
    const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible');
    const total = await fallback.count().catch(() => 0);
    if (total > 0) {
      // Pick the one that is closest to the field or just the first visible one
      panel = fallback.first();
    }
  }
  return panel;
}

export async function seleccionarDropdownEnScopePorTexto(
  page: Page,
  scope: Locator,
  labelText: string | RegExp,
  textoOpcion: RegExp,
  fallbackIndex: number,
  timeouts: ProductoSelectionTimeouts,
  opts?: {
    onBeforeAttempt?: (page: Page) => Promise<unknown>;
    onAfterDropdownClick?: (page: Page) => Promise<boolean | unknown>;
  }
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
    await page.waitForTimeout(timeouts.listScopeRetryWaitMs);
    return true;
  };

  for (let intento = 1; intento <= 3; intento++) {
    if (opts?.onBeforeAttempt) {
      await opts.onBeforeAttempt(page).catch(() => false);
    }
    await clickReintentarScope().catch(() => false);

    let label = scope.locator("label").filter({ hasText: pattern }).first();
    if ((await label.count().catch(() => 0)) === 0) {
      label = scope.getByText(pattern).first();
    }
    const labelVisible = await label.waitFor({ state: "visible", timeout: timeouts.listLabelTimeout })
      .then(() => true)
      .catch(() => false);
    if (!labelVisible) continue;

    let field = label
      .locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]')
      .first();
    if (!(await field.isVisible().catch(() => false))) {
      field = label
        .locator('xpath=following::*[self::div[contains(@class,"p-dropdown")] or @data-pc-name="dropdown"][1]')
        .first();
    }
    const fieldVisible = await field.waitFor({ state: "visible", timeout: timeouts.listLabelTimeout })
      .then(() => true)
      .catch(() => false);
    if (!fieldVisible) continue;

    await field.scrollIntoViewIfNeeded().catch(() => { });
    const targets = [
      field.locator('.p-dropdown-trigger, [data-pc-section="trigger"]').first(),
      field.locator('[role="combobox"]').first(),
      field,
    ];

    let panel: Locator | null = null;
    for (const target of targets) {
      const visibleTarget = await target.isVisible().catch(() => false);
      if (!visibleTarget) continue;

      // Check if panel is already open before clicking
      panel = await resolvePanel(page, field, { ...timeouts, listPanelQuickTimeout: 500 });
      if (!panel) {
        await target.click({ force: true }).catch(() => {});
        panel = await resolvePanel(page, field, timeouts);
      }
      
      if (panel) break;
    }

    if (!panel) {
      await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
      await page.keyboard.press("Escape").catch(() => { });
      await page.waitForTimeout(timeouts.listScopeRetryWaitMs);
      continue;
    }

    const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
    const count = await items.count().catch(() => 0);
    const prefix = String(labelText).includes('Producto') ? '[SeccionProductos]' : '[Dropdown]';
    console.log(`${prefix} Opciones disponibles: ${count} para '${String(labelText)}'`);

    const itemByText = items.filter({ hasText: textoOpcion }).first();
      
    // WAIT for the item to render, as the portal is an SPA and might be waiting for the backend API to return the options.
    const isReady = await itemByText.waitFor({ state: "visible", timeout: timeouts.listPanelTimeout })
      .then(() => true)
      .catch(() => false);
    
    console.log(`${prefix} listo=${isReady} para '${String(textoOpcion)}'`);
      
    if (isReady) {
      await itemByText.scrollIntoViewIfNeeded().catch(() => { });
      
      let selected = false;
      for (let clickIntento = 1; clickIntento <= 3; clickIntento++) {
          // Try standard click first to ensure events fire correctly
          let clicked = await itemByText.click({ timeout: 3000 }).then(() => true).catch(() => false);
          if (!clicked) {
              clicked = await itemByText.click({ force: true, timeout: 2000 }).then(() => true).catch(() => false);
          }
          
          if (!clicked) {
              const handle = await itemByText.elementHandle().catch(() => null);
              if (handle) {
                  await page.evaluate(el => {
                      const e = el as HTMLElement;
                      e.click();
                      e.dispatchEvent(new Event('click', { bubbles: true }));
                      e.dispatchEvent(new Event('change', { bubbles: true }));
                  }, handle).catch(() => {});
              }
          }

          // Verify the panel is hiding or hidden
          await page.waitForTimeout(200);
          const stillVisible = await panel.isVisible().catch(() => false);
          if (!stillVisible) {
              selected = true;
              break;
          }
      }
      
      if (selected) return;
    }

    if (fallbackIndex >= 0 && count > fallbackIndex) {
      const fallbackItem = items.nth(fallbackIndex);
      console.log(`${prefix}[WARN] Usando fallback por indice ${fallbackIndex}`);
      await fallbackItem.click({ force: true }).catch(async () => {
          const handle = await fallbackItem.elementHandle().catch(() => null);
          if (handle) await page.evaluate(el => (el as HTMLElement).click(), handle).catch(() => {});
      });
      return;
    }

    await page.mouse.click(10, 10).catch(() => { });
    await page.waitForTimeout(timeouts.listScopeRetryWaitMs);
  }

  throw new Error(`No se pudo seleccionar opcion '${String(textoOpcion)}' en '${String(labelText)}'.`);
}

export async function seleccionarProductoPorDropdownSecundario(
  page: Page,
  seccionProductos: Locator,
  tipoCuentaRegex: RegExp,
  codigoRegex: RegExp | null,
  timeouts: ProductoSelectionTimeouts,
  opts?: { onAfterDropdownClick?: (page: Page) => Promise<boolean | unknown> }
) {
  const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
  const count = await dropdowns.count().catch(() => 0);
  console.log(`[ProductoFallback] dropdowns visibles en seccion=${count}`);
  if (count < 2) return false;

  const productoDropdown = dropdowns.nth(1);
  await productoDropdown.scrollIntoViewIfNeeded().catch(() => { });

  const abrirProductoDropdown = async () => {
    const targets = [
      productoDropdown.locator('.p-dropdown-trigger, [data-pc-section="trigger"]').first(),
      productoDropdown.locator('[role="combobox"]').first(),
      productoDropdown,
    ];

    for (const target of targets) {
      const visible = await target.isVisible().catch(() => false);
      if (!visible) continue;

      const clicked = await target.click({ force: true, timeout: 1800 }).then(() => true).catch(() => false);
      if (!clicked) continue;

      if (opts?.onAfterDropdownClick) {
        const shouldRetryOpen = await opts.onAfterDropdownClick(page).catch(() => false);
        if (shouldRetryOpen) {
          console.log('[ProductoFallback] Click abrio modal de cancelar. Se cierra y se reintenta abrir dropdown.');
          await page.waitForTimeout(120);
          continue;
        }
      }

      const panelIntento = await resolvePanel(page, productoDropdown, timeouts);
      if (panelIntento) return panelIntento;
    }

    return null;
  };

  const panel = await abrirProductoDropdown();
  if (!panel) return false;

  const btnRetryPanel = panel
    .locator('button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button.p-button-warning')
    .first();
  if (await btnRetryPanel.isVisible().catch(() => false)) {
    await btnRetryPanel.click({ force: true }).catch(() => { });
    await page.waitForTimeout(timeouts.listScopeRetryWaitMs);
  }

  const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
  const hasItems = await items.first().waitFor({ state: 'visible', timeout: timeouts.listPanelTimeout })
    .then(() => true)
    .catch(() => false);
  const countItems = await items.count().catch(() => 0);
  console.log(`[ProductoFallback] panel visible, items=${countItems}`);
  if (!hasItems) return false;

  // Log items para diagnóstico
  for (let i = 0; i < Math.min(countItems, 15); i++) {
    const txt = ((await items.nth(i).innerText().catch(() => '')) || '').trim().replace(/\n/g, ' ');
    console.log(`[ProductoFallback] item[${i}]='${txt}'`);
  }
  console.log(`[ProductoFallback] buscando tipoCuentaRegex=${tipoCuentaRegex} codigoRegex=${codigoRegex}`);

  const clickMatchingItem = async (target: Locator) => {
    if (!(await target.isVisible().catch(() => false))) return false;
    await target.scrollIntoViewIfNeeded().catch(() => { });
    const clicked = await target.click({ force: true, timeout: 2500 }).then(() => true).catch(() => false);
    if (!clicked) {
      const handle = await target.elementHandle().catch(() => null);
      if (handle) {
        await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => { });
        return true;
      }
      return false;
    }
    return true;
  };

  const byTipo = items.filter({ hasText: tipoCuentaRegex }).first();
  if (await clickMatchingItem(byTipo)) { console.log(`[ProductoFallback] match por tipoCuentaRegex`); return true; }

  if (codigoRegex) {
    const byCodigo = items.filter({ hasText: codigoRegex }).first();
    if (await clickMatchingItem(byCodigo)) { console.log(`[ProductoFallback] match por codigoRegex`); return true; }
  }

  console.log(`[ProductoFallback] sin match, retornando false`);
  return false;
}
