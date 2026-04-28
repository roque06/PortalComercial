import type { Locator, Page } from '@playwright/test';

export interface EnsureProductosSectionResult {
  categoriaVisible: boolean;
  seccionProductos: Locator;
}

export interface EnsureProductosSectionOptions {
  maxIntentos: number;
  waitInicialMs: number;
  waitLoopMs: number;
  scrollPauseMs: number;
  step2WaitMs: number;
  locateSection: (page: Page) => Promise<Locator>;
  waitForUpdating: (page: Page, timeoutMs: number) => Promise<unknown>;
  closeCancelModal?: (page: Page) => Promise<unknown>;
  beforeStart?: (page: Page) => Promise<unknown>;
  isPreProductos: (page: Page) => Promise<boolean>;
  onPreProductos?: (page: Page) => Promise<unknown>;
  resolveCategoriaVisible?: (
    page: Page,
    seccionProductos: Locator
  ) => Promise<{ categoriaVisible: boolean; seccionProductos?: Locator }>;
  onFailure?: (page: Page, context: {
    alertSinProductosPaso: Locator;
    labelPropositoPaso: Locator;
    seccionProductos: Locator;
  }) => Promise<Locator | null>;
}

export async function asegurarSeccionProductosVisible(
  page: Page,
  options: EnsureProductosSectionOptions
): Promise<EnsureProductosSectionResult> {
  await options.beforeStart?.(page).catch(() => false);

  const alertSinProductosPaso = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
  const labelPropositoPaso = page.getByText(/Prop(?:o|\u00f3)sito/i).first();
  const tituloProductos = page.getByText(/Productos/i).first();
  let seccionProductos = await options.locateSection(page);

  await options.waitForUpdating(page, options.waitInicialMs).catch(() => false);

  let categoriaVisible = false;
  for (let i = 1; i <= options.maxIntentos; i++) {
    await options.waitForUpdating(page, options.waitLoopMs).catch(() => false);
    await options.closeCancelModal?.(page).catch(() => false);

    await tituloProductos.scrollIntoViewIfNeeded().catch(() => {});
    await page.mouse.wheel(0, 600).catch(() => {});
    await page.waitForTimeout(options.scrollPauseMs);

    seccionProductos = await options.locateSection(page);
    if (options.resolveCategoriaVisible) {
      const result = await options.resolveCategoriaVisible(page, seccionProductos);
      categoriaVisible = result.categoriaVisible;
      if (result.seccionProductos) seccionProductos = result.seccionProductos;
    } else {
      const seccionVisible = await seccionProductos.isVisible().catch(() => false);
      const dropdownsEnSeccion = await seccionProductos
        .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
        .count()
        .catch(() => 0);
      categoriaVisible = seccionVisible && dropdownsEnSeccion >= 2;
    }
    if (categoriaVisible) break;

    const cayoEnPasoSinProducto = await alertSinProductosPaso.isVisible().catch(() => false);
    const yaEnTaller = await labelPropositoPaso.isVisible().catch(() => false);
    const enPantallaPrevia = await options.isPreProductos(page);

    if (enPantallaPrevia) {
      const handled = await options.onPreProductos?.(page).catch(() => false);
      if (handled !== false) {
          // If the handler didn't explicitly return false, we assume it tried to do something (like clicking Continuar).
          // We must re-evaluate the page state because the DOM likely changed dramatically.
          continue;
      }
    }

    if (cayoEnPasoSinProducto || yaEnTaller || enPantallaPrevia) {
      const step2 = page.locator(
        'li.p-steps-item:has-text("2"), [data-pc-section="menuitem"]:has-text("2"), .p-steps-item:has(.p-steps-number:has-text("2"))'
      ).first();
      if (await step2.isVisible().catch(() => false)) {
        await step2.click({ force: true }).catch(() => {});
        await options.closeCancelModal?.(page).catch(() => false);
        await options.waitForUpdating(page, options.step2WaitMs).catch(() => false);
        await page.waitForTimeout(Math.min(options.scrollPauseMs + 130, 700));
      }
    }
  }

  if (!categoriaVisible) {
    const fallbackSection = await options.onFailure?.(page, {
      alertSinProductosPaso,
      labelPropositoPaso,
      seccionProductos,
    });
    if (fallbackSection) {
      return { categoriaVisible: true, seccionProductos: fallbackSection };
    }

    throw new Error("[CRITICO] No se pudo visualizar 'Categoria de producto' para agregar el producto.");
  }

  return { categoriaVisible, seccionProductos };
}
