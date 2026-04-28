import type { Locator, Page } from '@playwright/test';

export interface ProductoDetectionOptions {
  escapeRegexText: (value: string) => string;
  extraerCodigoProducto: (tipoCuenta: string) => string;
  useGlobalFallback?: boolean;
  modalRegex?: RegExp;
  modalMatcher?: (dialogs: Locator, tipoCuentaRegex: RegExp, codigoRegex: RegExp | null) => Promise<boolean>;
}

const GLOBAL_TEXT_SELECTOR =
  'xpath=//*[not(ancestor::*[contains(@class,"p-dropdown-panel") or @role="listbox" or @role="option"]) and (self::div or self::span or self::p or self::small or self::strong or self::a or self::li or self::td)]';

export async function detectarProductoSeleccionadoEnUI(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string,
  options: ProductoDetectionOptions
) {
  const tipoCuentaSafe = String(tipoCuenta ?? '').trim();
  if (!tipoCuentaSafe) return false;

  const escapedText = options.escapeRegexText(tipoCuentaSafe).replace(/\s+/g, '\\s+');
  const tipoCuentaRegex = new RegExp(escapedText, 'i');
  
  const codigoProducto = options.extraerCodigoProducto(tipoCuentaSafe);
  const codigoRegex = codigoProducto ? new RegExp(`\\b${options.escapeRegexText(codigoProducto)}\\b`, 'i') : null;

  const checkLocator = async (loc: Locator, regex: RegExp) => {
      const visible = await loc.filter({ hasText: regex }).first().isVisible().catch(() => false);
      if (visible) return true;
      
      // Fallback: check all texts in case of weird whitespace/nbsp
      const allTexts = await loc.allInnerTexts().catch(() => []);
      return allTexts.some(t => regex.test(t.replace(/\u00a0/g, ' ')));
  };

  if (await checkLocator(seccionProductos, tipoCuentaRegex)) return true;

  if (options.useGlobalFallback) {
    if (await checkLocator(page.locator(GLOBAL_TEXT_SELECTOR), tipoCuentaRegex)) return true;
  }

  if (codigoRegex) {
    if (await checkLocator(seccionProductos, codigoRegex)) return true;

    if (options.useGlobalFallback) {
      if (await checkLocator(page.locator(GLOBAL_TEXT_SELECTOR), codigoRegex)) return true;
    }
  }

  if (options.modalMatcher) {
    const dialogs = page.locator('.p-dialog:visible, [role="dialog"]:visible');
    if (await options.modalMatcher(dialogs, tipoCuentaRegex, codigoRegex)) return true;
  }

  if (options.modalRegex) {
    const modalProducto = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: options.modalRegex })
        .first();
    const modalVisible = await modalProducto.isVisible().catch(() => false);
    if (modalVisible) {
        if (await checkLocator(modalProducto, tipoCuentaRegex)) return true;
        if (codigoRegex && await checkLocator(modalProducto, codigoRegex)) return true;
    }
  }

  return false;
}
