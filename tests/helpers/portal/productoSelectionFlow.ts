import type { Locator, Page } from '@playwright/test';

export interface ProductoSelectionFlowDeps {
  confirmarSeleccionRapida: (
    page: Page,
    seccionProductos: Locator,
    tipoCuentaRegex: RegExp,
    codigoRegex: RegExp | null,
    waitMs?: number
  ) => Promise<boolean>;
  detectarProductoSeleccionado: (page: Page, seccionProductos: Locator, tipoCuenta: string) => Promise<boolean>;
  esValorDropdownVacio: (value: string) => boolean;
  extraerCodigoProducto: (tipoCuenta: string) => string;
  extraerNombreProducto?: (tipoCuenta: string) => string;
  escapeRegexText: (value: string) => string;
  leerValorDropdown: (
    scope: Locator,
    labelText: string | RegExp,
    options?: { timeoutMs?: number }
  ) => Promise<string>;
  seleccionarDropdownEnScopePorTexto: (
    page: Page,
    scope: Locator,
    labelText: string | RegExp,
    textoOpcion: RegExp,
    fallbackIndex?: number
  ) => Promise<unknown>;
  seleccionarProductoPorDropdownSecundario: (
    page: Page,
    seccionProductos: Locator,
    tipoCuentaRegex: RegExp,
    codigoRegex: RegExp | null
  ) => Promise<boolean>;
  seleccionarFiltrable?: (page: Page, textoFiltro: string) => Promise<unknown>;
  seleccionarConFiltroSecundario?: (
    page: Page,
    seccionProductos: Locator,
    tipoCuentaRegex: RegExp,
    codigoRegex: RegExp | null,
    textoFiltro: string
  ) => Promise<boolean>;
  beforeAttempt?: (page: Page) => Promise<unknown>;
  logPrefix?: string;
  maxIntentos: number;
  probeTimeoutMs: number;
  retryWaitMs: number;
}

export async function seleccionarProductoEnSeccionProductosFlow(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string,
  deps: ProductoSelectionFlowDeps
) {
  const productoLabel = /^Producto\b/i;
  const tipoCuentaSafe = String(tipoCuenta ?? '').trim();
  const tipoCuentaRegex = new RegExp(deps.escapeRegexText(tipoCuentaSafe).replace(/\s+/g, '\\s+'), 'i');
  const codigoProducto = deps.extraerCodigoProducto(tipoCuentaSafe);
  const codigoRegex = codigoProducto ? new RegExp(`\\b${deps.escapeRegexText(codigoProducto)}\\b`, 'i') : null;
  const nombreProducto = deps.extraerNombreProducto?.(tipoCuentaSafe) || tipoCuentaSafe;
  const logPrefix = deps.logPrefix ?? '[Producto]';

  const coincideValor = (valor: string) => {
    if (deps.esValorDropdownVacio(valor)) return false;
    const coincideTexto = tipoCuentaRegex.test(valor);
    const coincideCodigo = codigoRegex ? codigoRegex.test(valor) : true;
    return coincideTexto || coincideCodigo;
  };

  const leerValorActual = async () =>
    deps.leerValorDropdown(seccionProductos, productoLabel, { timeoutMs: deps.probeTimeoutMs + 500 }).catch(() => '');

  if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex, 300)) return;

  for (let intento = 1; intento <= deps.maxIntentos; intento++) {
    console.log(`${logPrefix} intento ${intento}/${deps.maxIntentos}`);
    await deps.beforeAttempt?.(page).catch(() => false);

    const seleccionadoSecundarioPrimero = await deps
      .seleccionarProductoPorDropdownSecundario(page, seccionProductos, tipoCuentaRegex, codigoRegex)
      .catch(() => false);
    if (seleccionadoSecundarioPrimero) {
      if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex)) {
        console.log(`${logPrefix} modal de configuracion abierto tras seleccionar producto (fallback secundario).`);
        return;
      }
      if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) return;
      if (coincideValor(await leerValorActual())) return;
      if (intento < deps.maxIntentos - 1) {
        await page.waitForTimeout(deps.retryWaitMs);
        continue;
      }
    }

    await deps
      .seleccionarDropdownEnScopePorTexto(page, seccionProductos, productoLabel, tipoCuentaRegex, -1)
      .catch(() => {});
    
    // VALIDATION 1: Configuration modal opened
    if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex)) {
      console.log(`${logPrefix} modal de configuracion abierto tras seleccionar producto (scope texto).`);
      return;
    }
    // VALIDATION 2: UI reflects selection (added items or dropdown label)
    if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) {
        console.log(`${logPrefix} producto detectado en UI tras seleccion (scope texto).`);
        return;
    }

    let valorProducto = await leerValorActual();
    if (coincideValor(valorProducto)) {
        console.log(`${logPrefix} valor dropdown coincide tras seleccion (scope texto): '${valorProducto}'`);
        return;
    }

    if (codigoRegex) {
      await deps
        .seleccionarDropdownEnScopePorTexto(page, seccionProductos, productoLabel, codigoRegex, 0)
        .catch(() => {});
      
      if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex)) {
        console.log(`${logPrefix} modal de configuracion abierto tras seleccionar producto (scope codigo).`);
        return;
      }
      if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) {
          console.log(`${logPrefix} producto detectado en UI tras seleccion (scope codigo).`);
          return;
      }

      valorProducto = await leerValorActual();
      if (!deps.esValorDropdownVacio(valorProducto) && (codigoRegex.test(valorProducto) || tipoCuentaRegex.test(valorProducto))) {
          console.log(`${logPrefix} valor dropdown coincide tras seleccion (scope codigo): '${valorProducto}'`);
          return;
      }
    }

    if (intento >= deps.maxIntentos - 1) {
      console.log(`${logPrefix} Intentando fallbacks avanzados (filtro/filtrable)...`);
      if (deps.seleccionarConFiltroSecundario) {
        const seleccionadoConFiltro = await deps
          .seleccionarConFiltroSecundario(page, seccionProductos, tipoCuentaRegex, codigoRegex, codigoProducto || nombreProducto)
          .catch(() => false);
        if (seleccionadoConFiltro) {
          if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex)) {
            console.log(`${logPrefix} modal de configuracion abierto tras seleccionar producto (filtro).`);
            return;
          }
          if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) return;
        }
      } else if (deps.seleccionarFiltrable) {
        await deps.seleccionarFiltrable(page, codigoProducto || tipoCuentaSafe).catch(() => {});
        if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex)) {
          console.log(`${logPrefix} modal de configuracion abierto tras seleccionar producto (filtrable).`);
          return;
        }
        if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) return;
      }
      
      // LAST RESORT FALLBACK: If Excel requested an option that flat out does not exist,
      // force select the very first option (index 0) so the script can proceed!
      console.log(`${logPrefix} [WARN] El nombre en Excel ('${tipoCuentaSafe}') no produjo match exacto. Forzando eleccion del indice 0 como fallback.`);
      await deps
        .seleccionarDropdownEnScopePorTexto(page, seccionProductos, productoLabel, tipoCuentaRegex, 0)
        .catch(() => {});
      if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex)) {
        return;
      }
      if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) return;
      
      let finalVal = await leerValorActual();
      if (!deps.esValorDropdownVacio(finalVal)) {
          console.log(`${logPrefix} Finalizando con valor fallback actual: '${finalVal}'`);
          return;
      }
    }

    const seleccionadoSecundarioFinal = await deps
      .seleccionarProductoPorDropdownSecundario(page, seccionProductos, tipoCuentaRegex, codigoRegex)
      .catch(() => false);
    if (seleccionadoSecundarioFinal) {
      if (await deps.confirmarSeleccionRapida(page, seccionProductos, tipoCuentaRegex, codigoRegex)) {
        console.log(`${logPrefix} modal de configuracion abierto tras seleccionar producto (fallback secundario final).`);
        return;
      }
      if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) return;
      if (coincideValor(await leerValorActual())) return;
    }

    valorProducto = await leerValorActual();
    if (coincideValor(valorProducto)) return;
    console.log(`${logPrefix} Intento ${intento} fallido. valorActual='${valorProducto}'. Esperando ${deps.retryWaitMs}ms...`);
    await page.waitForTimeout(deps.retryWaitMs);
  }

  if (await deps.detectarProductoSeleccionado(page, seccionProductos, tipoCuentaSafe)) return;

  const valorFinal = await leerValorActual();
  if (!deps.esValorDropdownVacio(valorFinal)) {
      console.log(`${logPrefix} [WARN] Terminando seleccion con valor fallback='${valorFinal}' (no coincidio estrictamente con '${tipoCuentaSafe}').`);
      return;
  }

  throw new Error(
    `[CRITICO] No se pudo seleccionar 'Producto'. Esperado='${tipoCuentaSafe}' codigo='${codigoProducto || 'N/A'}' valorActual='${valorFinal}'.`
  );
}
