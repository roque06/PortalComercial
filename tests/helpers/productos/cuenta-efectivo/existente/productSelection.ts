import type { Locator, Page } from '@playwright/test';
import { seleccionarProductoEnSeccionProductosFlow } from '../../../portal/productoSelectionFlow';

export interface ProductoCuentaEfectivoExistenteDeps {
  confirmarSeleccionProductoRapida: (
    page: Page,
    seccionProductos: Locator,
    tipoCuentaRegex: RegExp,
    codigoRegex: RegExp | null,
    waitMs?: number
  ) => Promise<boolean>;
  detectarProductoSeleccionadoEnUI: (page: Page, seccionProductos: Locator, tipoCuenta: string) => Promise<boolean>;
  esValorDropdownVacio: (value: string) => boolean;
  escapeRegexText: (value: string) => string;
  extraerCodigoProducto: (tipoCuenta: string) => string;
  leerValorDropdownEnScope: (
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
  seleccionarDropdownFiltrableConReintentar: (page: Page, textoFiltro: string) => Promise<unknown>;
  beforeAttempt: (page: Page) => Promise<unknown>;
  maxIntentos: number;
  probeTimeoutMs: number;
  retryWaitMs: number;
}

export async function seleccionarProductoCuentaEfectivoExistente(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string,
  deps: ProductoCuentaEfectivoExistenteDeps
) {
  return seleccionarProductoEnSeccionProductosFlow(page, seccionProductos, tipoCuenta, {
    confirmarSeleccionRapida: deps.confirmarSeleccionProductoRapida,
    detectarProductoSeleccionado: deps.detectarProductoSeleccionadoEnUI,
    esValorDropdownVacio: deps.esValorDropdownVacio,
    escapeRegexText: deps.escapeRegexText,
    extraerCodigoProducto: deps.extraerCodigoProducto,
    leerValorDropdown: deps.leerValorDropdownEnScope,
    seleccionarDropdownEnScopePorTexto: deps.seleccionarDropdownEnScopePorTexto,
    seleccionarProductoPorDropdownSecundario: deps.seleccionarProductoPorDropdownSecundario,
    seleccionarFiltrable: deps.seleccionarDropdownFiltrableConReintentar,
    beforeAttempt: deps.beforeAttempt,
    logPrefix: '[Producto]',
    maxIntentos: deps.maxIntentos,
    probeTimeoutMs: deps.probeTimeoutMs,
    retryWaitMs: deps.retryWaitMs,
  });
}
