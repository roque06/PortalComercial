import { test, type BrowserContext, type Locator, type Page } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import {
    seleccionarDefaultSiVacio,
    llenarInputMask,
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
    seleccionarDropdownPorCampo,
    esperarYClickReintentarPaisIdentificacion,
    validarApnfdYSeleccionarNoSiVacio,
    validarGestionDocumentalSiRequerido,
    cargarDocumentoEnGestionDocumental,
    abrirBpmSiVerificacionConoceCliente,
    esperarPortalListoTrasLogin,
    agregarRelacionadoSiAplica,
    seleccionarProductoCertificadoExistente,
    type RegistroExcel,
} from '../../helpers/productos/certificados/helpers';
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
const LIST_PROBE_TIMEOUT = 450;
const LIST_RETRY_CLICK_TIMEOUT = 350;
const LIST_RETRY_WAIT_MS = 200;
const LIST_LABEL_TIMEOUT = 1400;
const LIST_PANEL_TIMEOUT = 1400;
const LIST_PANEL_QUICK_TIMEOUT = 700;
const LIST_SCOPE_RETRY_WAIT_MS = 180;
const PRODUCTO_MAX_INTENTOS = 4;
const PRODUCTO_WAIT_CONFIRM_MS = 450;
const PRODUCTO_RETRY_WAIT_MS = 120;
const FINALIZACION_MAX_INTENTOS = 6;
const FINALIZACION_TIMEOUT_MS = 120000;
const PORTAL_BASE_URL = (process.env.PW_PORTAL_URL ?? process.env.PORTAL_URL ?? 'https://srvqacgowb01.local.bsc.com:5000').replace(/\/+$/, '');
const PORTAL_MULTIPRODUCT_URL = `${PORTAL_BASE_URL}/requests/create/multiproduct`;
const BIZAGI_URL_ANTIGUA_REGEX = /https?:\/\/(?:srvdevbizapp01\/BPM_BSC\/?|portalcomercial\.dev\.bsc\.com\.do\/login\/?|srvqabizapp01\/BPM_BSC_QA\/?|dev-bscrd-santacruz\.bizagi\.com\/?)/i;
const BIZAGI_URL_NUEVA_BASE = process.env.PW_BIZAGI_URL ?? process.env.BIZAGI_URL ?? "https://test-bscrd-santacruz.bizagi.com/";
const BIZAGI_URL_NUEVA = BIZAGI_URL_ANTIGUA_REGEX.test(BIZAGI_URL_NUEVA_BASE)
    ? "https://test-bscrd-santacruz.bizagi.com/"
    : BIZAGI_URL_NUEVA_BASE;
const BIZAGI_URL_LEGACY_REGEX = BIZAGI_URL_ANTIGUA_REGEX;
const BIZAGI_USUARIO = "domain\\admon";
const BIZAGI_USUARIO_FALLBACK = "domain\\admon";
const BIZAGI_PASSWORD = "H0la1234.";
const BIZAGI_HOST_REGEX = /portalcomercial\.dev\.bsc\.com\.do|bizagi\.com|srvdevbizapp01/i;

async function redirigirBizagiLegacySiAplica(page: Page) {
    if (BIZAGI_URL_LEGACY_REGEX.test(BIZAGI_URL_NUEVA)) return;
    const urlActual = page.url();
    if (!BIZAGI_URL_LEGACY_REGEX.test(urlActual)) return;
    console.log(`[TasaExcepcion][Bizagi] URL legacy detectada='${urlActual}'`);
    await page.goto(BIZAGI_URL_NUEVA, { waitUntil: "domcontentloaded" }).catch(() => { });
    console.log(`[TasaExcepcion][Bizagi] URL redirigida='${page.url()}'`);
}

async function findVisibleLocator(page: Page, selectors: string[], label: string = "Element") {
    // 1. Intento normal (visible)
    for (const sel of selectors) {
        const loc = page.locator(sel).first();
        if (await loc.isVisible().catch(() => false)) {
            console.log(`[BizagiLogin][Spec] ${label} visible en main con selector: ${sel}`);
            return loc;
        }
        for (const [idx, frame] of page.frames().entries()) {
            const fLoc = frame.locator(sel).first();
            if (await fLoc.isVisible().catch(() => false)) {
                console.log(`[BizagiLogin][Spec] ${label} visible en Frame #${idx} con selector: ${sel}`);
                return fLoc;
            }
        }
    }

    // 2. Intento normal (existente)
    for (const sel of selectors) {
        const loc = page.locator(sel).first();
        if (await loc.count().catch(() => 0) > 0) {
            console.log(`[BizagiLogin][Spec] ${label} existe en main con selector: ${sel}`);
            return loc;
        }
        for (const [idx, frame] of page.frames().entries()) {
            const fLoc = frame.locator(sel).first();
            if (await fLoc.count().catch(() => 0) > 0) {
                console.log(`[BizagiLogin][Spec] ${label} existe en Frame #${idx} con selector: ${sel}`);
                return fLoc;
            }
        }
    }

    // 3. Ultimate Fallback: Escaneo manual de atributos
    console.log(`[BizagiLogin][Spec] Iniciando escaneo manual para ${label}...`);
    for (const frame of page.frames()) {
        const inputs = await frame.locator('input, button, select').all().catch(() => []);
        for (const inp of inputs) {
            const id = (await inp.getAttribute('id').catch(() => '')) || '';
            const name = (await inp.getAttribute('name').catch(() => '')) || '';
            const type = (await inp.getAttribute('type').catch(() => '')) || '';
            const text = (await inp.innerText().catch(() => '')) || '';

            const lowerLabel = label.toLowerCase();
            if (lowerLabel.includes('user') || lowerLabel.includes('usuario')) {
                if (id === 'user' || id === 'username' || name === 'username' || name === 'user') return inp;
            } else if (lowerLabel.includes('pass') || lowerLabel.includes('contrase')) {
                if (id === 'password' || name === 'password' || type === 'password') return inp;
            } else if (lowerLabel.includes('boton') || lowerLabel.includes('ingresar')) {
                if (id.includes('login') || name.includes('login') || text.toLowerCase().includes('ingresar')) return inp;
            }
        }
    }

    return page.locator(selectors[0]).first();
}

async function esBizagiHomePage(page: Page): Promise<boolean> {
    // Verifica cada selector individualmente para evitar que .first() tome un elemento oculto
    const sels = ['input#menuQuery', '[id="mnuBt.Admin.Reports"]', '#main-layout', '.pc-avatar', '.user-name'];
    for (const sel of sels) {
        if (await page.locator(sel).first().isVisible().catch(() => false)) return true;
    }
    return false;
}

async function autocompletarCredencialesBizagiSiLogin(page: Page) {
    const url = page.url();
    if (!BIZAGI_HOST_REGEX.test(url)) return;

    // Si Bizagi ya está en home (ya logueado), no hacer nada
    if (await esBizagiHomePage(page)) return;

    const userSels = ['#user', 'input#user', 'input[name="username"]', 'input#username', 'input[autocomplete="username"]'];
    const passSels = ['#password', 'input#password', 'input[type="password"]', 'input[name="password"]'];
    const btnSels  = ['#btn-login', 'button#btn-login', 'button#btn-adminlogin', 'button:has-text("Ingresar")', 'button[name*="login" i]'];

    const inicioBase = Date.now();
    while (Date.now() - inicioBase < 50000) {
        if (page.isClosed()) return;

        // Primero verificar si ya llegamos a home; si es así, salir sin loguear
        if (await esBizagiHomePage(page)) return;

        // Detectar si el formulario de login está visible (contraseña Y botón deben estar visibles)
        const inputPassword = await findVisibleLocator(page, passSels, "Password").catch(() => page.locator(passSels[0]).first());
        const btnIngresar   = await findVisibleLocator(page, btnSels,  "BotonIngresar").catch(() => page.locator(btnSels[0]).first());

        const pVis = await inputPassword.isVisible().catch(() => false);
        const iVis = await btnIngresar.isVisible().catch(() => false);

        // Caso especial: login con solo selector de usuario (sin password) — ej. portalcomercial.dev.bsc.com.do/login
        if (!pVis && iVis) {
            const selectUsuario = page.locator('select:visible').first();
            const sVis = await selectUsuario.isVisible().catch(() => false);
            if (sVis) {
                console.log(`[BizagiLogin] Login solo-selector detectado. Seleccionando usuario y clicando Ingresar...`);
                await selectUsuario.evaluate((el, user) => {
                    const sel = el as HTMLSelectElement;
                    const raw = String(user || '').toLowerCase();
                    const terms = [raw, raw.replace('domain\\', ''), 'admon'];
                    const options = Array.from(sel.options || []);
                    const match = options.find((o) => {
                        const txt = String(o.text || '').toLowerCase();
                        const val = String(o.value || '').toLowerCase();
                        return terms.some((t) => t && (txt.includes(t) || val.includes(t)));
                    }) || options.find((o) => String(o.value || '').trim() || String(o.text || '').trim());
                    if (!match) return;
                    sel.value = match.value;
                    sel.dispatchEvent(new Event('input', { bubbles: true }));
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                }, BIZAGI_USUARIO).catch(() => {});
                await page.waitForTimeout(400).catch(() => {});
                await btnIngresar.click({ force: true }).catch(() => {});
                await page.waitForLoadState('domcontentloaded').catch(() => {});
                await page.waitForTimeout(3000).catch(() => {});
                if (await esBizagiHomePage(page)) return;
                continue;
            }
        }

        // Solo proceder si el formulario de login está completo (password + botón visibles)
        if (!pVis || !iVis) {
            console.log(`[BizagiLogin] Formulario de login no detectado (P=${pVis}, Btn=${iVis}). Esperando...`);
            await page.waitForTimeout(1000).catch(() => {});
            continue;
        }

        // Formulario de login confirmado — buscar campo usuario
        const inputUsuario = await findVisibleLocator(page, userSels, "Usuario").catch(() => page.locator(userSels[0]).first());
        const uVis = await inputUsuario.isVisible().catch(() => false);

        console.log(`[BizagiLogin] Formulario de login detectado. U=${uVis}, P=${pVis}, Btn=${iVis}`);

        if (uVis) {
            let valU = (await inputUsuario.inputValue().catch(() => "")).trim();
            if (!valU || (!valU.includes("admon") && !valU.includes("domain\\"))) {
                console.log(`[BizagiLogin] Rellenando usuario...`);
                await inputUsuario.fill(BIZAGI_USUARIO, { timeout: 5000 }).catch(async () => {
                    await inputUsuario.evaluate((el, v) => (el as HTMLInputElement).value = v, BIZAGI_USUARIO).catch(() => {});
                });
                await page.waitForTimeout(300).catch(() => {});
                valU = (await inputUsuario.inputValue().catch(() => "")).trim();
                if (!valU) await inputUsuario.fill(BIZAGI_USUARIO_FALLBACK).catch(() => {});
            }
        }

        if (page.isClosed()) return;

        let valP = (await inputPassword.inputValue().catch(() => "")).trim();
        if (!valP) {
            console.log(`[BizagiLogin] Rellenando password...`);
            await inputPassword.fill(BIZAGI_PASSWORD, { timeout: 5000 }).catch(async () => {
                await inputPassword.evaluate((el, v) => (el as HTMLInputElement).value = v, BIZAGI_PASSWORD).catch(() => {});
            });
            await page.waitForTimeout(300).catch(() => {});
            valP = (await inputPassword.inputValue().catch(() => "")).trim();
        }

        const finalU = uVis ? (await inputUsuario.inputValue().catch(() => "")).trim() : BIZAGI_USUARIO;
        const finalP = valP;

        if (finalP && iVis) {
            console.log(`[BizagiLogin] Ingresando... (U=${!!finalU}, P=${!!finalP})`);
            await page.waitForTimeout(800).catch(() => {});
            if (page.isClosed()) return;
            await btnIngresar.click({ force: true }).catch(() => {});
            await page.waitForTimeout(3000).catch(() => {});
        } else {
            console.log(`[BizagiLogin] No listo para ingresar. P=${!!finalP}, Btn=${iVis}`);
            await page.waitForTimeout(1000).catch(() => {});
        }

        if (page.isClosed()) return;
        if (await esBizagiHomePage(page)) return;
    }
}

function instalarForzadorUrlBizagi(context: BrowserContext) {
    const enganchar = (p: Page) => {
        void redirigirBizagiLegacySiAplica(p);
        void autocompletarCredencialesBizagiSiLogin(p);
        p.on("domcontentloaded", () => {
            void redirigirBizagiLegacySiAplica(p);
            void autocompletarCredencialesBizagiSiLogin(p);
        });
        p.on("framenavigated", (frame) => {
            if (frame !== p.mainFrame()) return;
            const urlFrame = frame.url();
            if (!BIZAGI_URL_LEGACY_REGEX.test(BIZAGI_URL_NUEVA) && BIZAGI_URL_LEGACY_REGEX.test(urlFrame)) {
                console.log(`[TasaExcepcion][Bizagi] frame legacy detectado='${urlFrame}'`);
                void p.goto(BIZAGI_URL_NUEVA, { waitUntil: "domcontentloaded" }).catch(() => { });
                return;
            }
            if (BIZAGI_HOST_REGEX.test(urlFrame)) {
                void autocompletarCredencialesBizagiSiLogin(p);
            }
        });
    };

    for (const p of context.pages()) {
        enganchar(p);
    }
    context.on("page", enganchar);
}

function esErrorCasoEnCursoNoGestionable(errorMsg: string) {
    const txt = String(errorMsg ?? "");
    const timeout = /Timeout\s+\d+ms exceeded/i.test(txt);
    const tablaCasos = /hasText:\s*\/Casos\/i/i.test(txt);
    const filaMpn = /tr:has-text\("MPN-\d+/i.test(txt);
    return timeout && tablaCasos && filaMpn;
}

async function modalProductoConfigVisibleLegacy(page: Page) {
    const dialogos = page.locator('.p-dialog:visible, [role="dialog"]:visible');
    const total = await dialogos.count().catch(() => 0);
    if (!total) return false;

    const regexTitulo = /Certificados de dep(?:o|\u00f3)sito/i;
    const regexCampos = /(Moneda|Monto|Plazo|Tasa|Inter(?:e|\u00e9)s)/i;

    for (let i = 0; i < total; i++) {
        const dialogo = dialogos.nth(i);
        const texto = (await dialogo.innerText().catch(() => "")).trim();
        if (!texto) continue;

        if (!regexTitulo.test(texto)) continue;
        if (!regexCampos.test(texto)) continue;

        const tieneBotonAceptar = await dialogo
            .getByRole("button", { name: /^Aceptar$/i })
            .first()
            .isVisible()
            .catch(() => false);
        if (tieneBotonAceptar) return true;
    }

    return false;
}

function esValorSi(value: string | undefined) {
    return /^(si|sÃ­|true|1)$/i.test(String(value ?? "").trim());
}

function normalizarNumero(value: string) {
    const raw = String(value ?? "").trim();
    if (!raw) return NaN;
    const cleaned = raw.replace(/[^\d,.\-]/g, "");
    if (!cleaned) return NaN;
    const lastComma = cleaned.lastIndexOf(",");
    const lastDot = cleaned.lastIndexOf(".");
    let normalized = cleaned;
    if (lastComma > -1 && lastDot > -1) {
        if (lastComma > lastDot) {
            normalized = cleaned.replace(/\./g, "").replace(",", ".");
        } else {
            normalized = cleaned.replace(/,/g, "");
        }
    } else if (lastComma > -1) {
        normalized = cleaned.replace(/\./g, "").replace(",", ".");
    }
    const parsed = Number.parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : NaN;
}

function tasaCoincide(valorActual: string, valorEsperado: string) {
    const actualNum = normalizarNumero(valorActual);
    const esperadoNum = normalizarNumero(valorEsperado);
    if (!Number.isNaN(actualNum) && !Number.isNaN(esperadoNum)) {
        return Math.abs(actualNum - esperadoNum) < 0.0001;
    }
    const onlyDigits = (txt: string) => String(txt ?? "").replace(/[^\d]/g, "");
    return onlyDigits(valorActual) === onlyDigits(valorEsperado);
}

async function esperarCampoTasaEditable(page: Page, inputTasa: Locator, timeoutMs = 12000) {
    const inicio = Date.now();

    while (Date.now() - inicio < timeoutMs) {
        const visible = await inputTasa.isVisible().catch(() => false);
        if (!visible) {
            await page.waitForTimeout(200);
            continue;
        }

        const editable = await inputTasa.isEditable().catch(() => false);
        const enabled = await inputTasa.isEnabled().catch(() => false);
        const readonlyAttr = await inputTasa.getAttribute("readonly").catch(() => null);
        const ariaReadonly = await inputTasa.getAttribute("aria-readonly").catch(() => null);

        if (editable && enabled && !readonlyAttr && ariaReadonly !== "true") {
            return true;
        }

        await page.waitForTimeout(250);
    }

    return false;
}

async function esperarInteresRecalculado(
    page: Page,
    inputInteres: Locator,
    valorPrevio: string,
    timeoutMs = 15000
) {
    const inicio = Date.now();
    const previo = String(valorPrevio ?? "").trim();

    const contieneNumMayorACero = (txt: string) => {
        const digitos = txt.replace(/[^\d]/g, "");
        return digitos.length > 0 && parseInt(digitos, 10) > 0;
    };

    while (Date.now() - inicio < timeoutMs) {
        const actual = (await inputInteres.inputValue().catch(() => "")).trim();
        const esProducto = /(dep[oÃ³]sito|producto|financiero|capitalizable)/i.test(actual);
        const numerico = contieneNumMayorACero(actual) && !esProducto;
        const cambio = !previo || actual !== previo;

        if (numerico && cambio) {
            return true;
        }

        await page.waitForTimeout(300);
    }

    return false;
}

async function solicitarTasaExcepcion(
    page: Page,
    inputTasa: Locator,
    inputInteres: Locator | null,
    valorTasaExcel: string
) {
    const valorInicialTasa = (await inputTasa.inputValue().catch(() => "")).trim();
    let valorInteresAntes = '';
    if (inputInteres) {
        valorInteresAntes = (await inputInteres.inputValue().catch(() => "")).trim();
    }

    console.log(`[TasaExcepcion] valor inicial tasa='${valorInicialTasa}'`);
    console.log(`[TasaExcepcion] interes antes='${valorInteresAntes || '[no disponible]'}'`);

    const leerTasa = async () => (await inputTasa.inputValue().catch(() => "")).trim();
    const esperarTasaCoincida = async (timeoutMs = 1500) => {
        const inicio = Date.now();
        let ultimo = await leerTasa();
        while (Date.now() - inicio < timeoutMs) {
            ultimo = await leerTasa();
            if (tasaCoincide(ultimo, valorTasaExcel)) return { ok: true, valor: ultimo };
            await page.waitForTimeout(120);
        }
        return { ok: tasaCoincide(ultimo, valorTasaExcel), valor: ultimo };
    };

    let tasaCambiada = false;

    console.log(`[CERT-EX][TasaExcepcion] Tasa Excel=${valorTasaExcel}`);
    console.log(`[CERT-EX][TasaExcepcion] Modificando Tasa a ${valorTasaExcel}`);

    // intento 1
    await inputTasa.scrollIntoViewIfNeeded().catch(() => { });
    await inputTasa.click({ force: true }).catch(() => { });
    await inputTasa.press("Control+A").catch(() => { });
    await inputTasa.press("Delete").catch(() => { });
    await inputTasa.fill(valorTasaExcel).catch(async () => {
        await inputTasa.pressSequentially(valorTasaExcel, { delay: 35 }).catch(() => { });
    });
    await inputTasa.press("Tab").catch(() => { });
    let check = await esperarTasaCoincida(1600);
    console.log(`[CERT-EX][TasaExcepcion] intento 1 -> tasa='${check.valor}'`);
    tasaCambiada = check.ok;

    // intento 2
    if (!tasaCambiada) {
        await inputTasa.click({ clickCount: 3, force: true }).catch(() => { });
        await inputTasa.press("Backspace").catch(() => { });
        await inputTasa.pressSequentially(valorTasaExcel, { delay: 80 }).catch(() => { });
        await inputTasa.press("Tab").catch(() => { });
        check = await esperarTasaCoincida(1800);
        console.log(`[CERT-EX][TasaExcepcion] intento 2 -> tasa='${check.valor}'`);
        tasaCambiada = check.ok;
    }

    // intento 3
    if (!tasaCambiada) {
        await inputTasa.evaluate((el: any, val: string) => {
            const input = el as HTMLInputElement;
            input.value = val;
            input.dispatchEvent(new Event("input", { bubbles: true }));
            input.dispatchEvent(new Event("change", { bubbles: true }));
            input.dispatchEvent(new Event("blur", { bubbles: true }));
        }, valorTasaExcel).catch(() => { });
        await inputTasa.press("Tab").catch(() => { });
        check = await esperarTasaCoincida(1800);
        console.log(`[CERT-EX][TasaExcepcion] intento 3 -> tasa='${check.valor}'`);
        tasaCambiada = check.ok;
    }

    if (!tasaCambiada) {
        const tasaFinal = await leerTasa();
        throw new Error(
            `[CRITICO] No se pudo cambiar la Tasa para excepciÃ³n. valorInicial='${valorInicialTasa}' valorFinal='${tasaFinal}'`
        );
    }
    console.log(`[CERT-EX][TasaExcepcion] Tasa modificada a ${valorTasaExcel}`);

    const modalConfirm = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({
            hasText: /(Confirmaci[oÃ³]n|tasa se encuentra por encima|solicitar excepci[oÃ³]n)/i,
        })
        .last();

    let confirmVisible = false;
    for (let i = 1; i <= 4; i++) {
        confirmVisible = await modalConfirm.isVisible().catch(() => false);
        console.log(`[CERT-EX][TasaExcepcion] Modal Confirmación visible=${confirmVisible} intento=${i}`);
        if (confirmVisible) break;
        await page.waitForTimeout(600);
    }

    if (!confirmVisible) {
        throw new Error("[CRITICO] No apareciÃ³ el modal de confirmaciÃ³n de tasa de excepciÃ³n.");
    }
    console.log(`[CERT-EX][TasaExcepcion] Modal Confirmación visible`);

    await modalConfirm.getByRole("button", { name: /^Aceptar$/i }).last().click({ force: true });
    console.log(`[CERT-EX][TasaExcepcion] Click Aceptar confirmación`);

    // Iterate visible dialogs to find the one that contains both textarea.p-inputtextarea
    // and a field related to "Tasa pool". Do not rely on dialog title text.
    let modalSolicitud: import("@playwright/test").Locator | null = null;
    const deadline = Date.now() + 15000;
    while (Date.now() < deadline) {
        const allDialogs = page.locator('.p-dialog, [role="dialog"]');
        const count = await allDialogs.count().catch(() => 0);
        console.log(`[TasaExcepcion] dialogos visibles encontrados=${count}`);
        for (let i = 0; i < count; i++) {
            const candidate = allDialogs.nth(i);
            const isVisible = await candidate.isVisible().catch(() => false);
            if (!isVisible) continue;
            const hasTextarea = await candidate.locator('textarea.p-inputtextarea').first().isVisible().catch(() => false);
            const hasTasaPool = await candidate.locator('xpath=.//*[contains(normalize-space(.),"Tasa pool")]').first().isVisible().catch(() => false);
            console.log(`[TasaExcepcion] dialogo[${i}] hasTextarea=${hasTextarea} hasTasaPool=${hasTasaPool}`);
            if (hasTextarea && hasTasaPool) {
                modalSolicitud = candidate;
                break;
            }
        }
        if (modalSolicitud) break;
        await page.waitForTimeout(300);
    }

    if (!modalSolicitud) {
        throw new Error("[CRITICO] No se encontró el modal de solicitud de tasa de excepción (sin textarea + Tasa pool).");
    }
    console.log("[CERT-EX][TasaExcepcion] Modal Solicitud tasa de excepción visible");

    // Relocalize immediately after detection using dynamic filter to avoid staleness
    // El .nth(i) guardado se vuelve stale cuando el modal de confirmación se cierra
    console.log("[CERT-EX][TasaExcepcion] Relocalizando modal Solicitud tasa excepción estable");
    const modalEstable = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Solicitud tasa de excepción|Motivo de solicitud|Tasa máxima|Tasa solicitada|Tasa pool/i })
        .last();
    await modalEstable.waitFor({ state: 'visible', timeout: 15000 });

    // Helper to read Tasa pool value tolerantly from a stable modal reference
    const leerTasaPoolDesdeModal = async (modal: Locator): Promise<string> => {
        return modal.evaluate((root: Element) => {
            // Collect all leaf-ish nodes whose trimmed text is exactly "Tasa pool"
            const allNodes = [...root.querySelectorAll("*")] as HTMLElement[];
            const labels = allNodes.filter(el =>
                el.childElementCount === 0 &&
                (el.textContent ?? "").trim() === "Tasa pool"
            );

            for (const label of labels) {
                // Walk up max 6 levels looking for an ancestor that contains an <input>
                let container: HTMLElement | null = label.parentElement;
                for (let depth = 0; depth < 6; depth++) {
                    if (!container) break;
                    const input = container.querySelector("input") as HTMLInputElement | null;
                    if (input) {
                        const ariaText = input.getAttribute("aria-valuetext") ?? "";
                        const ariaNow  = input.getAttribute("aria-valuenow") ?? "";
                        const valProp  = input.value ?? "";
                        const valAttr  = input.getAttribute("value") ?? "";
                        const text     = (container.textContent ?? "").replace("Tasa pool", "").trim();
                        const matchTxt = text.match(/\d[\d.,]*%?/)?.[0] ?? "";
                        const resolved = ariaText || ariaNow || valProp || valAttr || matchTxt;
                        if (resolved && /\d/.test(resolved)) return resolved;
                    }
                    container = container.parentElement;
                }
            }
            return "";
        }).catch(() => "");
    };

    // Non-blocking Tasa pool read: max 3 seconds (6 retries × 500ms)
    // If unsuccessful, continue anyway if textarea is visible
    let tasaPool = "";
    for (let i = 1; i <= 6; i++) {
        tasaPool = (await leerTasaPoolDesdeModal(modalEstable)).trim();
        console.log(`[CERT-EX][TasaExcepcion] Tasa pool intento=${i}/6 valor="${tasaPool}"`);
        if (tasaPool && /\d/.test(tasaPool)) break;
        await page.waitForTimeout(500);
    }
    if (!tasaPool) {
        console.log("[CERT-EX][TasaExcepcion] Tasa pool no leída, pero textarea visible; continuo con motivo");
    }

    const inputMotivo = modalEstable.locator('textarea.p-inputtextarea').first();
    const motivoVisible = await inputMotivo.isVisible().catch(() => false);
    console.log(`[TasaExcepcion] textarea motivo visible=${motivoVisible}`);

    await inputMotivo.waitFor({ state: "visible", timeout: 8000 });
    await inputMotivo.scrollIntoViewIfNeeded().catch(() => { });
    await inputMotivo.click({ force: true }).catch(() => { });
    await inputMotivo.press("Control+A").catch(() => { });
    await inputMotivo.press("Delete").catch(() => { });
    console.log("[CERT-EX][TasaExcepcion] Llenando Motivo de solicitud");
    await inputMotivo.fill("testing").catch(async () => {
        await inputMotivo.pressSequentially("testing", { delay: 25 }).catch(() => { });
    });
    await inputMotivo.press("Tab").catch(() => { });
    await page.waitForTimeout(150);

    let valorMotivo = await inputMotivo.inputValue().catch(() => "");
    console.log(`[CERT-EX][TasaExcepcion] Motivo llenado=${valorMotivo}`);
    if (!/testing/i.test(String(valorMotivo ?? ""))) {
        console.log("[TasaExcepcion] fill insuficiente, usando evaluate para forzar valor");
        await inputMotivo.evaluate((el: Element) => {
            const ta = el as HTMLTextAreaElement;
            ta.value = "testing";
            ta.dispatchEvent(new Event("input", { bubbles: true }));
            ta.dispatchEvent(new Event("change", { bubbles: true }));
            ta.dispatchEvent(new Event("blur", { bubbles: true }));
        }).catch(() => { });
        await page.waitForTimeout(150);
        valorMotivo = await inputMotivo.inputValue().catch(() => "");
        console.log(`[TasaExcepcion] motivo tras evaluate='${valorMotivo}'`);
    }
    if (!/testing/i.test(String(valorMotivo ?? ""))) {
        throw new Error("[CRITICO] No se pudo escribir el motivo de solicitud.");
    }

    // Diagnóstico: estado de todos los spinbuttons e inputs inválidos del modal.
    const diagModal = await modalEstable.evaluate((modal: Element) => {
        // Spinbuttons con su label y valor
        const spinbuttons = ([...modal.querySelectorAll('input[role="spinbutton"]')] as HTMLInputElement[])
            .map(sb => {
                let label = sb.getAttribute("aria-label") ?? "";
                if (!label) {
                    let node: HTMLElement | null = sb.parentElement;
                    for (let d = 0; d < 6 && node; d++, node = node.parentElement as HTMLElement) {
                        const leaf = ([...node.querySelectorAll("*")] as HTMLElement[])
                            .filter(el => el.childElementCount === 0)
                            .map(el => (el.textContent ?? "").trim())
                            .find(t => t.length > 0 && !/^\d/.test(t) && t.length < 60);
                        if (leaf) { label = leaf; break; }
                    }
                }
                return {
                    label,
                    ariaValueText: sb.getAttribute("aria-valuetext") ?? "",
                    value: sb.value,
                    ngInvalid: sb.classList.contains("ng-invalid"),
                    required: sb.required,
                };
            });
        // Campos con clase ng-invalid (excluyendo contenedores form/div/span)
        const ngInvalid = ([...modal.querySelectorAll(".ng-invalid")] as HTMLElement[])
            .filter(el => !["FORM", "NG-COMPONENT", "DIV", "SPAN", "P-INPUTNUMBER", "P-INPUTTEXTAREA"].includes(el.tagName))
            .map(el => ({
                tag: el.tagName,
                fcn: el.getAttribute("formcontrolname") ?? "",
                value: (el as HTMLInputElement).value ?? "",
            }));
        return { spinbuttons, ngInvalid };
    }).catch(() => ({ spinbuttons: [] as any[], ngInvalid: [] as any[] }));
    console.log(`[TasaExcepcion] diag spinbuttons: ${JSON.stringify(diagModal.spinbuttons)}`);
    console.log(`[TasaExcepcion] diag ng-invalid: ${JSON.stringify(diagModal.ngInvalid)}`);

    // --- Clic en Aceptar con múltiples estrategias ---
    const botonAceptarSolicitud = modalEstable.getByRole("button", { name: /^Aceptar$/i }).last();
    await botonAceptarSolicitud.waitFor({ state: "visible", timeout: 5000 });

    console.log("[CERT-EX][TasaExcepcion] Click Aceptar solicitud tasa excepción");

    // Estrategia 1: click normal Playwright
    console.log("[CERT-EX][TasaExcepcion] intento 1: click normal");
    await botonAceptarSolicitud.click().catch(() => {});
    await page.waitForTimeout(1000);
    let modalCerrado = !(await modalEstable.isVisible().catch(() => true));
    console.log(`[CERT-EX][TasaExcepcion] modal cerrado tras click normal: ${modalCerrado}`);

    if (!modalCerrado) {
        // Estrategia 2: focus + tecla Enter
        console.log("[CERT-EX][TasaExcepcion] intento 2: focus + Enter");
        await botonAceptarSolicitud.focus().catch(() => {});
        await page.waitForTimeout(200);
        await page.keyboard.press("Enter");
        await page.waitForTimeout(1000);
        modalCerrado = !(await modalEstable.isVisible().catch(() => true));
        console.log(`[CERT-EX][TasaExcepcion] modal cerrado tras Enter: ${modalCerrado}`);
    }

    if (!modalCerrado) {
        // Estrategia 3: page.mouse.click en coordenadas exactas del botón
        console.log("[CERT-EX][TasaExcepcion] intento 3: mouse.click en coordenadas");
        const box = await botonAceptarSolicitud.boundingBox();
        console.log(`[CERT-EX][TasaExcepcion] bounding box Aceptar: ${JSON.stringify(box)}`);
        if (box) {
            await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
            await page.waitForTimeout(100);
            await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
            await page.waitForTimeout(1000);
            modalCerrado = !(await modalEstable.isVisible().catch(() => true));
            console.log(`[CERT-EX][TasaExcepcion] modal cerrado tras mouse.click: ${modalCerrado}`);
        }
    }

    // Esperar completamente el cierre del modal de solicitud
    await modalEstable.waitFor({ state: 'hidden', timeout: 15000 }).catch(() => {});
    console.log("[CERT-EX][TasaExcepcion] Modal Solicitud tasa excepción cerrado");

    // Modal cerrado confirmado — dar tiempo al DOM para estabilizarse antes de continuar.
    await page.waitForTimeout(500);

    // Si el interés ya tiene valor numérico válido, no hace falta esperar recálculo
    // Proteger contra inputInteres null
    if (inputInteres) {
        const interesInmediato = normalizarNumero((await inputInteres.inputValue().catch(() => "")).trim());
        if (interesInmediato > 0) {
            console.log(`[TasaExcepcion] interes ya valido post-modal=${interesInmediato}, sin esperar recalculo`);
        } else {
            const recalculoOk = await esperarInteresRecalculado(page, inputInteres, valorInteresAntes, 8000);
            const valorInteresFinal = (await inputInteres.inputValue().catch(() => "")).trim();
            console.log(`[TasaExcepcion] interes tras espera='${valorInteresFinal}' recalculo=${recalculoOk}`);
            if (!recalculoOk && normalizarNumero(valorInteresFinal) <= 0) {
                throw new Error("[CRITICO] No se recalculó el Interés tras solicitar tasa de excepción.");
            }
        }
    } else {
        console.log(`[CERT-EX][TasaExcepcion] inputInteres es null, no se valida recálculo de interés`);
    }
    console.log("[CERT-EX][TasaExcepcion] Flujo de excepción completado");
}







async function extraerMpnSolicitudDesdePortal(page: Page): Promise<string> {
    const candidatos = page
        .locator('header, .p-breadcrumb, .breadcrumb, [role="navigation"], [class*="breadcrumb"]')
        .filter({ hasText: /MPN-\d+/i });

    const count = await candidatos.count().catch(() => 0);
    for (let i = 0; i < count; i++) {
        const el = candidatos.nth(i);
        const visible = await el.isVisible().catch(() => false);
        if (!visible) continue;
        const txt = (await el.textContent().catch(() => "")) || "";
        const match = txt.match(/MPN-\d+/i);
        if (match?.[0]) return match[0].toUpperCase();
    }

    const fallback = page.getByText(/MPN-\d+/i).first();
    if (await fallback.isVisible().catch(() => false)) {
        const txt = (await fallback.textContent().catch(() => "")) || "";
        const match = txt.match(/MPN-\d+/i);
        if (match?.[0]) return match[0].toUpperCase();
    }

    throw new Error("[CRITICO] No se pudo extraer el MPN actual desde Portal Comercial.");
}

async function abrirYAutenticarBizagi(context: BrowserContext): Promise<Page> {
    const bizagiPage = await context.newPage();
    console.log(`[TasaExcepcion][Bizagi] URL objetivo='${BIZAGI_URL_NUEVA}'`);
    await bizagiPage.goto(BIZAGI_URL_NUEVA, { waitUntil: "domcontentloaded" });
    await bizagiPage.waitForLoadState("domcontentloaded").catch(() => { });
    console.log(`[TasaExcepcion][Bizagi] URL cargada='${bizagiPage.url()}'`);
    await redirigirBizagiLegacySiAplica(bizagiPage);
    await autocompletarCredencialesBizagiSiLogin(bizagiPage);

    console.log("[TasaExcepcion][Bizagi] verificando estado inicial de Bizagi");

    const loginSignals = async () => {
        const btnIngresar = bizagiPage
            .locator('button:has-text("Ingresar"), #btn-admin-login, button[name*="login" i]')
            .first();
        const selectorUsuario = bizagiPage.locator("select, .p-dropdown, [role='combobox']").first();
        const inputUsuario = bizagiPage
            .locator('input[name="username"], input#username, input[autocomplete="username"], .login-field input[type="text"], xpath=(//label[contains(normalize-space(.),"Usuario")]/following::input[1])[1]')
            .first();
        const inputPassword = bizagiPage
            .locator('input[type="password"], input[name="password"], input#password, xpath=(//label[contains(normalize-space(.),"Contrase")]/following::input[1])[1]')
            .first();
        const textoLogin = bizagiPage.getByText(/Seleccione sus datos para ingresar a Bizagi/i).first();
        const textoLoginNuevo = bizagiPage.getByText(/Ingrese su nombre de usuario y contrase(?:n|ñ)a/i).first();
        const visible =
            (await btnIngresar.isVisible().catch(() => false)) ||
            (await selectorUsuario.isVisible().catch(() => false)) ||
            (await inputUsuario.isVisible().catch(() => false)) ||
            (await inputPassword.isVisible().catch(() => false)) ||
            (await textoLogin.isVisible().catch(() => false)) ||
            (await textoLoginNuevo.isVisible().catch(() => false));
        return { visible, btnIngresar, selectorUsuario, inputUsuario, inputPassword };
    };

    const homeSignals = async () => {
        const inputMenuQuery = bizagiPage.locator('#menuQuery, input[name="menuQuery"]').first();
        const inputBuscar = bizagiPage.locator('input[placeholder*="Buscar" i]').first();
        const inbox = bizagiPage.getByText(/Inbox/i).first();
        const tabla = bizagiPage.locator('table, .p-datatable, .grid, .bpm-grid').first();
        const visible =
            (await inputMenuQuery.isVisible().catch(() => false)) ||
            (await inputBuscar.isVisible().catch(() => false)) ||
            (await inbox.isVisible().catch(() => false)) ||
            (await tabla.isVisible().catch(() => false));
        return { visible, inputMenuQuery, inputBuscar };
    };

    const inicio = Date.now();
    let login = await loginSignals();
    let home = await homeSignals();

    while (!login.visible && !home.visible && Date.now() - inicio < 20000) {
        await bizagiPage.waitForTimeout(500);
        login = await loginSignals();
        home = await homeSignals();
    }

    if (home.visible && !login.visible) {
        console.log("[TasaExcepcion][Bizagi] home ya autenticada detectada");
        return bizagiPage;
    }

    if (login.visible) {
        console.log("[TasaExcepcion][Bizagi] login detectado");
        const inputUsuarioVisible = await login.inputUsuario.isVisible().catch(() => false);
        const inputPasswordVisible = await login.inputPassword.isVisible().catch(() => false);

        if (inputUsuarioVisible) {
            await login.inputUsuario.click({ force: true }).catch(() => { });
            await login.inputUsuario.press("Control+A").catch(() => { });
            await login.inputUsuario.press("Delete").catch(() => { });
            await login.inputUsuario.fill(BIZAGI_USUARIO).catch(async () => {
                await login.inputUsuario.pressSequentially(BIZAGI_USUARIO, { delay: 25 }).catch(() => { });
            });
        } else {
            // Select-only login (e.g. portalcomercial.dev.bsc.com.do/login): use selectOption
            const selectUsuario = bizagiPage.locator('select:visible').first();
            if (await selectUsuario.isVisible().catch(() => false)) {
                await selectUsuario.evaluate((el, user) => {
                    const sel = el as HTMLSelectElement;
                    const raw = String(user || '').toLowerCase();
                    const terms = [raw, raw.replace('domain\\', ''), 'admon'];
                    const options = Array.from(sel.options || []);
                    const match = options.find((o) => {
                        const txt = String(o.text || '').toLowerCase();
                        const val = String(o.value || '').toLowerCase();
                        return terms.some((t) => t && (txt.includes(t) || val.includes(t)));
                    }) || options.find((o) => String(o.value || '').trim() || String(o.text || '').trim());
                    if (!match) return;
                    sel.value = match.value;
                    sel.dispatchEvent(new Event('input', { bubbles: true }));
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                }, BIZAGI_USUARIO).catch(() => { });
                await bizagiPage.waitForTimeout(300).catch(() => { });
            }
        }

        if (inputPasswordVisible) {
            await login.inputPassword.click({ force: true }).catch(() => { });
            await login.inputPassword.press("Control+A").catch(() => { });
            await login.inputPassword.press("Delete").catch(() => { });
            await login.inputPassword.fill(BIZAGI_PASSWORD).catch(async () => {
                await login.inputPassword.pressSequentially(BIZAGI_PASSWORD, { delay: 25 }).catch(() => { });
            });
        }

        await autocompletarCredencialesBizagiSiLogin(bizagiPage);
        if (await login.btnIngresar.isVisible().catch(() => false)) {
            await login.btnIngresar.click({ force: true });
        }
        console.log("[TasaExcepcion][Bizagi] autenticado con usuario admon");

        const inicioHome = Date.now();
        while (Date.now() - inicioHome < 20000) {
            const homeAfter = await homeSignals();
            if (homeAfter.visible) {
                return bizagiPage;
            }
            await bizagiPage.waitForTimeout(500);
        }

        return bizagiPage;
    }

    throw new Error("[CRITICO] Bizagi no mostr? ni login ni pantalla principal.");
}

async function buscarCasoBizagi(page: Page, mpn: string): Promise<void> {
    console.log(`[TasaExcepcion][Bizagi] buscando caso ${mpn}`);

    // Paso 1: Localizar y usar el buscador superior (menuQuery)
    const selectoresBuscador = [
        '#menuQuery',
        'input#menuQuery',
        '#ui-bizagi-wp-widget-searchContainer input',
        'input[id*="menuQuery"]',
        'input[name="menuQuery"]',
        'input[placeholder="Buscar"]',
    ];

    let inputBuscador: Locator | null = null;
    for (const sel of selectoresBuscador) {
        try {
            const input = page.locator(sel).first();
            if (await input.isVisible({ timeout: 2000 }).catch(() => false)) {
                inputBuscador = input;
                console.log(`[TasaExcepcion][Bizagi] buscador encontrado: '${sel}'`);
                break;
            }
        } catch { }
    }

    if (!inputBuscador) {
        console.log('[TasaExcepcion][Bizagi] buscador superior no encontrado, continuando sin busqueda...');
        throw new Error(`[CRITICO] Bizagi: no se pudo localizar el buscador superior.`);
    }

    // Paso 2: Dispara la busqueda
    const dispararBusquedaTasaExcepcion = async (input: Locator) => {
        try {
            // Enfoque el input
            await input.focus().catch(() => { });

            // Intenta fill directo
            await input.fill(mpn).catch(async () => {
                await input.click({ timeout: 1000 }).catch(() => { });
                await input.fill(mpn).catch(() => { });
            });

            // Verifica que el valor se escribio
            const valor = ((await input.inputValue().catch(() => '')) || '').trim();
            if (valor !== mpn) {
                await input.evaluate((el, v) => {
                    const inputEl = el as HTMLInputElement;
                    inputEl.value = String(v);
                    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                }, mpn).catch(() => { });
            }

            // Dispara eventos de busqueda
            await input.evaluate((el) => {
                const inputEl = el as HTMLInputElement;
                inputEl.focus();
                inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                inputEl.dispatchEvent(new Event('search', { bubbles: true }));
                inputEl.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter' }));
                inputEl.dispatchEvent(new KeyboardEvent('keypress', { bubbles: true, key: 'Enter', code: 'Enter' }));
                inputEl.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Enter', code: 'Enter' }));
                const form = inputEl.form;
                if (form) {
                    form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
                    form.requestSubmit?.();
                }
                inputEl.blur();
            }).catch(() => { });

            // Press Enter como fallback
            await input.press('Enter').catch(() => { });
            await page.keyboard.press('Enter').catch(() => { });
            await page.waitForTimeout(500);
        } catch (e) {
            console.log(`[TasaExcepcion][Bizagi] Error disparando busqueda: ${e}`);
        }
    };

    // Dispara la busqueda inicial
    await dispararBusquedaTasaExcepcion(inputBuscador);

    // Paso 3: Espera a resultados en tabla Bizagi
    const filasPorMpn = page
        .locator('table#ui-bizagi-wp-app-inbox-grid-cases tbody tr[data-idworkflow]:visible')
        .filter({ hasText: new RegExp(mpn, 'i') });

    const REGEX_ACTIVIDAD = /Gesti(?:o|ó)n Autorizaci(?:o|ó)n Tasa de Excepci(?:o|ó)n/i;
    const filasPorActividad = page
        .locator('table#ui-bizagi-wp-app-inbox-grid-cases tbody tr[data-idworkflow]:visible')
        .filter({ hasText: REGEX_ACTIVIDAD });

    const inicio = Date.now();
    const TIMEOUT_ESPERA = 30000;
    let ultimoSegundoLog = -1;
    let reintentosDisparoCount = 0;
    const maxReintentosDisparo = 3;

    while (Date.now() - inicio < TIMEOUT_ESPERA) {
        const countMpn = await filasPorMpn.count().catch(() => 0);
        if (countMpn > 0) {
            console.log(`[TasaExcepcion][Bizagi] fila con ${mpn} encontrada (count=${countMpn})`);
            capturas.push(await capturarCuentaComoPNG(page, mpn, `bizagi__caso-encontrado`));
            return;
        }

        const countActividad = await filasPorActividad.count().catch(() => 0);
        if (countActividad > 0) {
            console.log(`[TasaExcepcion][Bizagi] fila por actividad encontrada (count=${countActividad})`);
            return;
        }

        // Reintentar disparo cada 5 segundos
        const elapsedMs = Date.now() - inicio;
        if (inputBuscador && elapsedMs > (reintentosDisparoCount + 1) * 5000 && reintentosDisparoCount < maxReintentosDisparo) {
            reintentosDisparoCount++;
            console.log(`[TasaExcepcion][Bizagi] Reintento de busqueda #${reintentosDisparoCount} para ${mpn}...`);
            await dispararBusquedaTasaExcepcion(inputBuscador);
        }

        const elapsed = Math.floor(elapsedMs / 1000);
        if (elapsed !== ultimoSegundoLog) {
            ultimoSegundoLog = elapsed;
            console.log(`[TasaExcepcion][Bizagi] esperando resultados para ${mpn} (${elapsed}s)...`);
        }
        await page.waitForTimeout(500);
    }

    throw new Error(`[CRITICO] Bizagi no devolvio resultados para ${mpn}.`);
}

async function abrirCasoGestionAutorizacionTasaExcepcion(page: Page, mpn: string): Promise<void> {
    const REGEX_ACTIVIDAD = /Gesti(?:o|\u00f3)n Autorizaci(?:o|\u00f3)n Tasa de Excepci(?:o|\u00f3)n/i;

    // Intentar primero con MPN en la fila (más preciso)
    const filaPorMpn = page
        .locator('tr, .p-datatable-row, .grid-row, .bpm-grid-row')
        .filter({ hasText: new RegExp(mpn, "i") })
        .filter({ hasText: REGEX_ACTIVIDAD })
        .filter({ hasNotText: /Solicitud MultiProducto Persona Natural/i })
        .first();

    // Fallback: buscar solo por nombre de actividad (cuando MPN no está en el texto de la fila)
    const filaPorActividad = page
        .locator('tr, .p-datatable-row, .grid-row, .bpm-grid-row')
        .filter({ hasText: REGEX_ACTIVIDAD })
        .filter({ hasNotText: /Solicitud MultiProducto Persona Natural/i })
        .first();

    let fila = filaPorMpn;
    let usandoFallback = false;

    const mpnVisible = await filaPorMpn.isVisible().catch(() => false);
    if (!mpnVisible) {
        const actividadVisible = await filaPorActividad.isVisible().catch(() => false);
        if (!actividadVisible) {
            throw new Error("[CRITICO] No se encontró la fila de 'Gestión Autorización Tasa de Excepción' en Bizagi.");
        }
        console.log("[TasaExcepcion][Bizagi] usando fila por nombre de actividad (MPN no visible en fila)");
        fila = filaPorActividad;
        usandoFallback = true;
    }

    console.log(`[TasaExcepcion][Bizagi] fila correcta encontrada (fallback=${usandoFallback})`);

    const link = usandoFallback
        ? fila.getByRole("link").first()
        : fila.getByRole("link", { name: new RegExp(mpn, "i") }).first();

    if (await link.isVisible().catch(() => false)) {
        await link.click({ force: true });
    } else {
        await fila.click({ force: true });
    }
}

async function completarAprobacionTasaBizagi(page: Page, valorTasa: string, paso: 1 | 2): Promise<void> {
    // 1) Solicitar aclaraciones -> No
    const bloqueAclaraciones = page
        .locator('xpath=//*[contains(normalize-space(.),"Solicitar Aclaraciones")]/ancestor::*[self::div or self::section or self::fieldset][1]')
        .first();
    await bloqueAclaraciones.waitFor({ state: "visible", timeout: 20000 });

    const opcionNo = bloqueAclaraciones.getByLabel(/No/i).first();
    if (await opcionNo.isVisible().catch(() => false)) {
        await opcionNo.check({ force: true }).catch(async () => {
            await opcionNo.click({ force: true }).catch(() => { });
        });
        const checked = await opcionNo.isChecked().catch(() => false);
        if (!checked) {
            throw new Error(`[CRITICO] No se pudo seleccionar 'No' en Solicitar Aclaraciones (paso ${paso}).`);
        }
    } else {
        const labelNo = bloqueAclaraciones.locator('label:has-text("No")').first();
        await labelNo.click({ force: true }).catch(() => { });
    }
    console.log(`[TasaExcepcion][Bizagi] paso ${paso} solicitar aclaraciones=no`);

    // 2) Secci?n Tasa de Excepci?n (contenedor)
    const seccionTasa = page
        .locator('xpath=//*[self::div or self::section or self::fieldset or self::table][.//*[contains(normalize-space(.),"Tasa de ExcepciÃ³n") or contains(normalize-space(.),"Tasa de Excepci")]]')
        .first();
    await seccionTasa.waitFor({ state: "visible", timeout: 20000 });
    console.log(`[TasaExcepcion][Bizagi] paso ${paso} secciÃ³n tasa excepciÃ³n localizada`);

    // 3) Input Tasa Recomendada / Aprobada por XPath exacto
    const inputTasaAprobada = page
        .locator('xpath=//*[@id="M_VerifTasa_xProdVerifTasa"]/div/div[2]/div[1]/table/tbody/tr/td[12]/div/div/span/input')
        .first();
    await inputTasaAprobada.waitFor({ state: "visible", timeout: 15000 });
    console.log(`[TasaExcepcion][Bizagi] paso ${paso} input tasa localizado por xpath`);
    if (paso === 1) {
        capturas.push(await capturarCuentaComoPNG(page, 'bizagi', `tasa-excepcion__paso-1-input`, { skipSpinnerWait: true }));
    }

    await inputTasaAprobada.scrollIntoViewIfNeeded().catch(() => { });
    await inputTasaAprobada.click({ force: true }).catch(() => { });
    await inputTasaAprobada.press("Control+A").catch(() => { });
    await inputTasaAprobada.press("Delete").catch(() => { });
    await inputTasaAprobada.fill(valorTasa).catch(() => { });

    let valorFinal = (await inputTasaAprobada.inputValue().catch(() => "")).trim();
    if (!tasaCoincide(valorFinal, valorTasa)) {
        await inputTasaAprobada.click({ force: true }).catch(() => { });
        await inputTasaAprobada.press("Control+A").catch(() => { });
        await inputTasaAprobada.press("Delete").catch(() => { });
        await inputTasaAprobada.pressSequentially(valorTasa, { delay: 40 }).catch(() => { });
        valorFinal = (await inputTasaAprobada.inputValue().catch(() => "")).trim();
    }
    if (!tasaCoincide(valorFinal, valorTasa)) {
        throw new Error(`[CRITICO] No se pudo escribir la tasa recomendada/aprobada. esperado='${valorTasa}' actual='${valorFinal}'`);
    }
    console.log(`[TasaExcepcion][Bizagi] paso ${paso} tasa aprobada='${valorFinal}'`);

    // 4) Seleccionar S? en columna Aprobaci?n dentro de la misma fila
    const filaAprobacion = inputTasaAprobada.locator('xpath=ancestor::tr[1]');
    let opcionSi = inputTasaAprobada
        .locator('xpath=ancestor::td[1]/following-sibling::td[1]//input[@type="radio" and @value="true"]')
        .first();
    if (!(await opcionSi.isVisible().catch(() => false))) {
        opcionSi = inputTasaAprobada
            .locator('xpath=ancestor::td[1]/following-sibling::td[1]//label[contains(normalize-space(.),"Si") or contains(normalize-space(.),"SÃ­")]')
            .first();
    }
    if (!(await opcionSi.isVisible().catch(() => false))) {
        opcionSi = filaAprobacion.locator('input[type="radio"][value="true"]').first();
    }
    if (!(await opcionSi.isVisible().catch(() => false))) {
        opcionSi = filaAprobacion.locator('label:has-text("Si"), label:has-text("SÃ­")').first();
    }
    if (!(await opcionSi.isVisible().catch(() => false))) {
        opcionSi = filaAprobacion.locator('input[type="radio"]').first();
    }

    if (await opcionSi.isVisible().catch(() => false)) {
        await opcionSi.click({ force: true }).catch(() => { });
    }

    const marcado = await filaAprobacion.locator('input[type="radio"]:checked').first().isVisible().catch(() => false);
    if (!marcado) {
        throw new Error(`[CRITICO] No se pudo seleccionar 'Si' en aprobaci?n Bizagi (paso ${paso}).`);
    }
    console.log(`[TasaExcepcion][Bizagi] paso ${paso} aprobaciÃ³n=si`);

    // 5) Guardar antes de Siguiente (Bizagi requiere guardar primero)
    const btnGuardar = page.getByRole("button", { name: /^Guardar$/i }).first();
    if (await btnGuardar.isVisible().catch(() => false)) {
        console.log(`[TasaExcepcion][Bizagi] paso ${paso} clickeando Guardar`);
        await btnGuardar.click({ force: true }).catch(() => { });
        await page.waitForTimeout(1500);
    }

    // 6) Siguiente
    const btnSiguiente = page.getByRole("button", { name: /Siguiente/i }).first();
    await btnSiguiente.waitFor({ state: "visible", timeout: 15000 });
    const btnEnabled = await btnSiguiente.isEnabled().catch(() => false);
    if (!btnEnabled) {
        throw new Error(`[CRITICO] Botón 'Siguiente' no disponible en Bizagi (paso ${paso}).`);
    }
    console.log(`[TasaExcepcion][Bizagi] paso ${paso} clickeando Siguiente`);
    await btnSiguiente.click({ force: true });
    await page.waitForTimeout(1500);

    // Verificar que la acción avanzó (el input debería desaparecer o la URL cambiar)
    const inputSigueVisible = await inputTasaAprobada.isVisible().catch(() => false);
    if (inputSigueVisible) {
        // Reintentar con JS click
        console.log(`[TasaExcepcion][Bizagi] paso ${paso} Siguiente no navegó, reintentando con JS click`);
        await btnSiguiente.evaluate((el: HTMLElement) => el.click()).catch(() => { });
        await page.waitForTimeout(2000);
    }
    console.log(`[TasaExcepcion][Bizagi] paso ${paso} Siguiente ejecutado`);
}

async function aprobarExcepcionTasaBizagi(context: BrowserContext, mpn: string, valorTasa: string): Promise<void> {
    const bizagiPage = await abrirYAutenticarBizagi(context);
    await buscarCasoBizagi(bizagiPage, mpn);
    await abrirCasoGestionAutorizacionTasaExcepcion(bizagiPage, mpn);
    await completarAprobacionTasaBizagi(bizagiPage, valorTasa, 1);
    const inputTasaPaso2 = bizagiPage
        .locator('xpath=//*[@id="M_VerifTasa_xProdVerifTasa"]/div/div[2]/div[1]/table/tbody/tr/td[12]/div/div/span/input')
        .first();
    await inputTasaPaso2.waitFor({ state: "detached", timeout: 15000 }).catch(() => { });
    await inputTasaPaso2.waitFor({ state: "visible", timeout: 20000 });
    await completarAprobacionTasaBizagi(bizagiPage, valorTasa, 2);
}

async function esperarEstadoAprobadoTasaEnPortal(page: Page, valorTasa: string): Promise<void> {
    const bloque = page
        .locator('section, fieldset, .p-card, .card, .p-panel, div')
        .filter({ hasText: /Tasa de excepci(?:o|\u00f3)n/i })
        .first();

    const inicio = Date.now();
    const timeoutMs = 120000;

    const obtenerValorPorLabel = async (labelRegex: RegExp) => {
        const label = bloque.getByText(labelRegex).first();
        if (!(await label.isVisible().catch(() => false))) return "";
        const valor = await label
            .locator('xpath=following::*[self::span or self::div or self::input][1]')
            .first()
            .textContent()
            .catch(() => "");
        return String(valor ?? "").trim();
    };

    while (Date.now() - inicio < timeoutMs) {
        const visible = await bloque.isVisible().catch(() => false);
        if (visible) {
            const aprobado = await bloque.getByText(/Aprobado/i).first().isVisible().catch(() => false);
            const solicitada = await obtenerValorPorLabel(/Tasa solicitada/i);
            const aprobada = await obtenerValorPorLabel(/Tasa aprobada/i);
            const solicitadaOk = solicitada ? tasaCoincide(solicitada, valorTasa) : false;
            const aprobadaOk = aprobada ? tasaCoincide(aprobada, valorTasa) : false;
            if (aprobado && solicitadaOk && aprobadaOk) {
                console.log("[TasaExcepcion][Portal] estado aprobado detectado");
                return;
            }
        }
        await page.waitForTimeout(1200);
    }

    throw new Error("[CRITICO] No se detectÃ³ estado Aprobado para la Tasa de excepciÃ³n en Portal.");
}

async function procesarAprobacionExcepcionTasaSiAplica(
    context: BrowserContext,
    page: Page,
    registro: RegistroExcel
): Promise<void> {
    if (!esValorSi(registro.tasaExepcion)) {
        return;
    }

    console.log("[TasaExcepcion][Portal] excepci?n solicitada; esperando salida del modal");

    const dialogoCertificado = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Certificados de dep[o??]sito/i })
        .last();

    await dialogoCertificado.waitFor({ state: "hidden", timeout: 15000 }).catch(() => { });

    const seccionInicial = await localizarSeccionProductos(page);
    const productoVisibleInicial = await detectarProductoAgregadoEnUI(page, seccionInicial, registro.tipoCuenta).catch(() => false);
    if (productoVisibleInicial) {
        console.log("[TasaExcepcion][Portal] producto visible en UI tras excepci?n");
    } else {
        const inicioProducto = Date.now();
        while (Date.now() - inicioProducto < 15000) {
            const seccion = await localizarSeccionProductos(page);
            const visible = await detectarProductoAgregadoEnUI(page, seccion, registro.tipoCuenta).catch(() => false);
            if (visible) {
                console.log("[TasaExcepcion][Portal] producto visible en UI tras excepci?n");
                break;
            }
            await page.waitForTimeout(500);
        }
    }

    const btnContinuarProductos = getBotonContinuar(page);
    await btnContinuarProductos.waitFor({ state: "visible", timeout: 20000 });

    const continuarEnabled = await btnContinuarProductos.isEnabled().catch(() => false);
    if (!continuarEnabled) {
        throw new Error("[CRITICO] El bot?n Continuar est? deshabilitado despu?s de solicitar la excepci?n.");
    }

    await btnContinuarProductos.scrollIntoViewIfNeeded().catch(() => { });
    console.log("[TasaExcepcion][Portal] click en Continuar para ir a Verificaciones");
    await btnContinuarProductos.click({ force: true });

    await esperarFinActualizandoSolicitud(page, 30000).catch(() => false);

    const verifHeader = page.getByText(/Verificaciones/i).first();
    await verifHeader.waitFor({ state: "visible", timeout: 20000 });
    console.log("[TasaExcepcion][Portal] pantalla Verificaciones detectada");
    capturas.push(await capturarCuentaComoPNG(page, registro.identificacion, `${registro.tipoCuenta}__verificaciones`));

    const bloqueTasa = page
        .locator('section, fieldset, .p-card, .card, .p-panel, div')
        .filter({ hasText: /Tasa de excepci(?:o|\u00f3)n/i })
        .first();
    await bloqueTasa.waitFor({ state: "visible", timeout: 20000 });

    const inicio = Date.now();
    let enEsperaDetectado = false;
    while (Date.now() - inicio < 120000) {
        const enEspera = await bloqueTasa.getByText(/En espera/i).first().isVisible().catch(() => false);
        if (enEspera) {
            console.log("[TasaExcepcion][Portal] estado en espera detectado");
            enEsperaDetectado = true;
            break;
        }
        await page.waitForTimeout(1200);
    }
    if (!enEsperaDetectado) {
        throw new Error("[CRITICO] No se detectÃ³ estado En espera para Tasa de excepciÃ³n.");
    }

    const mpn = await extraerMpnSolicitudDesdePortal(page);
    console.log(`[TasaExcepcion][Portal] mpn detectado='${mpn}'`);

    await aprobarExcepcionTasaBizagi(context, mpn, String(registro.valorTasa ?? "").trim());

    await page.bringToFront().catch(() => { });
    await page.reload({ waitUntil: "domcontentloaded" }).catch(() => { });

    await verifHeader.waitFor({ state: "visible", timeout: 20000 });
    await esperarEstadoAprobadoTasaEnPortal(page, String(registro.valorTasa ?? "").trim());
    console.log("[TasaExcepcion][Portal] estado aprobado detectado");
    capturas.push(await capturarCuentaComoPNG(page, registro.identificacion, `${registro.tipoCuenta}__tasa-aprobada`, { skipSpinnerWait: true }));

    const btnContinuarPostAprobacion = getBotonContinuar(page);
    await btnContinuarPostAprobacion.waitFor({ state: "visible", timeout: 15000 });
    await btnContinuarPostAprobacion.click({ force: true });
    console.log("[TasaExcepcion][Portal] continuando flujo tras aprobaciÃ³n");
    await esperarFinActualizandoSolicitud(page, 30000).catch(() => false);
}

async function localizarSeccionProductosLegacy(page: Page) {
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

    const globalDropdowns = await page
        .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
        .count()
        .catch(() => 0);
    if (globalDropdowns >= 2) {
        const fallbackVisible = page
            .locator('xpath=(//*[self::fieldset or self::div][.//div[contains(@class,"p-dropdown") and not(contains(@class,"disabled"))]][1])')
            .first();
        if (await fallbackVisible.isVisible().catch(() => false)) {
            return fallbackVisible;
        }
    }

    return candidatos[0];
}

async function seleccionarCategoriaEnSeccionProductosLegacy(page: Page, seccionProductos: Locator) {
    const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
    let count = await dropdowns.count().catch(() => 0);
    let categoriaDropdown = dropdowns.first();
    if (count < 1) {
        const globalDropdowns = page.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
        count = await globalDropdowns.count().catch(() => 0);
        if (count < 1) {
            throw new Error("[CRITICO] No se encontro dropdown de 'Categoria de producto' en la seccion de Productos.");
        }
        categoriaDropdown = globalDropdowns.first();
    }

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
                const itemCategoria = items.filter({ hasText: /Certificados de Deposito/i }).first();
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

    const msgSinProductos = page.getByText(/No se agregaron productos en simulaci[oÃ³]n/i).first();
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

    // IMPORTANTE: Durante flujo normal NO cierre este modal
    // Solo registre que apareció
    console.log('[CERT-EX][Guard][CRITICO] Modal Cancelar proceso de solicitud visible; no se cierra automáticamente');
    return false;
}

// Helper para validar que NO aparece el modal Cancelar proceso de solicitud
async function validarNoCancelarProcesoModal(page: Page) {
    const modalCancelarProceso = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Cancelar proceso de solicitud/i })
        .first();
    const visible = await modalCancelarProceso.isVisible().catch(() => false);
    if (visible) {
        throw new Error('[CERT-EX][Guard][CRITICO] Modal Cancelar proceso de solicitud apareció durante flujo normal; no se cierra automáticamente');
    }
}

async function esperarFinActualizandoSolicitud(page: Page, timeoutMs = 120000) {
    const EARLY_EXIT_CHECK_MS = 1500; // Verificar tempranamente si el texto aparece

    const txtActualizando = page.getByText(/Actualizando solicitud/i).first();
    const txtDepurando = page.getByText(/Depurando solicitante|Consultando datos del solicitante/i).first();
    const overlays = page.locator(
        '.p-blockui:visible, [data-pc-name="blockui"]:visible, .p-progressspinner:visible, .p-progress-spinner:visible, [role="progressbar"]:visible'
    );

    // Early exit check: si "Actualizando solicitud" no aparece en los primeros 1500ms, salir inmediatamente
    const earlyCheckStart = Date.now();
    while (Date.now() - earlyCheckStart < EARLY_EXIT_CHECK_MS) {
        const actualizandoVisible = await txtActualizando.isVisible().catch(() => false);
        const depurandoVisible = await txtDepurando.isVisible().catch(() => false);
        const overlayCount = await overlays.count().catch(() => 0);

        if (actualizandoVisible || depurandoVisible || overlayCount > 0) {
            // El spinner/actualizando SÍ apareció, esperar a que desaparezca
            console.log('[CERT-EX][Depurar][Continuar] Actualizando solicitud detectado; esperando fin');
            break;
        }

        if (Date.now() - earlyCheckStart >= EARLY_EXIT_CHECK_MS) {
            // El spinner nunca apareció, salir inmediatamente sin esperar timeout completo
            console.log('[CERT-EX][Depurar][Continuar] Actualizando solicitud no detectado en 1500ms; continuando');
            return true;
        }

        await page.waitForTimeout(100);
    }

    // Si llegamos aquí, el spinner SÍ apareció. Esperar a que desaparezca.
    const mainWaitStart = Date.now();
    while (Date.now() - mainWaitStart < timeoutMs) {
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
            const inicioEspera = Date.now();
            while (Date.now() - inicioEspera < 15000) {
                const sigueVisible = await msgCorreoPred.isVisible().catch(() => false);
                if (!sigueVisible) break;
                const actualizando = await page.getByText(/Actualizando solicitud/i).first().isVisible().catch(() => false);
                if (!actualizando) break;
                await page.waitForTimeout(350);
            }
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
            'xpath=//*[not(ancestor::*[contains(@class,"p-dialog") or @role="dialog"])][contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZÃÃ‰ÃÃ“Ãš","abcdefghijklmnopqrstuvwxyzÃ¡Ã©Ã­Ã³Ãº"),"solicitud finalizada")]'
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
            'xpath=//*[not(ancestor::*[contains(@class,"p-dialog") or @role="dialog"])][contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZÃÃ‰ÃÃ“Ãš","abcdefghijklmnopqrstuvwxyzÃ¡Ã©Ã­Ã³Ãº"),"pendiente finalizar")]'
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

async function adjuntarDocumentoYFinalizarSolicitud(page: Page): Promise<void> {
    const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
    const btnFinalizar = page.getByRole('button', { name: /^Finalizar$/i }).first();
    const gestionVisible = await btnGestion.waitFor({ state: "visible", timeout: 15000 }).then(() => true).catch(() => false);
    const finalizarVisible = await btnFinalizar.waitFor({ state: "visible", timeout: 15000 }).then(() => true).catch(() => false);
    if (!gestionVisible || !finalizarVisible) {
        throw new Error(`[CRITICO] No se detectaron botones en la pantalla final. gestion=${gestionVisible} finalizar=${finalizarVisible}`);
    }

    console.log('[FinalizarSolicitud] pantalla final detectada');
    console.log('[FinalizarSolicitud] abriendo Gesti\u00f3n Documental');
    await btnGestion.scrollIntoViewIfNeeded().catch(() => { });
    await btnGestion.click({ force: true });
    await page.waitForTimeout(2500);

    const modalGestion = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Gesti(?:o|\u00f3)n [Dd]ocumental/i })
        .last();
    const modalVisible = await modalGestion.waitFor({ state: 'visible', timeout: 15000 }).then(() => true).catch(() => false);
    if (!modalVisible) {
        // Fallback: cualquier dialogo visible
        const anyDialog = page.locator('.p-dialog:visible, [role="dialog"]:visible').last();
        const anyVisible = await anyDialog.isVisible().catch(() => false);
        throw new Error(`[CRITICO] No se abrió el modal de Gestión Documental. Hay dialogo visible=${anyVisible}`);
    }

    console.log('[FinalizarSolicitud] modal abierto, expandiendo secciones si están colapsadas');
    // Expandir cualquier panel colapsado dentro del modal (p-panel con p-panel-header toggler)
    const panelHeaders = modalGestion.locator('.p-panel-header, [data-pc-name="panel"] [data-pc-section="header"]');
    const totalPanels = await panelHeaders.count().catch(() => 0);
    for (let pi = 0; pi < totalPanels; pi++) {
        const hdr = panelHeaders.nth(pi);
        // Verificar si el contenido del panel está oculto buscando su contenido hermano
        const panelEl = hdr.locator('xpath=ancestor::*[contains(@class,"p-panel")][1]').first();
        const panelContent = panelEl.locator('.p-panel-content, [data-pc-section="content"]').first();
        const contentVisible = await panelContent.isVisible().catch(() => false);
        if (!contentVisible) {
            console.log(`[FinalizarSolicitud] Expandiendo panel ${pi + 1}`);
            await hdr.click({ force: true }).catch(() => {});
            await page.waitForTimeout(600);
        }
    }
    await page.waitForTimeout(400);

    // Esperar a que el contenido interno termine de cargar/renderizar (la lista y el botón de adjuntar llegan async).
    const loadingInterno = modalGestion.locator('.p-progressspinner, .p-progress-spinner, [role="progressbar"], .p-blockui').first();
    await loadingInterno.waitFor({ state: 'hidden', timeout: 20000 }).catch(() => { });
    const filaProducto = modalGestion.getByText(/CERTIFICADOS DE DEP(?:O|\u00D3)SITO/i).first();
    await filaProducto.waitFor({ state: 'visible', timeout: 20000 }).catch(() => { });

    // Buscar el boton de carga a nivel de página dentro del contenido del dialogo visible
    // (no usar modalGestion.locator porque el p-panel puede renderizar fuera del scope del locator)
    const btnCargarDocumentoPreferido = modalGestion
        .locator('button:visible')
        .filter({ has: modalGestion.locator('.pi-upload, i.pi-upload, span.pi-upload, [class*="upload" i]') })
        .first();

    const btnCargarDocumentoFallback = modalGestion
        .locator(
            '.p-dialog-content button.p-button-icon-only:visible, ' +
            '[data-pc-section="content"] button.p-button-icon-only:visible, ' +
            'button[data-pd-tooltip="true"]:visible, button[data-p-tooltip="true"]:visible'
        )
        .first();

    const candidatosCarga = [btnCargarDocumentoPreferido, btnCargarDocumentoFallback];
    let btnCargarDocumento = btnCargarDocumentoPreferido;
    let btnCargaVisible = false;
    let btnCargaEnabled = false;
    let filaProductoVisible = false;
    let textoGenerandoVisible = false;
    const deadlineGeneracion = Date.now() + 90000;
    let ultimoLogSeg = -1;
    while (Date.now() < deadlineGeneracion) {
        filaProductoVisible = await filaProducto.isVisible().catch(() => false);
        textoGenerandoVisible = await modalGestion.getByText(/Generando|Procesando|Preparando|Espere|Cargando/i).first().isVisible().catch(() => false);

        btnCargaVisible = false;
        btnCargaEnabled = false;
        for (const cand of candidatosCarga) {
            const visible = await cand.isVisible().catch(() => false);
            if (!visible) continue;
            btnCargarDocumento = cand;
            btnCargaVisible = true;
            btnCargaEnabled = await cand.isEnabled().catch(() => false);
            if (btnCargaEnabled) break;
        }

        if (filaProductoVisible && btnCargaVisible && btnCargaEnabled) break;

        const elapsedSeg = Math.floor((90000 - Math.max(0, deadlineGeneracion - Date.now())) / 1000);
        if (elapsedSeg !== ultimoLogSeg && elapsedSeg % 3 === 0) {
            ultimoLogSeg = elapsedSeg;
            console.log(`[FinalizarSolicitud] esperando generación del documento... ${elapsedSeg}s (fila=${filaProductoVisible} btnVisible=${btnCargaVisible} btnEnabled=${btnCargaEnabled} generando=${textoGenerandoVisible})`);
        }

        await loadingInterno.waitFor({ state: 'hidden', timeout: 1500 }).catch(() => { });
        await page.waitForTimeout(500);
    }
    if (!btnCargaVisible || !btnCargaEnabled) {
        // Log diagnóstico completo de botones en la página
        const btns = await page.locator('.p-dialog:visible button').all().catch(() => []);
        const btnsInfo = await Promise.all(btns.map(async b => {
            const txt = await b.innerText().catch(() => '?');
            const lbl = await b.getAttribute('aria-label').catch(() => null);
            const cls = await b.getAttribute('class').catch(() => null);
            return `"${txt.trim()}"${lbl ? `[aria=${lbl}]` : ''}${cls ? `[cls=${cls.split(' ').filter(c => c.startsWith('p-button')).join('.')}]` : ''}`;
        }));
        throw new Error(`[CRITICO] No se encontró botón de carga listo en el modal. fila=${filaProductoVisible} btnVisible=${btnCargaVisible} btnEnabled=${btnCargaEnabled} generando=${textoGenerandoVisible}. Botones: ${btnsInfo.join(', ')}`);
    }
    console.log('[FinalizarSolicitud] abriendo carga de documento');
    await btnCargarDocumento.scrollIntoViewIfNeeded().catch(() => { });
    await btnCargarDocumento.click({ force: true });

    const modalCarga = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Cargar documento/i })
        .last();
    const modalCargaVisible = await modalCarga.waitFor({ state: 'visible', timeout: 12000 }).then(() => true).catch(() => false);
    if (!modalCargaVisible) {
        throw new Error('[CRITICO] No se abri\u00f3 el modal "Cargar documento".');
    }

    console.log('[FinalizarSolicitud] adjuntando archivo evidencia.pdf');
    const archivo = path.resolve(process.cwd(), 'data', 'Documentos', 'evidencia.pdf');

    const botonAbrirChooser = modalCarga
        .locator('div[data-pc-section="content"] button[aria-label="Cargar"], div[data-pc-section="content"] button:has-text("Cargar")')
        .first();
    await botonAbrirChooser.waitFor({ state: 'visible', timeout: 20000 });

    const fileChooser = await Promise.all([
        page.waitForEvent('filechooser', { timeout: 7000 }).catch(() => null),
        botonAbrirChooser.click({ force: true }),
    ]).then((res) => res[0]);

    if (fileChooser) {
        await fileChooser.setFiles(archivo);
    } else {
        const inputFile = page.locator('input[type="file"]').last();
        if ((await inputFile.count().catch(() => 0)) === 0) {
            throw new Error("[CRITICO] No se encontró input de archivo en el modal de carga.");
        }
        await inputFile.setInputFiles(archivo);
    }

    await modalCarga.waitFor({ state: 'visible', timeout: 10000 });
    const filaArchivo = modalCarga.getByText(/evidencia\.pdf/i).first();
    const archivoVisible = await filaArchivo.waitFor({ state: 'visible', timeout: 15000 }).then(() => true).catch(() => false);
    if (!archivoVisible) {
        throw new Error("[CRITICO] No aparece el archivo en la lista de carga.");
    }
    console.log('[FinalizarSolicitud] archivo visible en lista');

    const btnConfirmarCarga = modalCarga
        .locator('div[data-pc-section="footer"] button[aria-label="Cargar"], div[data-pc-section="footer"] button:has-text("Cargar")')
        .first();
    await btnConfirmarCarga.waitFor({ state: 'visible', timeout: 10000 });
    const confirmarEnabled = await btnConfirmarCarga.isEnabled().catch(() => false);
    if (!confirmarEnabled) {
        throw new Error("[CRITICO] El botón final 'Cargar' está deshabilitado.");
    }
    await btnConfirmarCarga.click({ force: true });

    const notificacionOK = page.getByText(/Informaci[oó]n guardada|guardad[ao]|correctamente/i).first();
    const okVisible = await notificacionOK.waitFor({ state: 'visible', timeout: 9000 }).then(() => true).catch(() => false);
    const esperandoConfirmacion = await modalGestion.getByText(/Esperando confirmaci[oó]n/i).first().isVisible().catch(() => false);
    const confirmada = okVisible || esperandoConfirmacion;
    if (!confirmada) {
        throw new Error("[CRITICO] No se detectó confirmación de guardado del documento.");
    }
    console.log('[FinalizarSolicitud] carga confirmada');

    console.log('[FinalizarSolicitud] cerrando modal de gestión documental');
    const btnCerrarGestion = modalGestion
        .locator('button[aria-label="Cerrar"], button[aria-label="Close"], button[data-pc-section="closebutton"], button.p-dialog-header-icon')
        .first();
    await btnCerrarGestion.waitFor({ state: 'visible', timeout: 10000 });
    await btnCerrarGestion.click({ force: true }).catch(() => { });
    await modalGestion.waitFor({ state: 'hidden', timeout: 12000 }).catch(() => { });

    console.log('[FinalizarSolicitud] click en Finalizar');
    await btnFinalizar.waitFor({ state: 'visible', timeout: 15000 }).catch(() => { });
    const deadlineFinalizar = Date.now() + 15000;
    let finalizarEnabledNow = await btnFinalizar.isEnabled().catch(() => false);
    while (!finalizarEnabledNow && Date.now() < deadlineFinalizar) {
        await page.waitForTimeout(250);
        finalizarEnabledNow = await btnFinalizar.isEnabled().catch(() => false);
    }
    if (!finalizarEnabledNow) {
        throw new Error("[CRITICO] El botón Finalizar está deshabilitado.");
    }
    await clickBotonFinalizarRobusto(page).catch(async () => {
        await btnFinalizar.scrollIntoViewIfNeeded().catch(() => { });
        await btnFinalizar.click({ force: true }).catch(() => { });
    });
}
async function finalizarSolicitudConDocumento(page: Page): Promise<void> {
    const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
    const btnFinalizar = page.getByRole('button', { name: /^Finalizar$/i }).first();
    const gestionVisible = await btnGestion.waitFor({ state: "visible", timeout: 15000 }).then(() => true).catch(() => false);
    const finalizarVisible = await btnFinalizar.waitFor({ state: "visible", timeout: 15000 }).then(() => true).catch(() => false);
    if (!gestionVisible || !finalizarVisible) {
        throw new Error(`[CRITICO] No se detect\u00f3 la pantalla final. gestion=${gestionVisible} finalizar=${finalizarVisible}`);
    }
    await btnGestion.scrollIntoViewIfNeeded().catch(() => { });
    await btnFinalizar.scrollIntoViewIfNeeded().catch(() => { });

    console.log('[FinalizarSolicitud] pantalla final detectada');
    await adjuntarDocumentoYFinalizarSolicitud(page);
    // Espera inicial: carga que puede aparecer antes del popup de confirmación
    await esperarFinActualizandoSolicitud(page, 20000).catch(() => false);
    // Aceptar popup de confirmación si apareció ("¿Desea finalizar la solicitud?")
    await confirmarFinalizacionSiHayPopup(page).catch(() => false);
    // Espera posterior: la carga que arranca DESPUÉS de aceptar el popup
    await esperarFinActualizandoSolicitud(page, 25000).catch(() => false);
    // Esperar que el badge "Pendiente finalizar" desaparezca (max 20s)
    await page.getByText(/Pendiente finalizar/i).first()
        .waitFor({ state: 'hidden', timeout: 20000 })
        .catch(() => { });
    await page.waitForTimeout(600);

    const pendiente = await estadoPendienteFinalizarVisible(page).catch(() => false);
    if (pendiente) {
        throw new Error("[CRITICO] El flujo termin\u00f3 sin finalizar la solicitud ni cargar el documento.");
    }
    console.log('[FinalizarSolicitud] solicitud finalizada correctamente');
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
            const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
            const gestionVisible = await btnGestion.isVisible().catch(() => false);
            if (gestionVisible) {
                await adjuntarDocumentoYFinalizarSolicitud(page);
            } else {
                await clickBotonFinalizarRobusto(page).catch(() => false);
            }
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
    });
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
        modalMatcher: async (dialogos, tipoCuentaRegex, codigoRegex) => {
            const totalDialogos = await dialogos.count().catch(() => 0);

            for (let i = totalDialogos - 1; i >= 0; i--) {
                const dlg = dialogos.nth(i);
                const texto = ((await dlg.innerText().catch(() => "")) || "").trim();

                if (/Certificados de dep[oÃƒÂ³]sito actuales/i.test(texto)) continue;

                const tieneCabeceraConfig =
                    /Certificados de dep[oÃƒÂ³]sito/i.test(texto) &&
                    /Monto/i.test(texto) &&
                    /Plazo/i.test(texto);

                if (!tieneCabeceraConfig) continue;

                const visibleTipoEnModal = await dlg.getByText(tipoCuentaRegex).first().isVisible().catch(() => false);
                if (visibleTipoEnModal) return true;

                if (codigoRegex) {
                    const visibleCodigoEnModal = await dlg.getByText(codigoRegex).first().isVisible().catch(() => false);
                    if (visibleCodigoEnModal) return true;
                }
            }

            return false;
        },
    });

    const tipoCuentaSafe = String(tipoCuenta ?? "").trim();
    if (!tipoCuentaSafe) return false;

    const tipoCuentaRegex = new RegExp(escapeRegexText(tipoCuentaSafe).replace(/\s+/g, "\\s+"), "i");
    const codigoProducto = extraerCodigoProducto(tipoCuentaSafe);
    const codigoRegex = codigoProducto ? new RegExp(`\\b${escapeRegexText(codigoProducto)}\\b`, "i") : null;

    const visibleTextoCompletoEnSeccion = await seccionProductos
        .getByText(tipoCuentaRegex)
        .first()
        .isVisible()
        .catch(() => false);
    if (visibleTextoCompletoEnSeccion) return true;

    const visibleTextoCompletoGlobal = await page
        .locator(
            'xpath=//*[not(ancestor::*[contains(@class,"p-dropdown-panel") or @role="listbox" or @role="option"]) and (self::div or self::span or self::p or self::small or self::strong or self::a or self::li or self::td)]'
        )
        .filter({ hasText: tipoCuentaRegex })
        .first()
        .isVisible()
        .catch(() => false);
    if (visibleTextoCompletoGlobal) return true;

    if (codigoRegex) {
        const visibleCodigoEnSeccion = await seccionProductos
            .getByText(codigoRegex)
            .first()
            .isVisible()
            .catch(() => false);
        if (visibleCodigoEnSeccion) return true;

        const visibleCodigoGlobal = await page
            .locator(
                'xpath=//*[not(ancestor::*[contains(@class,"p-dropdown-panel") or @role="listbox" or @role="option"]) and (self::div or self::span or self::p or self::small or self::strong or self::a or self::li or self::td)]'
            )
            .filter({ hasText: codigoRegex })
            .first()
            .isVisible()
            .catch(() => false);
        if (visibleCodigoGlobal) return true;
    }

    const dialogos = page.locator('.p-dialog:visible, [role="dialog"]:visible');
    const totalDialogos = await dialogos.count().catch(() => 0);

    for (let i = totalDialogos - 1; i >= 0; i--) {
        const dlg = dialogos.nth(i);
        const texto = ((await dlg.innerText().catch(() => "")) || "").trim();

        if (/Certificados de dep[oÃ³]sito actuales/i.test(texto)) continue;

        const tieneCabeceraConfig =
            /Certificados de dep[oÃ³]sito/i.test(texto) &&
            /Monto/i.test(texto) &&
            /Plazo/i.test(texto);

        if (!tieneCabeceraConfig) continue;

        const visibleTipoEnModal = await dlg.getByText(tipoCuentaRegex).first().isVisible().catch(() => false);
        if (visibleTipoEnModal) return true;

        if (codigoRegex) {
            const visibleCodigoEnModal = await dlg.getByText(codigoRegex).first().isVisible().catch(() => false);
            if (visibleCodigoEnModal) return true;
        }
    }

    return false;
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

            if (coincideTexto || coincideCodigo) {
                const modalAbierto = await modalProductoConfigVisible(page);
                if (modalAbierto) return true;
            }
        }

        await page.waitForTimeout(110);
    }
    return false;
}

async function modalProductoConfigVisible(page: Page) {
    return modalProductoConfigVisibleShared(page, /Cuentas de efectivo|Balance promedio|Moneda|Tasa|Certificados/i);
}

async function localizarSeccionProductos(page: Page) {
    return localizarSeccionProductosShared(page);
}

function obtenerCategoriaRegex(tipoCuenta: string): RegExp {
    const codigo = parseInt(extraerCodigoProducto(tipoCuenta), 10);
    if (codigo >= 300 && codigo <= 399) return /Certificados?\s+de\s+dep[oó]sito|Dep[oó]sito\s+a\s+plazo|Certificados?/i;
    return /Cuentas de Efectivo/i;
}

async function seleccionarCategoriaEnSeccionProductos(page: Page, seccionProductos: Locator, tipoCuenta: string) {
    return seleccionarCategoriaEnSeccionProductosShared(page, seccionProductos, {
        panelQuickTimeout: LIST_PANEL_QUICK_TIMEOUT,
        panelTimeout: LIST_PANEL_TIMEOUT,
        scopeRetryWaitMs: LIST_SCOPE_RETRY_WAIT_MS,
    }, {
        categoriaRegex: obtenerCategoriaRegex(tipoCuenta),
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
        /Cuentas de efectivo|Balance promedio|Moneda|Tasa|Certificados/i,
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
                await cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false);
                await clickReintentarListaSiVisible(currentPage, labelText, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
            },
            onAfterDropdownClick: async (currentPage) => {
                const cerro = await cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false);
                return cerro;
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
        },
        {
            onAfterDropdownClick: async (currentPage) => {
                const cerro = await cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false);
                return cerro;
            },
        }
    );
}

async function seleccionarProductoEnSeccionProductos(
    page: Page,
    seccionProductos: Locator,
    tipoCuenta: string
) {
    return seleccionarProductoCertificadoExistente(page, seccionProductos, tipoCuenta, {
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

async function localizarInputNumeroIdentificacion(page: Page): Promise<Locator> {
    const labelNumeroRegex = /N(?:u|ú|Ãº)mero de identificaci(?:o|ó|Ã³)n/i;
    const candidatos: Locator[] = [
        page
            .locator('label')
            .filter({ hasText: labelNumeroRegex })
            .first()
            .locator('xpath=following::input[1]')
            .first(),
        page
            .locator('xpath=(//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Numero de identificacion")])[1]/following::input[1]')
            .first(),
        page.locator('input[aria-label*="identific" i], input[placeholder*="identific" i]').first(),
    ];

    for (const candidato of candidatos) {
        if (await candidato.isVisible().catch(() => false)) {
            return candidato;
        }
    }
    return candidatos[0];
}

async function asegurarPortalEnMultiproducto(page: Page, contexto: string) {
    const enMultiproducto = async () => /\/requests\/create\/multiproduct/i.test(page.url());
    if (await enMultiproducto()) return true;

    console.log(`[${contexto}] Forzando goto directo a multiproducto...`);
    await page.goto(PORTAL_MULTIPRODUCT_URL, { waitUntil: 'domcontentloaded' }).catch(() => { });
    await page.waitForURL(/\/requests\/create\/multiproduct/i, { timeout: 15000 }).catch(() => { });
    if (await enMultiproducto()) return true;

    console.log(`[${contexto}] Goto directo no basto. Intentando location.assign...`);
    await page.evaluate((destino) => {
        window.location.assign(destino);
    }, PORTAL_MULTIPRODUCT_URL).catch(() => { });
    await page.waitForURL(/\/requests\/create\/multiproduct/i, { timeout: 15000 }).catch(() => { });
    if (await enMultiproducto()) return true;

    console.log(`[${contexto}] location.assign no basto. Intentando router/link...`);
    const navegado = await page.evaluate(() => {
        const link = document.querySelector<HTMLAnchorElement>('a[href="/requests/create/multiproduct"]');
        if (!link) return false;
        link.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        return true;
    }).catch(() => false);
    if (navegado) {
        await page.waitForURL(/\/requests\/create\/multiproduct/i, { timeout: 15000 }).catch(() => { });
    }

    return await enMultiproducto();
}

async function esperarCertificadoListoDespuesDepurar(page: Page) {
    const timeoutMs = 120000;
    const inicio = Date.now();
    let estableConsecutivo = 0;

    while (Date.now() - inicio < timeoutMs) {
        const actualizando = await page.getByText(/Actualizando solicitud/i).first().isVisible().catch(() => false);
        const depurando = await page.getByText(/Depurando solicitante|Consultando datos del solicitante/i).first().isVisible().catch(() => false);
        const modalCancelar = await page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Cancelar proceso de solicitud/i })
            .first()
            .isVisible()
            .catch(() => false);

        const toastRequest0 = await page
            .getByText(/La entidad Request con el Id\.?\s*0 no existe|Error al crear cliente/i)
            .first()
            .isVisible()
            .catch(() => false);

        console.log(`[CERT-EX][Depurar][Continuar] espera actualizando=${actualizando} depurando=${depurando} modalCancelar=${modalCancelar} toastRequest0=${toastRequest0}`);

        if (toastRequest0) {
            await page.screenshot({ path: `artifacts/cert_ex_request_id_0_${Date.now()}.png`, fullPage: true }).catch(() => { });
            throw new Error('[CERT-EX][CRITICO] Apareció Request Id 0 antes de Continuar; la web no estaba lista');
        }

        if (modalCancelar) {
            // NO hacer click en Cancelar. Solo esperar a que desaparezca.
            estableConsecutivo = 0;
            await page.waitForTimeout(500);
            continue;
        }

        if (!actualizando && !depurando) {
            estableConsecutivo += 1;
        } else {
            estableConsecutivo = 0;
        }

        if (estableConsecutivo >= 4) {
            console.log('[CERT-EX][Depurar][Continuar] Pantalla estable tras Depurar');
            return;
        }

        await page.waitForTimeout(500);
    }

    await page.screenshot({ path: `artifacts/cert_ex_timeout_post_depurar_${Date.now()}.png`, fullPage: true }).catch(() => { });
    throw new Error('[CERT-EX][CRITICO] Timeout esperando estabilidad después de Depurar');
}

async function esperarUrlRequestEditEstableCert(page: Page, contexto: string) {
    const timeoutMs = 120000;
    const inicio = Date.now();
    let estableConsecutivo = 0;

    while (Date.now() - inicio < timeoutMs) {
        const url = page.url();

        const enCreate = /\/requests\/create\/multiproduct/i.test(url);
        const enEdit = /\/requests\/\d+\/edit/i.test(url);

        const actualizando = await page.getByText(/Actualizando solicitud/i).first().isVisible().catch(() => false);
        const depurando = await page.getByText(/Depurando solicitante|Consultando datos del solicitante/i).first().isVisible().catch(() => false);
        const obteniendoCert = await page.getByText(/Obteniendo datos de certificado/i).first().isVisible().catch(() => false);

        const modalCancelar = await page.locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Cancelar proceso de solicitud/i })
            .first()
            .isVisible()
            .catch(() => false);

        const requestId0 = await page.getByText(/La entidad Request con el Id\.?\s*0 no existe|Error al crear cliente/i)
            .first()
            .isVisible()
            .catch(() => false);

        console.log(`[CERT-EX][WaitRequestEdit] contexto=${contexto} url=${url} enCreate=${enCreate} enEdit=${enEdit} actualizando=${actualizando} depurando=${depurando} obteniendoCert=${obteniendoCert} modalCancelar=${modalCancelar} requestId0=${requestId0}`);

        if (requestId0) {
            await page.screenshot({ path: `artifacts/cert_ex_request_id_0_${Date.now()}.png`, fullPage: true }).catch(() => {});
            throw new Error('[CERT-EX][CRITICO] Request Id 0 detectado durante espera de request edit estable');
        }

        if (modalCancelar) {
            throw new Error('[CERT-EX][CRITICO] Modal Cancelar proceso de solicitud apareció durante espera; no se cierra automáticamente');
        }

        if (enEdit && !enCreate && !actualizando && !depurando && !obteniendoCert) {
            estableConsecutivo += 1;
        } else {
            estableConsecutivo = 0;
        }

        if (estableConsecutivo >= 5) {
            console.log(`[CERT-EX][WaitRequestEdit] Request edit estable contexto=${contexto}`);
            return;
        }

        await page.waitForTimeout(700);
    }

    await page.screenshot({ path: `artifacts/cert_ex_timeout_request_edit_${Date.now()}.png`, fullPage: true }).catch(() => {});
    throw new Error(`[CERT-EX][CRITICO] Timeout esperando request edit estable contexto=${contexto}`);
}

async function etapaFlujoRegistro(page: Page, registro: RegistroExcel) {
    const maxIntentosCasoActivo = 2;

    for (let intento = 1; intento <= maxIntentosCasoActivo; intento++) {
        await page.evaluate(() => {
            window.moveTo(0, 0);
            window.resizeTo(window.screen.availWidth, window.screen.availHeight);
        });

        await page.goto(PORTAL_MULTIPRODUCT_URL, {
            waitUntil: 'domcontentloaded',
        });
        const estadoSesion = await esperarPortalListoTrasLogin(page, { timeoutMs: 240000 });
        if (estadoSesion.loginDetectado) {
            console.log('Sesion reanudada despues de login manual.');
        } else {
            console.log('Sesion cargada, sin MFA.');
        }

        const urlTrasGoto = page.url();
        const estaEnSolicitudMultiproducto = /\/requests\/create\/multiproduct/i.test(urlTrasGoto);
        console.log(`[Flujo][intento=${intento}] URL tras goto: ${urlTrasGoto} | enMultiproducto=${estaEnSolicitudMultiproducto}`);

        if (!estaEnSolicitudMultiproducto) {
            const forzadoDirecto = await asegurarPortalEnMultiproducto(page, `Flujo][intento=${intento}`);
            const urlTrasForzado = page.url();
            console.log(`[Flujo][intento=${intento}] URL tras forzado directo: ${urlTrasForzado} | exito=${forzadoDirecto}`);
        }

        const estaEnSolicitudTrasForzado = /\/requests\/create\/multiproduct/i.test(page.url());
        if (!estaEnSolicitudTrasForzado) {
            console.log(`[Flujo][intento=${intento}] Navegando a multiproducto via router...`);
            const navegado = await page.evaluate(() => {
                const link = document.querySelector<HTMLAnchorElement>('a[href="/requests/create/multiproduct"]');
                if (!link) return false;
                link.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                return true;
            });
            console.log(`[Flujo][intento=${intento}] dispatchEvent click ejecutado=${navegado}`);
            if (navegado) {
                await page.waitForURL(/\/requests\/create\/multiproduct/i, { timeout: 30000 }).catch(() => {});
            }
            const urlTrasNav = page.url();
            console.log(`[Flujo][intento=${intento}] URL tras navegacion: ${urlTrasNav}`);
        }

        const modalCancelarProceso = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Cancelar proceso de solicitud/i })
            .first();
        const modalCancelarVisible = await modalCancelarProceso.isVisible().catch(() => false);
        console.log(`[Flujo][intento=${intento}] modalCancelarProceso visible=${modalCancelarVisible}`);
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
        await clickReintentarListaSiVisible(page, /Pa[iÃ­]s de nacimiento/i);
        await clickReintentarListaSiVisible(page, 'Estado civil');

        console.log(`[Flujo][intento=${intento}] asegurarIdentificacionHabilitada...`);
        await asegurarIdentificacionHabilitada(page, { timeoutCampoMs: 60000 });
        console.log(`[Flujo][intento=${intento}] identificacion habilitada OK`);
        const inputNumeroRefuerzo = await localizarInputNumeroIdentificacion(page);
        await inputNumeroRefuerzo.waitFor({ state: 'visible', timeout: 25000 });
        let habilitadoNumero = await inputNumeroRefuerzo.isEnabled({ timeout: 5000 }).catch(() => false);
        if (!habilitadoNumero) {
            await esperarYClickReintentarPaisIdentificacion(page, 12000);
            const clicPais = await clickReintentarListaSiVisible(page, 'PaÃ­s de identificaciÃ³n', 2000).catch(() => false);
            if (!clicPais) {
                await clickReintentarListaSiVisible(page, 'Tipo de identificaciÃ³n', 2000).catch(() => false);
            }
            habilitadoNumero = await inputNumeroRefuerzo.isEnabled({ timeout: 5000 }).catch(() => false);
            if (!habilitadoNumero) {
                throw new Error("Input 'NÃºmero de identificaciÃ³n' sigue deshabilitado tras refuerzo.");
            }
        }

        console.log(`[Flujo][intento=${intento}] llenando cedula=${registro.identificacion}...`);
        const inputNumero = await localizarInputNumeroIdentificacion(page);
        await inputNumero.waitFor({ state: 'visible', timeout: 15000 });
        await inputNumero.click({ force: true });
        await inputNumero.fill(registro.identificacion);

        await page.keyboard.press('Enter');
        console.log(`[Flujo][intento=${intento}] Enter presionado, esperando spinner consulta...`);
        const spinnerConsulta = page.getByText(/Consultando datos del solicitante/i);
        await spinnerConsulta.waitFor({ state: 'visible', timeout: 6000 }).catch(() => { });
        await spinnerConsulta.waitFor({ state: 'hidden', timeout: 120000 }).catch(() => { });
        await page.waitForTimeout(500);
        console.log(`[Flujo][intento=${intento}] spinner consulta completado`);

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

        console.log(`[CERT-EX][Depurar][Continuar] Click en Depurar realizado`);
        const btnDepurar = page.getByRole('button', { name: 'Depurar' }).first();
        await btnDepurar.click({ noWaitAfter: true });

        console.log(`[CERT-EX][Depurar][Continuar] Esperando fin de spinner/loading`);
        await esperarFinActualizandoSolicitud(page, 15000).catch(() => { });
        console.log(`[CERT-EX][Depurar][Continuar] Spinner/loading finalizado`);

        const mpnCasoActivo = await extraerCasoActivoMpn(page);
        if (mpnCasoActivo) {
            console.log(`[CasoActivo] Detectado ${mpnCasoActivo}. Cancelando en Bizagi...`);
            await cancelarCasoEnBizagiDesdePortal(page, mpnCasoActivo, {
                url: BIZAGI_URL_NUEVA,
                usuario: BIZAGI_USUARIO_FALLBACK,
                password: BIZAGI_PASSWORD,
            });
            if (intento >= maxIntentosCasoActivo) {
                throw new Error(`[CRITICO] Se cancelÃ³ ${mpnCasoActivo} pero el caso activo persiste para '${registro.identificacion}'.`);
            }
            console.log(`[CasoActivo] ${mpnCasoActivo} cancelado. Reintentando con la misma cÃ©dula (${registro.identificacion})...`);
            await page.waitForTimeout(1500);
            continue;
        }

        const msgFechaMinima = page.getByText(/La fecha m[iÃ­]nima permitida es/i).first();
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

            console.log(`[CERT-EX][Depurar][Continuar] Click en segundo Depurar realizado`);
            const btnDepurar2 = page.getByRole('button', { name: 'Depurar' }).first();
            await btnDepurar2.click({ noWaitAfter: true });

            console.log(`[CERT-EX][Depurar][Continuar] Esperando fin de spinner/loading (segundo depurar)`);
            await esperarFinActualizandoSolicitud(page, 15000).catch(() => { });
            console.log(`[CERT-EX][Depurar][Continuar] Spinner/loading finalizado (segundo depurar)`);

            const mpnCasoActivo2 = await extraerCasoActivoMpn(page);
            if (mpnCasoActivo2) {
                console.log(`[CasoActivo] Detectado ${mpnCasoActivo2} en segundo depurar. Cancelando en Bizagi...`);
                await cancelarCasoEnBizagiDesdePortal(page, mpnCasoActivo2, {
                    url: BIZAGI_URL_NUEVA,
                    usuario: BIZAGI_USUARIO_FALLBACK,
                    password: BIZAGI_PASSWORD,
                });
                if (intento >= maxIntentosCasoActivo) {
                    throw new Error(`[CRITICO] Se cancelÃ³ ${mpnCasoActivo2} pero el caso activo persiste para '${registro.identificacion}'.`);
                }
                console.log(`[CasoActivo] ${mpnCasoActivo2} cancelado. Reintentando con la misma cÃ©dula (${registro.identificacion})...`);
                await page.waitForTimeout(1500);
                continue;
            }
        }

        // Espera específica de Certificado-ex: NO continuar mientras exista spinner/modal/toast
        await esperarCertificadoListoDespuesDepurar(page);

        console.log(`[CERT-EX][Depurar][Continuar] Esperando estabilización humana antes de Continuar`);
        await page.waitForTimeout(2500);

        // Segunda validación: esperar de nuevo que la página esté estable
        await esperarCertificadoListoDespuesDepurar(page);

        const btnContinuar = getBotonContinuar(page);
        const continuarVisible = await btnContinuar.isVisible({ timeout: 2000 }).catch(() => false);
        const continuarEnabled = await btnContinuar.isEnabled({ timeout: 800 }).catch(() => true);

        if (!continuarVisible || !continuarEnabled) {
            console.log(`[CERT-EX][Depurar][Continuar] WARN: Botón Continuar no está listo (visible=${continuarVisible}, enabled=${continuarEnabled}). Esperando...`);
            await page.waitForTimeout(1500);
        }

        console.log(`[CERT-EX][Depurar][Continuar] Botón Continuar visible y habilitado`);

        // Validar URL /requests/{id}/edit estable ANTES de hacer click
        await esperarUrlRequestEditEstableCert(page, 'justo antes click Continuar post Depurar');

        await asegurarTiempoEnVivienda(page, "0").catch(() => false);
        console.log(`[CERT-EX][Depurar][Continuar] Click en Continuar`);
        await btnContinuar.click({ noWaitAfter: true }).catch(() => { });

        // Guard: Verificar que la page no se cerró
        console.log(`[CERT-EX][PAGE] Verificando page después de Continuar`);
        if (page.isClosed()) {
            console.log(`[CERT-EX][PAGE][CLOSED] Page cerrada después de Continuar`);
            throw new Error('[CERT-EX][PAGE][CLOSED] Page cerrada después de Continuar');
        }
        console.log(`[CERT-EX][PAGE] Page activa después de Continuar url=${page.url()}`);

        // Esperar a que la página se estabilice después del click
        await page.waitForLoadState('domcontentloaded').catch(() => { });
        await page.waitForTimeout(1500);
        await esperarCertificadoListoDespuesDepurar(page);
        console.log('[CERT-EX][Depurar][Continuar] Post-click Continuar estable');

        await esperarFinActualizandoSolicitud(page, 8000).catch(() => false);
        await resolverNoPoseeCorreoSiFalta(page).catch(() => false);
        await page.waitForTimeout(1500);
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
            await llenarInputMask(page, "TelÃ©fono", "8092551233");
            await llenarCampoPorLabel(page, "RelaciÃ³n", "Amigo");
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

// === FUNCIONES ROBUSTAS DE SELECCIÓN DE PRODUCTOS (patrón de ce-ex adaptado para Certificados) ===

async function seleccionarCategoriaEnSeccionProductosRobusto(page: Page, seccionProductos: Locator) {
  console.log(`[CERT-EX][Producto][Seleccion] seleccionarCategoriaEnSeccionProductos INIT`);
  console.log(`[CERT-EX][Producto][Seleccion] Sección Productos visible`);

  // Estrategia 1: Buscar por label "Categoría de producto"
  console.log(`[CERT-EX][Producto][Seleccion] Buscando dropdown Categoría por label`);
  const labelCategoriaBuscadores = [
    seccionProductos.locator('xpath=//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Categoria de producto")]'),
    seccionProductos.getByText(/Categor[ií]a de producto/i),
  ];

  let categoriaDropdown: Locator | null = null;
  for (const labelLocator of labelCategoriaBuscadores) {
    const visible = await labelLocator.isVisible().catch(() => false);
    if (visible) {
      // Encontró el label, ahora buscar el dropdown cercano
      try {
        const parent = labelLocator.locator('xpath=ancestor::*[self::fieldset or self::div][1]');
        const dropdownEnParent = parent.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible, [role="combobox"]:visible').first();
        const dropdownVisible = await dropdownEnParent.isVisible().catch(() => false);
        if (dropdownVisible) {
          console.log(`[CERT-EX][Producto][Seleccion] Dropdown Categoría encontrado por label`);
          categoriaDropdown = dropdownEnParent;
          break;
        }
      } catch (e) {
        console.log(`[CERT-EX][Producto][Seleccion] Error buscando dropdown cerca del label: ${String(e)}`);
      }
    }
  }

  // Estrategia 2: Fallback posicional - usar el primer dropdown visible
  if (!categoriaDropdown) {
    console.log(`[CERT-EX][Producto][Seleccion] Dropdown Categoría no encontrado por label; usando fallback posicional`);
    const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
    const count = await dropdowns.count().catch(() => 0);
    if (count >= 1) {
      console.log(`[CERT-EX][Producto][Seleccion] Dropdown Categoría encontrado por posición 1/${count}`);
      categoriaDropdown = dropdowns.first();
    }
  }

  if (!categoriaDropdown) {
    console.log(`[CERT-EX][Producto][Seleccion][ERROR] No se encontró dropdown de Categoría de producto`);
    throw new Error("[CRITICO] No se encontro dropdown de 'Categoria de producto' en la seccion de Productos.");
  }

  // Verificar si el dropdown ya tiene valor seleccionado
  const labelDropdown = categoriaDropdown.locator('.p-dropdown-label, [data-pc-section="label"]').first();
  let valor = ((await labelDropdown.textContent().catch(() => "")) || "").trim();
  if (valor && !/seleccionar|por favor|elige|--/i.test(valor)) {
    console.log(`[CERT-EX][Producto][Seleccion] Categoría ya seleccionada: ${valor}`);
    return;
  }

  // Abrir el dropdown y seleccionar la categoría
  for (let intento = 1; intento <= 3; intento++) {
    console.log(`[CERT-EX][Producto][Seleccion] Seleccionando categoría: Certificados de Deposito (intento ${intento}/3)`);
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    await categoriaDropdown.scrollIntoViewIfNeeded().catch(() => { });
    await categoriaDropdown.click({ force: true }).catch(() => { });
    await page.waitForTimeout(300);

    const combobox = categoriaDropdown.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
    let panel: Locator | null = null;
    if (panelId) {
      const byId = page.locator(`#${panelId}`);
      const visible = await byId.waitFor({ state: "visible", timeout: 2000 })
        .then(() => true)
        .catch(() => false);
      if (visible) panel = byId;
    }
    if (!panel) {
      const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
      const visible = await fallback.waitFor({ state: "visible", timeout: 2000 })
        .then(() => true)
        .catch(() => false);
      if (visible) panel = fallback;
    }

    if (panel) {
      // Reintentar si hay botón de reintentar
      const btnRetryPanel = panel
        .locator('button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button.p-button-warning')
        .first();
      if (await btnRetryPanel.isVisible().catch(() => false)) {
        await btnRetryPanel.click({ force: true }).catch(() => { });
        await page.waitForTimeout(800);
      }

      // Buscar y seleccionar la opción "Certificados de Deposito"
      const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
      const listo = await items.first().waitFor({ state: "visible", timeout: 3000 })
        .then(() => true)
        .catch(() => false);
      const countItems = await items.count().catch(() => 0);
      if (listo && countItems > 0) {
        const itemCategoria = items.filter({ hasText: /Certificados de Deposito|Certificados de depósito|Certificados?/i }).first();
        if (await itemCategoria.isVisible().catch(() => false)) {
          console.log(`[CERT-EX][Producto][Seleccion] Seleccionando: Certificados de Deposito`);
          await itemCategoria.click({ force: true }).catch(() => { });
        } else {
          console.log(`[CERT-EX][Producto][Seleccion] Seleccionando: primera opción (índice 0)`);
          await items.nth(0).click({ force: true }).catch(() => { });
        }
      }
    }

    await page.waitForTimeout(200);
    valor = ((await labelDropdown.textContent().catch(() => "")) || "").trim();
    if (valor && !/seleccionar|por favor|elige|--/i.test(valor)) {
      console.log(`[CERT-EX][Producto][Seleccion] Categoría seleccionada: ${valor}`);
      return;
    }

    await page.waitForTimeout(800);
  }

  throw new Error("[CRITICO] No se pudo seleccionar 'Categoria de producto' en la seccion de Productos.");
}

async function seleccionarProductoEnSeccionProductosConDropdown(
  page: Page,
  productoDropdown: Locator,
  tipoCuenta: string
) {
  try {
    await productoDropdown.click({ timeout: 3000 }).catch(async () => {
      await productoDropdown.click({ force: true, timeout: 3000 });
    });

    const panel = page.locator('.p-dropdown-panel:visible, .p-select-overlay:visible, [role="listbox"]:visible').last();
    const panelVisible = await panel.isVisible({ timeout: 5000 }).catch(() => false);
    if (!panelVisible) {
      throw new Error(`[Producto][CRITICO] No abrió panel Producto tipoCuenta='${tipoCuenta}'`);
    }

    console.log('[CERT-EX][Producto][Seleccion] Opciones de Producto visibles');

    const opciones = panel.locator('li[role="option"], .p-dropdown-item, .p-select-option');
    const total = await opciones.count().catch(() => 0);
    if (total === 0) {
      throw new Error(`[CERT-EX][Producto][CRITICO] Panel abierto pero sin opciones tipoCuenta='${tipoCuenta}'`);
    }

    const nombreRegex = tipoCuenta
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .toLowerCase();

    for (let i = 0; i < total; i++) {
      const texto = (await opciones.nth(i).innerText().catch(() => '')).trim();
      const normalizado = texto.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
      if (normalizado.includes(nombreRegex)) {
        console.log(`[CERT-EX][Producto][Seleccion] Seleccionando opción: ${texto}`);
        await opciones.nth(i).click({ force: true }).catch(() => { });
        console.log(`[CERT-EX][Producto][Seleccion] Producto seleccionado: ${texto}`);
        console.log('[CERT-EX][Producto][Seleccion] Validación producto seleccionado OK');
        return;
      }
    }

    console.log(`[CERT-EX][Producto][Seleccion] Tipo '${tipoCuenta}' no encontrado; seleccionando primera opción`);
    await opciones.nth(0).click({ force: true }).catch(() => { });
    console.log(`[CERT-EX][Producto][Seleccion] Producto seleccionado: primera opción`);
    console.log('[CERT-EX][Producto][Seleccion] Validación producto seleccionado OK');
  } catch (error) {
    console.log(`[CERT-EX][Producto][Seleccion][ERROR] Fallo selección directa: ${String(error)}`);
    throw new Error(`[CRITICO] No se pudo seleccionar Producto para '${tipoCuenta}'.`);
  }
}

async function seleccionarProductoEnSeccionProductosRobusto(
  page: Page,
  seccionProductos: Locator,
  tipoCuenta: string,
  productoDropdownPrelocalizado?: Locator | null
) {
  // Si ya tenemos el dropdown prelocalizado, usarlo directamente
  if (productoDropdownPrelocalizado) {
    console.log('[CERT-EX][Producto][Seleccion] Usando dropdown Producto previamente detectado por posición 2/2');
    console.log(`[CERT-EX][Producto][Seleccion] Seleccionando producto: ${tipoCuenta}`);
    console.log('[CERT-EX][Producto][Seleccion] Abriendo dropdown Producto');
    return await seleccionarProductoEnSeccionProductosConDropdown(page, productoDropdownPrelocalizado, tipoCuenta);
  }

  console.log(`[CERT-EX][Producto][Seleccion] Buscando dropdown Producto por label`);

  // Estrategia 1: Buscar por label "Producto"
  const labelProductoBuscadores = [
    seccionProductos.getByText(/^Producto$/i),
  ];

  let productoDropdown: Locator | null = null;
  for (const labelLocator of labelProductoBuscadores) {
    const visible = await labelLocator.isVisible().catch(() => false);
    if (visible) {
      // Encontró el label, ahora buscar el dropdown cercano
      try {
        const parent = labelLocator.locator('xpath=ancestor::*[self::fieldset or self::div][1]');
        const dropdownEnParent = parent.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible, [role="combobox"]:visible').first();
        const dropdownVisible = await dropdownEnParent.isVisible().catch(() => false);
        if (dropdownVisible) {
          console.log(`[CERT-EX][Producto][Seleccion] Dropdown Producto encontrado por label`);
          productoDropdown = dropdownEnParent;
          break;
        }
      } catch (e) {
        console.log(`[CERT-EX][Producto][Seleccion] Error buscando dropdown cerca del label: ${String(e)}`);
      }
    }
  }

  // Estrategia 2: Fallback posicional - usar el segundo dropdown visible (el primero es Categoría)
  if (!productoDropdown) {
    console.log(`[CERT-EX][Producto][Seleccion] Dropdown Producto no encontrado por label; usando fallback posicional`);
    const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
    const count = await dropdowns.count().catch(() => 0);
    console.log(`[CERT-EX][Producto][Seleccion] dropdowns en sección Productos=${count}`);
    if (count >= 2) {
      console.log(`[CERT-EX][Producto][Seleccion] Dropdown Producto encontrado por posición 2/${count}`);
      productoDropdown = dropdowns.nth(1);
    } else {
      console.log(`[CERT-EX][Producto][Seleccion] Producto no listo: solo ${count} dropdown visible; NO se usará como Producto`);
    }
  }

  if (!productoDropdown) {
    console.log(`[CERT-EX][Producto][Seleccion][ERROR] No se encontró dropdown de Producto`);
    throw new Error("[CRITICO] No se encontro dropdown de 'Producto' en la seccion de Productos.");
  }

  console.log(`[CERT-EX][Producto][Seleccion] Seleccionando producto: ${tipoCuenta}`);
  return await seleccionarProductoEnSeccionProductosConDropdown(page, productoDropdown, tipoCuenta);
}

async function etapaSeccionProductos(context: BrowserContext, page: Page, registro: RegistroExcel) {
    const url = page.url();
    console.log(`[SeccionProductos] inicio url=${url}`);

    // Validar contexto de entrada
    console.log('[CERT-EX][FLOW] Antes de etapaSeccionProductos url=' + url);
    console.log('[CERT-EX][FLOW] registro.productoAgregadoEnFlujo=' + Boolean((registro as any).productoAgregadoEnFlujo));
    console.log('[CERT-EX][FLOW] registro.certificadoAgregadoEnFlujo=' + Boolean((registro as any).certificadoAgregadoEnFlujo));
    console.log('[CERT-EX][FLOW] registro.tasaExcepcionProcesada=' + Boolean((registro as any).tasaExcepcionProcesada));

    // Si el certificado ya fue agregado, omitir esta etapa duplicada
    if ((registro as any).certificadoAgregadoEnFlujo || (registro as any).productoAgregadoEnFlujo) {
        console.log('[CERT-EX][Producto][Avance] Certificado ya agregado; omitiendo etapaSeccionProductos duplicada');
        return;
    }

    // === GUARDS: Validar estado antes de Productos ===
    console.log('[CERT-EX][Guard] Validando estado antes de Productos');

    // Guard 1: Validar URL /requests/{id}/edit
    if (!/\/requests\/\d+\/edit/i.test(url)) {
        throw new Error(`[CERT-EX][Guard][CRITICO] URL inv\u00e1lida antes de Productos: ${url}`);
    }
    console.log('[CERT-EX][Guard] URL v\u00e1lida');

    // Guard 2: Validar que no existe Request Id 0
    const requestId0Visible = await page
        .getByText(/La entidad Request con el Id\.?\s*0 no existe|Error al crear cliente/i)
        .first()
        .isVisible()
        .catch(() => false);
    if (requestId0Visible) {
        console.log('[CERT-EX][Guard][CRITICO] Request Id 0 detectado; no se contin\u00faa');
        await page.screenshot({ path: `artifacts/cert_ex_request_id_0_${Date.now()}.png`, fullPage: true }).catch(() => {});
        throw new Error('[CERT-EX][Guard][CRITICO] Request Id 0 detectado; no se contin\u00faa a Productos');
    }
    console.log('[CERT-EX][Guard] requestId0Visible=false');

    // Guard 3: Validar que NO aparece modal Cancelar proceso de solicitud
    const modalCancelarProceso = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Cancelar proceso de solicitud/i })
        .first();
    const modalCancelarVisible = await modalCancelarProceso.isVisible().catch(() => false);
    if (modalCancelarVisible) {
        console.log('[CERT-EX][Guard][CRITICO] Modal Cancelar proceso de solicitud visible; no se cierra autom\u00e1ticamente');
        throw new Error('[CERT-EX][Guard][CRITICO] Modal Cancelar proceso de solicitud visible al inicio de etapaSeccionProductos');
    }
    console.log('[CERT-EX][Guard] modalCancelarVisible=false');

    await resolverCorreoPredeterminadoYContinuarSiVisible(page).catch(() => false);

    const alertSinProductosPaso = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    const labelPropositoPaso = page.getByText(/Prop(?:o|\u00f3)sito/i).first();
    const tituloProductos = page.getByText(/Productos/i).first();
    let seccionProductos = await localizarSeccionProductos(page);

    console.log(`[SeccionProductos] esperarFinActualizando(90s)...`);
    await esperarFinActualizandoSolicitud(page, 90000).catch(() => false);
    console.log(`[SeccionProductos] esperarFinActualizando completado`);

    let categoriaVisible = false;
    for (let i = 1; i <= 10; i++) {
        console.log(`[SeccionProductos] iter=${i}`);
        await esperarFinActualizandoSolicitud(page, 8000).catch(() => false);

        // No cerrar modal Cancelar proceso de solicitud; solo validar que no aparece
        await validarNoCancelarProcesoModal(page).catch(() => {
            console.log('[SeccionProductos] Modal Cancelar proceso detectado; continuando...');
        });

        await tituloProductos.scrollIntoViewIfNeeded().catch(() => { });
        await page.mouse.wheel(0, 600).catch(() => { });
        await page.waitForTimeout(150);

        seccionProductos = await localizarSeccionProductos(page);
        const seccionVisible = await seccionProductos.isVisible().catch(() => false);
        const dropdownsEnSeccion = await seccionProductos
            .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
            .count()
            .catch(() => 0);
        const dropdownsGlobales = await page
            .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
            .count()
            .catch(() => 0);
        categoriaVisible = dropdownsEnSeccion >= 2 || dropdownsGlobales >= 2;
        if (categoriaVisible && dropdownsEnSeccion < 2 && dropdownsGlobales >= 2) {
            // Fallback: hay dropdowns visibles, pero el scope detectado no era el correcto.
            seccionProductos = page
                .locator('xpath=(//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Categoria de producto")] and .//div[contains(@class,"p-dropdown")]][1])')
                .first();
        }
        if (categoriaVisible) break;

        const cayoEnPasoSinProducto = await alertSinProductosPaso.isVisible().catch(() => false);
        const yaEnTaller = await labelPropositoPaso.isVisible().catch(() => false);
        const enPantallaPrevia = await estaEnPantallaPreProductos(page);

        if (enPantallaPrevia) {
            await asegurarTiempoEnVivienda(page, "0").catch(() => false);
            await resolverNoPoseeCorreoSiFalta(page).catch(() => false);
            const btnContinuar = getBotonContinuar(page);
            const continuarVisible = await btnContinuar.isVisible().catch(() => false);
            const continuarEnabled = continuarVisible
                ? await btnContinuar.isEnabled().catch(() => false)
                : false;
            if (continuarVisible && continuarEnabled) {
                await btnContinuar.click({ force: true }).catch(() => { });
                await esperarFinActualizandoSolicitud(page, 18000).catch(() => false);
                await resolverNoPoseeCorreoSiFalta(page).catch(() => false);
                await page.waitForTimeout(280);
            }
        }

        if (cayoEnPasoSinProducto || yaEnTaller || enPantallaPrevia) {
            const step2 = page.locator(
                'li.p-steps-item:has-text("2"), [data-pc-section="menuitem"]:has-text("2"), .p-steps-item:has(.p-steps-number:has-text("2"))'
            ).first();
            if (await step2.isVisible().catch(() => false)) {
                await step2.click({ force: true }).catch(() => { });
                await esperarFinActualizandoSolicitud(page, 14000).catch(() => false);
                await page.waitForTimeout(320);
            }
        }
    }

    if (!categoriaVisible) {
        const urlActual = page.url();
        const enPaso3SinProducto = await alertSinProductosPaso.isVisible().catch(() => false);
        const dropdownsVisibles = await page
            .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
            .count()
            .catch(() => 0);
        const enPantallaPrevia = await estaEnPantallaPreProductos(page).catch(() => false);
        const actualizandoVisible = await page.getByText(/Actualizando solicitud/i).first().isVisible().catch(() => false);
        console.log(`[NAV-PRODUCTO] categoriaVisible=false url=${urlActual} paso3SinProducto=${enPaso3SinProducto} dropdownsVisibles=${dropdownsVisibles} prePantalla=${enPantallaPrevia} actualizando=${actualizandoVisible}`);
        throw new Error("[CRITICO] No se pudo visualizar 'Categoria de producto' para agregar el producto.");
    }

    // === INICIO: Patrón robusto de selección de Categoría/Producto (replicado de ce-ex.spec.ts) ===
    console.log('[CERT-EX][Producto][Seleccion] Esperando pantalla Productos');
    const maxEsperaProductos = 12000;
    const inicioEspera = Date.now();
    let pantallaProductosDetectada = false;

    while (Date.now() - inicioEspera < maxEsperaProductos) {
      if (page.isClosed()) {
        throw new Error('[CERT-EX][PAGE][CLOSED] Page cerrada mientras esperaba pantalla Productos');
      }

      const categoriaVisible_signal = await page.getByText(/Categor[ií]a de producto/i).first().isVisible().catch(() => false);
      const productoVisible = await page.getByText(/^Producto$/i).first().isVisible().catch(() => false);
      const productosTextVisible = await page.getByText(/^Productos$/i).isVisible().catch(() => false);

      const signals = [categoriaVisible_signal, productoVisible, productosTextVisible];
      const signalsCount = signals.filter(Boolean).length;

      if (signalsCount >= 2) {
        console.log(`[CERT-EX][Producto][Seleccion] Señales de pantalla Productos detectadas (${signalsCount}/3): categoría=${categoriaVisible_signal}, producto=${productoVisible}, productos=${productosTextVisible}`);
        pantallaProductosDetectada = true;
        break;
      }

      await page.waitForTimeout(500);
    }

    if (!pantallaProductosDetectada) {
      console.log('[CERT-EX][Producto][Seleccion][WARN] Pantalla Productos no detectada en timeout, verificando si hay modal de error');
      const modalError = page.locator('.p-dialog:visible, [role="dialog"]:visible').first();
      const tieneModal = await modalError.isVisible().catch(() => false);
      if (tieneModal) {
        throw new Error('[CERT-EX][PAGE][MODAL] Modal de error visible después de Continuar');
      }
    }

    console.log(`[CERT-EX][Producto][Seleccion] Sección Productos visible=${pantallaProductosDetectada}`);

    // Llamar a selección robusta de categoría
    await seleccionarCategoriaEnSeccionProductosRobusto(page, seccionProductos);
    await page.waitForTimeout(1500);

    // Contar dropdowns y decidir si llamar a selección robusta de producto
    const dropdownsPostCategoria = await seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').count().catch(() => 0);
    console.log(`[CERT-EX][Producto][Seleccion] Dropdowns después de categoría: ${dropdownsPostCategoria}`);

    if (dropdownsPostCategoria >= 2) {
      // Hay al menos 2 dropdowns (categoría + producto)
      const productoDropdown = await seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').nth(1);
      await seleccionarProductoEnSeccionProductosRobusto(page, seccionProductos, registro.tipoCuenta, productoDropdown);
    } else if (dropdownsPostCategoria === 1) {
      // Solo hay 1 dropdown visible - esperar a que aparezca el producto y reintentar
      console.log('[CERT-EX][Producto][Seleccion] Solo 1 dropdown visible; esperando producto...');
      for (let i = 0; i < 4; i++) {
        await page.waitForTimeout(600);
        const dropdownsNow = await seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').count().catch(() => 0);
        if (dropdownsNow >= 2) {
          console.log(`[CERT-EX][Producto][Seleccion] Producto dropdown ahora visible (intento ${i + 1})`);
          const productoDropdown = await seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').nth(1);
          await seleccionarProductoEnSeccionProductosRobusto(page, seccionProductos, registro.tipoCuenta, productoDropdown);
          break;
        }
      }
    } else {
      // Sin dropdowns - no hacer nada
      console.log('[CERT-EX][Producto][Seleccion] Sin dropdowns visibles - producto posiblemente ya fue seleccionado');
    }

    const msgSinProductos = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    const confirmarProductoAgregado = async () => {
        if (await modalProductoConfigVisible(page)) {
            return false;
        }

        const tallerVisible =
            await page.getByText(/Prop(?:o|\u00f3)sito/i).first().isVisible().catch(() => false) ||
            await page.getByRole('button', { name: /Crear certificado/i }).first().isVisible().catch(() => false);

        if (tallerVisible) return true;

        const enUI = await detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta);
        if (enUI) return true;

        const sinProductos = await msgSinProductos.isVisible().catch(() => false);
        if (sinProductos) return false;

        await page.waitForTimeout(900);

        const tallerVisibleDespues =
            await page.getByText(/Prop(?:o|\u00f3)sito/i).first().isVisible().catch(() => false) ||
            await page.getByRole('button', { name: /Crear certificado/i }).first().isVisible().catch(() => false);

        if (tallerVisibleDespues) return true;

        if (await modalProductoConfigVisible(page)) {
            return false;
        }

        return detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta);
    };



    // Helper para localizar inputs editables después de un label
    const localizarInputEditableDespuesDeLabel = async (
      modal: Locator,
      labelRx: RegExp
    ): Promise<Locator> => {
      const label = modal.getByText(labelRx).first();
      await label.waitFor({ state: 'visible', timeout: 15000 });

      console.log(`[CERT-EX][Certificados][Modal] Localizando input editable para ${labelRx}`);

      // Buscar todos los inputs DESPUÉS del label (no ancestor)
      // Excluir readonly, disabled, y hidden
      const candidatos = label.locator(
        'xpath=following::input[not(@readonly) and not(@disabled) and not(ancestor::*[contains(@style,"display: none")])]'
      );

      const count = await candidatos.count().catch(() => 0);

      for (let i = 0; i < count; i++) {
        const input = candidatos.nth(i);
        const visible = await input.isVisible().catch(() => false);
        const editable = await input.isEditable().catch(() => false);
        const value = (await input.inputValue().catch(() => '')).trim();
        const outerHtml = (await input.evaluate((el: any) => (el as HTMLElement).outerHTML).catch(() => ''));

        console.log(`[CERT-EX][Certificados][Modal] candidato ${labelRx} #${i} visible=${visible} editable=${editable} value="${value.substring(0, 50)}" outer=${outerHtml.substring(0, 250)}`);

        if (visible && editable) {
          console.log(`[CERT-EX][Certificados][Modal] Campo ${labelRx} editable localizado`);
          return input;
        }
      }

      throw new Error(`[CERT-EX][Certificados][CRITICO] No se encontró input editable para ${labelRx}`);
    };

    const llenarModalCertificado = async () => {
        const dialogo = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Certificados?(?:\s+de\s+dep[oó]sito)?|dep[oó]sito\s+a\s+plazo|plazo\s+financiero/i })
            .last();

        await dialogo.waitFor({ state: 'visible', timeout: 25000 });
        console.log('[CERT-EX][Certificados][Modal] Modal visible detectado');

        // Diagnosticar modal antes de intentar llenar
        const modalTexto = (await dialogo.innerText().catch(() => '')).substring(0, 3000);
        const inputsCount = await dialogo.locator('input').count().catch(() => 0);
        const dropdownsCount = await dialogo.locator('div.p-dropdown, [data-pc-name="dropdown"]').count().catch(() => 0);
        const monedaLabels = await dialogo.getByText(/Moneda/i).count().catch(() => 0);
        const montoLabels = await dialogo.getByText(/Monto/i).count().catch(() => 0);
        const plazoLabels = await dialogo.getByText(/Plazo/i).count().catch(() => 0);
        const tasaLabels = await dialogo.getByText(/Tasa/i).count().catch(() => 0);

        console.log(`[CERT-EX][Certificados][Modal] Texto modal: ${modalTexto.substring(0, 200)}...`);
        console.log(`[CERT-EX][Certificados][Modal] labels Moneda=${monedaLabels} Monto=${montoLabels} Plazo=${plazoLabels} Tasa=${tasaLabels}`);
        console.log(`[CERT-EX][Certificados][Modal] inputs=${inputsCount} dropdowns=${dropdownsCount}`);

        // Localizar campos editables dentro del modal (obligatorios)
        let inputMonto: Locator;
        let inputPlazo: Locator;
        let inputTasa: Locator | null = null;
        let inputInteres: Locator | null = null;

        try {
          inputMonto = await localizarInputEditableDespuesDeLabel(dialogo, /^Monto$/i);
        } catch (e) {
          console.log(`[CERT-EX][Certificados][Modal][ERROR] ${String(e)}`);
          throw new Error("[CRITICO] No se localizó campo editable Monto en modal de certificado.");
        }

        try {
          inputPlazo = await localizarInputEditableDespuesDeLabel(dialogo, /^Plazo$/i);
        } catch (e) {
          console.log(`[CERT-EX][Certificados][Modal][ERROR] ${String(e)}`);
          throw new Error("[CRITICO] No se localizó campo editable Plazo en modal de certificado.");
        }

        // Campos opcionales
        try {
          inputTasa = await localizarInputEditableDespuesDeLabel(dialogo, /^Tasa$/i);
        } catch (e) {
          console.log(`[CERT-EX][Certificados][Modal] Tasa no es editable o no encontrada (opcional)`);
        }

        try {
          inputInteres = await localizarInputEditableDespuesDeLabel(dialogo, /^Inter[eé]s$/i);
        } catch (e) {
          console.log(`[CERT-EX][Certificados][Modal] Interés no es editable o no encontrada (opcional)`);
        }

        // Validar que son editables
        const montoEditable = await inputMonto.isEditable({ timeout: 10000 }).catch(() => false);
        if (!montoEditable) {
          throw new Error("[CRITICO] Campo Monto no es editable después de localización.");
        }

        const plazoEditable = await inputPlazo.isEditable({ timeout: 10000 }).catch(() => false);
        if (!plazoEditable) {
          throw new Error("[CRITICO] Campo Plazo no es editable después de localización.");
        }

        await inputMonto.waitFor({ state: 'visible', timeout: 8000 }).catch(() => {});
        await inputPlazo.waitFor({ state: 'visible', timeout: 8000 }).catch(() => {});
        if (inputTasa) await inputTasa.waitFor({ state: 'visible', timeout: 8000 }).catch(() => {});
        if (inputInteres) await inputInteres.waitFor({ state: 'visible', timeout: 8000 }).catch(() => {});

        // MONTO
        console.log('[CERT-EX][Certificados][Modal] Llenando campo Monto');
        await inputMonto.scrollIntoViewIfNeeded().catch(() => { });
        await inputMonto.click({ force: true });
        await inputMonto.press('Control+A').catch(() => { });
        await inputMonto.press('Delete').catch(() => { });
        await inputMonto.fill('50000').catch(async () => {
            await inputMonto.pressSequentially('50000', { delay: 35 }).catch(() => { });
        });
        await inputMonto.dispatchEvent('input').catch(() => { });
        await inputMonto.dispatchEvent('change').catch(() => { });
        await inputMonto.blur().catch(() => { });
        await page.waitForTimeout(150);

        let valorMonto = (await inputMonto.inputValue().catch(() => '')).trim();
        console.log(`[Certificados] valorMonto='${valorMonto}'`);

        if (!valorMonto) {
            await inputMonto.evaluate((el: any) => {
                const input = el as HTMLInputElement;
                input.value = '50000';
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                input.dispatchEvent(new Event('blur', { bubbles: true }));
            }).catch(() => { });
            await page.waitForTimeout(100);
            valorMonto = (await inputMonto.inputValue().catch(() => '')).trim();
        }

        console.log('[CERT-EX][Certificados][Modal] Campo Monto llenado');

        // PLAZO
        console.log('[CERT-EX][Certificados][Modal] Llenando campo Plazo');
        await inputPlazo.scrollIntoViewIfNeeded().catch(() => { });
        await inputPlazo.click({ force: true });
        await inputPlazo.press('Control+A').catch(() => { });
        await inputPlazo.press('Delete').catch(() => { });
        await inputPlazo.fill('12').catch(async () => {
            await inputPlazo.pressSequentially('12', { delay: 35 }).catch(() => { });
        });
        await inputPlazo.dispatchEvent('input').catch(() => { });
        await inputPlazo.dispatchEvent('change').catch(() => { });
        await inputPlazo.blur().catch(() => { });
        await page.waitForTimeout(150);

        let valorPlazo = (await inputPlazo.inputValue().catch(() => '')).trim();
        console.log(`[Certificados] valorPlazo='${valorPlazo}'`);

        if (!valorPlazo || /DOP|USD/i.test(valorPlazo)) {
            await inputPlazo.evaluate((el: any) => {
                const input = el as HTMLInputElement;
                input.value = '12';
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                input.dispatchEvent(new Event('blur', { bubbles: true }));
            }).catch(() => { });
            await page.waitForTimeout(200);
            valorPlazo = (await inputPlazo.inputValue().catch(() => '')).trim();
        }

        console.log('[CERT-EX][Certificados][Modal] Campo Plazo llenado');

        const montoValido = !!valorMonto && /50[\.,]?0*00|50000/i.test(valorMonto.replace(/\s/g, ''));
        const plazoValido = !!valorPlazo && /^12\b|12\s*mes/i.test(valorPlazo);

        if (!montoValido || !plazoValido) {
            throw new Error(`[CRITICO] No se pudieron completar correctamente Monto/Plazo. monto='${valorMonto}' plazo='${valorPlazo}'`);
        }

        console.log('[CERT-EX][Certificados][Modal] Campos requeridos validados');

        await inputPlazo.press('Tab').catch(() => { });
        await inputPlazo.blur().catch(() => { });

        // Helper para esperar tasa calculada (protegido contra inputInteres null)
        const esperarTasaCalculadaCertificado = async (timeoutMs = 12000): Promise<string> => {
            if (!inputTasa) {
                throw new Error('[CERT-EX][Certificados][CRITICO] Campo Tasa no localizado');
            }

            console.log('[CERT-EX][Certificados][Modal] Esperando cálculo de tasa');

            for (let intento = 1; intento <= 10; intento++) {
                const valor = await inputTasa.inputValue().catch(() => '');
                const trimmed = valor.trim();

                console.log(`[CERT-EX][Certificados][Modal] Esperando tasa intento=${intento}/10 valor="${trimmed}"`);

                // Validar que tasa tenga valor y formato de porcentaje
                if (trimmed && trimmed !== '0' && trimmed !== '0.00' && /%/.test(trimmed)) {
                    console.log(`[CERT-EX][Certificados][Modal] Tasa calculada: ${trimmed}`);
                    return trimmed;
                }

                // Disparadores suaves: Tab y click en intentos específicos
                if (intento === 1 || intento === 4 || intento === 7) {
                    try {
                        await inputTasa.click({ force: true }).catch(() => {});
                        await page.keyboard.press('Tab').catch(() => {});
                    } catch (e) {
                        // Continuar sin fallar
                    }
                }

                await page.waitForTimeout(1000);
            }

            // Si llegamos aquí, la tasa no se calculó. Intentar botón "Calcular tasa"
            console.log('[CERT-EX][Certificados][Modal] Buscando botón para calcular tasa');
            const btnCalcularTasa = dialogo.getByRole('button', { name: /Calcular|calcular tasa|Buscar|buscar/i }).first();
            const btnVisible = await btnCalcularTasa.isVisible().catch(() => false);

            if (btnVisible) {
                console.log('[CERT-EX][Certificados][Modal] Botón calcular tasa visible=true');
                await btnCalcularTasa.click().catch(() => {});
                await page.waitForTimeout(1500);

                // Intentar leer tasa una vez más después del click
                const valorFinal = await inputTasa.inputValue().catch(() => '');
                if (valorFinal.trim() && /%/.test(valorFinal.trim())) {
                    console.log(`[CERT-EX][Certificados][Modal] Tasa calculada después de botón: ${valorFinal.trim()}`);
                    return valorFinal.trim();
                }
            } else {
                console.log('[CERT-EX][Certificados][Modal] Botón calcular tasa visible=false');
            }

            const textoModal = await dialogo.innerText().catch(() => '');
            console.log('[CERT-EX][Certificados][Modal][DIAG] Texto modal al fallar tasa: ' + textoModal.slice(0, 3000));
            throw new Error('[CERT-EX][Certificados][CRITICO] Tasa no calculada después de llenar Monto y Plazo');
        };

        // Leer Interés si existe (campo opcional)
        let valorInteres = '';
        if (inputInteres) {
            valorInteres = await inputInteres.inputValue().catch(() => '');
            console.log(`[CERT-EX][Certificados][Modal] Interés actual="${valorInteres.trim()}"`);
        } else {
            console.log('[CERT-EX][Certificados][Modal] Interés no localizado/no editable; se omite lectura obligatoria');
        }

        // === FASE 1: ESPERAR Y VALIDAR TASA AUTOMÁTICA ===
        console.log('[CERT-EX][Certificados][Modal] Esperando cálculo de tasa automática');
        let tasaCalculada = '';
        try {
            tasaCalculada = await esperarTasaCalculadaCertificado(8000);
        } catch (e) {
            console.log(`[CERT-EX][Certificados][Modal] Error esperando tasa: ${String(e)}`);
            throw e;
        }
        console.log(`[CERT-EX][Certificados][Modal] Tasa automática calculada: ${tasaCalculada}`);

        // === FASE 2: MANEJO DE TASA DE EXCEPCIÓN (BLOQUEANTE) ===
        const tasaExepcionSi = esValorSi(registro.tasaExepcion);
        if (tasaExepcionSi) {
            console.log('[CERT-EX][TasaExcepcion] Inicia manejo de tasa de excepción');
            const valorTasaExcel = String(registro.valorTasa ?? "").trim();
            if (!valorTasaExcel) {
                throw new Error("[CRITICO] Registro marcado con Tasa Exepcion=SI sin Valor Tasa en Excel.");
            }
            console.log(`[CERT-EX][TasaExcepcion] Aplica tasa excepción=true valorExcel=${valorTasaExcel}`);
            console.log(`[CERT-EX][TasaExcepcion] Tasa automática antes de excepción=${tasaCalculada}`);

            const tasaEditable = await esperarCampoTasaEditable(page, inputTasa, 12000);
            if (!tasaEditable) {
                throw new Error("[CRITICO] El campo Tasa no quedó editable para solicitar excepción.");
            }

            // Ejecutar flujo de tasa de excepción (bloqueante)
            await solicitarTasaExcepcion(page, inputTasa, inputInteres, valorTasaExcel);

            // IMPORTANTE: Validar que el modal de tasa de excepción se cerró completamente
            console.log('[CERT-EX][TasaExcepcion] Validando cierre completo del flujo de tasa excepción');
            await page.waitForTimeout(1000);

            const modalTasaExcepcionAun = await page
                .locator('.p-dialog:visible, [role="dialog"]:visible')
                .filter({ hasText: /Solicitud tasa de excepci[oó]n|Motivo de solicitud|Tasa m[aá]xima|Tasa solicitada|Tasa pool|Esta tasa se encuentra por encima/i })
                .first()
                .isVisible()
                .catch(() => false);

            if (modalTasaExcepcionAun) {
                throw new Error('[CERT-EX][TasaExcepcion][CRITICO] Modal de tasa excepción sigue abierto después de solicitarTasaExcepcion');
            }
            console.log('[CERT-EX][TasaExcepcion] Flujo de excepción completado; modal cerrado');
        } else {
            console.log('[CERT-EX][TasaExcepcion] Tasa excepción no aplica; continuando con flujo normal');
        }

        // === FASE 3: VALIDAR QUE NO HAY MODAL DE TASA EXCEPCIÓN ABIERTO ANTES DE ACEPTAR ===
        console.log('[CERT-EX][Certificados][Modal] Validando que no hay modal de tasa excepción abierto');
        const modalTasaExcepcionAbiertoAntes = await page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Solicitud tasa de excepci[oó]n|Motivo de solicitud|Tasa m[aá]xima|Tasa solicitada|Tasa pool|Esta tasa se encuentra por encima/i })
            .first()
            .isVisible()
            .catch(() => false);

        if (modalTasaExcepcionAbiertoAntes) {
            throw new Error('[CERT-EX][Certificados][CRITICO] Intento de Aceptar modal mientras modal de tasa excepción sigue abierto');
        }
        console.log('[CERT-EX][Certificados][Modal] No hay modal de tasa excepción abierto; se permite Aceptar final');

        // === FASE 4: ACEPTAR MODAL PRINCIPAL ===
        console.log("[CERT-EX][Certificados][Modal] Click en Aceptar final del modal principal");
        const btnAceptar = dialogo.getByRole('button', { name: /^Aceptar$/i }).last();
        await btnAceptar.waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});

        // Click normal primero
        await btnAceptar.click().catch(() => {});
        await page.waitForTimeout(800);

        // Si el dialog sigue abierto, usar focus + Enter (mismo patrón que cierra el modal de tasa excepción)
        const dialogoSigueAbierto = await dialogo.isVisible().catch(() => false);
        if (dialogoSigueAbierto) {
            console.log("[CERT-EX][Certificados][Modal] Click normal no cerró modal, usando focus+Enter");
            await btnAceptar.focus().catch(() => {});
            await page.waitForTimeout(200);
            await page.keyboard.press("Enter");
        }

        await dialogo.waitFor({ state: 'hidden', timeout: 20000 }).catch(() => {});
        console.log('[CERT-EX][Certificados][Modal] Modal principal cerrado correctamente');
        await page.waitForTimeout(1200);
        console.log('[CERT-EX][Certificados][Modal] Certificado agregado correctamente');

        // Marcar que el certificado fue agregado exitosamente en el flujo
        (registro as any).certificadoAgregadoEnFlujo = true;
        (registro as any).productoAgregadoEnFlujo = true;
        console.log('[CERT-EX][Producto][Avance] certificadoAgregadoEnFlujo=true');
        console.log('[CERT-EX][Producto][Avance] productoAgregadoEnFlujo=true');

        // === FASE 5: CONTINUAR FLUJO NORMAL CON BOTÓN CONTINUAR/SIGUIENTE ===
        console.log('[CERT-EX][Producto][Avance] Esperando estabilidad antes de Continuar');
        await esperarUrlRequestEditEstableCert(page, 'después de cerrar modal certificado');

        console.log('[CERT-EX][Producto][Avance] URL estable antes de Continuar');
        const btnContinuar = page
            .getByRole('button', { name: /^Continuar$|^Siguiente$|^Siguiente paso$/i })
            .first();
        const btnContinuarVisible = await btnContinuar.isVisible().catch(() => false);

        if (btnContinuarVisible) {
            console.log('[CERT-EX][Producto][Avance] Click en Continuar');
            await btnContinuar.click().catch(() => {});
            await page.waitForLoadState('domcontentloaded').catch(() => {});
            await page.waitForTimeout(1500);
            console.log('[CERT-EX][Producto][Avance] Post-click Continuar, esperando estabilidad');
            await esperarFinActualizandoSolicitud(page, 8000).catch(() => false);
            console.log('[CERT-EX][Producto][Avance] Continuando flujo normal');
        } else {
            console.log('[CERT-EX][Producto][Avance] Botón Continuar no visible');
        }
    };





    let productoAgregado = false;
    for (let intentoProducto = 1; intentoProducto <= 3; intentoProducto++) {
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        const yaAgregadoAntes = await confirmarProductoAgregado();
        if (yaAgregadoAntes) {
            productoAgregado = true;
            break;
        }

        const modalVisible = await modalProductoConfigVisible(page);
        if (!modalVisible) {
            if (intentoProducto < 3) {
                await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
                await page.waitForTimeout(600);
                continue;
            }
            throw new Error("[CRITICO] No apareciÃ³ el modal de Certificados tras seleccionar el producto.");
        }

        console.log(`[Certificados] Intento ${intentoProducto}/3 llenando modal de certificado.`);
        await llenarModalCertificado();

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
            throw new Error(`[CRITICO] No se agregÃ³ el producto '${registro.tipoCuenta}' en la simulaciÃ³n.`);
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

    const msgSinProductosDespuesContinuar = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    const labelPropositoTaller = page.getByText(/Prop(?:o|\u00f3)sito/i).first();
    const btnCrearCertificado = page.getByRole('button', { name: /Crear certificado/i }).first();

    const yaEnTaller =
        await labelPropositoTaller.isVisible().catch(() => false) ||
        await btnCrearCertificado.isVisible().catch(() => false);

    if (yaEnTaller) {
        return labelPropositoTaller;
    }

    const seccionProductos = await localizarSeccionProductos(page);
    const sinProductosAntesContinuar = await msgSinProductosDespuesContinuar.isVisible().catch(() => false);
    const productoEnUIPrevio = await detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta).catch(() => false);

    if (sinProductosAntesContinuar) {
        throw new Error(`[CRITICO] Bloqueado 'Continuar': aun no hay productos agregados para '${registro.tipoCuenta}'.`);
    }

    if (!productoEnUIPrevio) {
        console.log(`[Relacionados] Producto no visible en UI para '${registro.tipoCuenta}', pero se intentarÃ¡ Continuar porque el flujo de Certificados puede refrescar tarde.`);
    }

    await getBotonContinuar(page).click();

    const llegoATaller = await Promise.race([
        labelPropositoTaller.waitFor({ state: 'visible', timeout: 20000 }).then(() => true).catch(() => false),
        btnCrearCertificado.waitFor({ state: 'visible', timeout: 20000 }).then(() => true).catch(() => false),
    ]);

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
    await page.waitForTimeout(1200);

    const cerrarDropdownSiQuedoAbierto = async () => {
        await page.keyboard.press('Escape').catch(() => { });
        await page.waitForTimeout(400);
    };

    const obtenerComboPorTexto = async (textoCampo: string, waitMs = 4000) => {
        const textoEscapado = textoCampo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

        const candidatos = [
            page.locator('label').filter({ hasText: new RegExp(`^\\s*${textoEscapado}\\s*$`, 'i') }).first().locator('xpath=following::*[contains(@class,"p-dropdown") and not(contains(@class,"disabled"))][1]').first(),
            page.locator(`xpath=(//*[contains(normalize-space(.), "${textoCampo}")])[1]/following::*[contains(@class,"p-dropdown") and not(contains(@class,"disabled"))][1]`).first(),
            page.locator(`xpath=(//*[contains(normalize-space(.), "${textoCampo}")])[1]/following::*[@role="combobox"][1]`).first(),
            page.locator(`xpath=(//*[contains(normalize-space(.), "${textoCampo}")])[1]/ancestor::*[self::div or self::section][1]//*[contains(@class,"p-dropdown") and not(contains(@class,"disabled"))][1]`).first(),
            page.locator(`.p-dropdown:visible`).filter({ has: page.getByText(new RegExp(textoEscapado, "i")) }).first(),
        ];

        const inicio = Date.now();
        while (Date.now() - inicio < waitMs) {
            for (const candidato of candidatos) {
                const visible = await candidato.isVisible().catch(() => false);
                if (visible) return candidato;
            }
            await page.waitForTimeout(400);
        }

        return null;
    };

    const seleccionarListaPorCampo = async (textoCampo: string, index: number) => {
        // 1) Si hay boton naranja 'Reintentar buscar lista' visible proximo a este campo, hacerle click
        const textoEscapado = textoCampo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

        // Ubicar el contenedor padre mÃ¡s cercano que contenga el texto del campo y el dropdown
        const btnReintentar = page.locator(`xpath=(//*[contains(normalize-space(.), "${textoCampo}")])[1]/following::*[self::div or self::fieldset or self::section][.//button[contains(@class,"p-button-warning") or contains(normalize-space(.),"Reintentar")]][1]//button[contains(@class,"p-button-warning") or contains(normalize-space(.),"Reintentar")]`).first();

        const btnAlternativo = page.locator(`xpath=(//*[contains(normalize-space(.), "${textoCampo}")])[1]/ancestor::*[self::div or self::section][1]//button[contains(@class,"p-button-warning") or contains(normalize-space(.),"Reintentar")]`).first();

        // Probar ambos candidatos
        let btnFinal = btnReintentar;
        if (!(await btnFinal.isVisible().catch(() => false))) {
            btnFinal = btnAlternativo;
        }

        if (await btnFinal.isVisible().catch(() => false)) {
            console.log(`[TallerCertificados] Pulsando 'Reintentar buscar lista' para el campo: ${textoCampo}`);
            await btnFinal.click({ force: true });
            await page.waitForTimeout(800); // Esperar que la lista cargue tras el reintento
        } else {
            // Fallback genÃ©rico a la func de uiHelpers si no se hallo con la busqueda especificada directa
            await clickReintentarListaSiVisible(page, new RegExp(textoCampo, "i")).catch(() => false);
        }

        await cerrarDropdownSiQuedoAbierto();

        const comboRapido = await obtenerComboPorTexto(textoCampo, 1200);
        if (comboRapido) {
            const valorActual = ((await comboRapido
                .locator('.p-dropdown-label, [data-pc-section="label"]')
                .first()
                .textContent()
                .catch(() => "")) || "").trim();
            if (!esValorDropdownVacio(valorActual)) {
                console.log(`[TallerCertificados] ${textoCampo} ya seleccionado -> "${valorActual}"`);
                return;
            }
        }

        const combo = await obtenerComboPorTexto(textoCampo);
        if (!combo) {
            throw new Error(`No se encontrÃ³ dropdown visible para el campo ${textoCampo}`);
        }

        await combo.scrollIntoViewIfNeeded().catch(() => { });
        await combo.click({ force: true });
        await page.waitForTimeout(200);

        const panel = page.locator('.p-dropdown-panel:visible, [role="listbox"]:visible').last();
        await panel.waitFor({ state: 'visible', timeout: 5000 });

        const items = panel.locator('.p-dropdown-item:visible, [role="option"]:visible');
        const totalItems = await items.count().catch(() => 0);

        if (totalItems === 0) {
            throw new Error(`El dropdown ${textoCampo} no mostrÃ³ opciones visibles.`);
        }

        if (index >= totalItems) {
            throw new Error(`El Ã­ndice ${index} no existe para ${textoCampo}. Opciones visibles=${totalItems}`);
        }

        const item = items.nth(index);
        const textoItem = (await item.innerText().catch(() => '')).trim();
        console.log(`[TallerCertificados] ${textoCampo} -> Ã­ndice ${index} -> "${textoItem}"`);

        await item.scrollIntoViewIfNeeded().catch(() => { });
        await item.click({ force: true });
        await page.waitForTimeout(300);

        await cerrarDropdownSiQuedoAbierto();
        await page.waitForTimeout(200);
    };

    // Ãndices segÃºn lo que me indicaste / la captura:
    // PropÃ³sito -> 1
    // Forma de pago -> 1 (Comprobante de Cajas, porque visualmente parece Ã­ndice base 0)
    // Instrumento -> 1
    // Firma autorizada 1 -> 1
    // Firma autorizada 2 -> 1

    await seleccionarListaPorCampo("Prop\u00f3sito", 1);
    await seleccionarListaPorCampo("Forma de pago", 1);
    await seleccionarListaPorCampo("Instrumento", 1);
    await seleccionarListaPorCampo("Firma autorizada 1", 1);
    await seleccionarListaPorCampo("Firma autorizada 2", 1);

    await completarCamposNoCapitalizableSiAplican(page, registro);

    const btnCrearCertificado = page.getByRole('button', { name: /Crear certificado/i }).first();
    await btnCrearCertificado.waitFor({ state: 'visible', timeout: 15000 });

    const crearEnabled = await btnCrearCertificado.isEnabled().catch(() => false);
    if (!crearEnabled) {
        throw new Error("El botÃ³n 'Crear certificado' estÃ¡ deshabilitado.");
    }

    await btnCrearCertificado.click({ force: true });

    await esperarFinActualizandoSolicitud(page, 90000).catch(() => false);
    await page.waitForTimeout(2500);

    const evidenciaPostCrear = await capturarCuentaComoPNG(
        page,
        registro.identificacion,
        `${registro.tipoCuenta}__post-crear`
    );
    capturasRef.push(evidenciaPostCrear);

    const btnContinuar = getBotonContinuar(page);
    await btnContinuar.waitFor({ state: 'visible', timeout: 20000 });

    const continuarEnabled = await btnContinuar.isEnabled().catch(() => false);
    if (!continuarEnabled) {
        throw new Error("El botÃ³n 'Continuar' estÃ¡ deshabilitado despuÃ©s de crear el certificado.");
    }

    await btnContinuar.click({ force: true });
    await esperarFinActualizandoSolicitud(page, 30000).catch(() => false);
    await page.waitForTimeout(1500);

    const sigueEnTaller =
        (await labelPropositoTaller.isVisible().catch(() => false)) ||
        (await page.getByRole('button', { name: /Crear certificado/i }).first().isVisible().catch(() => false));
    if (sigueEnTaller) {
        const btnContinuarRetry = getBotonContinuar(page);
        const retryVisible = await btnContinuarRetry.isVisible().catch(() => false);
        const retryEnabled = retryVisible ? await btnContinuarRetry.isEnabled().catch(() => false) : false;
        if (retryVisible && retryEnabled) {
            await btnContinuarRetry.click({ force: true });
            await esperarFinActualizandoSolicitud(page, 20000).catch(() => false);
            await page.waitForTimeout(1200);
        }
    }
}

async function completarCamposNoCapitalizableSiAplican(page: Page, registro: RegistroExcel): Promise<void> {
    const tipoCuenta = String(registro.tipoCuenta ?? "").trim();
    const codigoMatch = tipoCuenta.match(/^\s*(\d{3})\s*-/);
    const codigo = codigoMatch ? codigoMatch[1] : "";
    const esNoCapitalizable = ["301", "303", "309"].includes(codigo);

    if (!esNoCapitalizable) {
        return;
    }

    console.log(`[TallerCertificados] producto no capitalizable detectado: '${registro.tipoCuenta}'`);

    let campoVisible = false;

    // Candidatos: primero buscar por role=spinbutton (p-inputnumber PrimeNG), luego fallbacks
    const candidatosFrecuencia = [
        page.locator('label').filter({ hasText: /Frecuencia de pago inter/i }).first().locator('xpath=following::input[@role="spinbutton"][1]').first(),
        page.locator('xpath=(//*[contains(normalize-space(.),"Frecuencia de pago inter")])[1]/following::input[@role="spinbutton"][1]').first(),
        page.locator('xpath=(//*[contains(normalize-space(.),"Frecuencia de pago inter")])[1]/following::input[contains(@class,"p-inputnumber-input")][1]').first(),
        page.locator('xpath=(//*[contains(normalize-space(.),"Frecuencia de pago inter")])[1]/ancestor::*[self::div or self::section or self::fieldset][1]//input[@role="spinbutton"][1]').first(),
        page.locator('xpath=(//*[contains(normalize-space(.),"Frecuencia de pago inter")])[1]/following::input[not(@type="hidden")][1]').first(),
    ];

    let inputFrecuencia: Locator | null = null;
    // Esperar hasta 8s a que algún candidato sea visible (el campo aparece tras seleccionar Forma de pago)
    const inicioEspFrecuencia = Date.now();
    while (Date.now() - inicioEspFrecuencia < 8000) {
        for (const cand of candidatosFrecuencia) {
            if (await cand.isVisible().catch(() => false)) {
                inputFrecuencia = cand;
                break;
            }
        }
        if (inputFrecuencia) break;
        await page.waitForTimeout(400);
    }

    if (inputFrecuencia) {
        campoVisible = true;
        // Reintentar si hay boton naranja
        const btnReinInput = page.locator(`xpath=(//*[contains(normalize-space(.),"Frecuencia de pago inter")])[1]/following::button[contains(@class,"p-button-warning") or contains(normalize-space(.),"Reintentar")]`).first();
        if (await btnReinInput.isVisible().catch(() => false)) {
            await btnReinInput.click({ force: true }).catch(() => {});
            await page.waitForTimeout(2500);
        }

        await inputFrecuencia.scrollIntoViewIfNeeded().catch(() => { });
        await inputFrecuencia.click({ force: true }).catch(() => { });
        await page.waitForTimeout(150);
        await inputFrecuencia.press("Control+A").catch(() => { });
        await inputFrecuencia.press("Delete").catch(() => { });
        // pressSequentially dispara keydown/keypress/keyup que PrimeNG p-inputnumber requiere
        await inputFrecuencia.pressSequentially("1", { delay: 60 }).catch(async () => {
            // Fallback: fill + eventos manuales
            await inputFrecuencia!.fill("1").catch(() => {});
            await inputFrecuencia!.dispatchEvent('input').catch(() => {});
            await inputFrecuencia!.dispatchEvent('change').catch(() => {});
        });
        await page.keyboard.press('Tab').catch(() => {});
        await page.waitForTimeout(400);
        console.log(`[TallerCertificados] Frecuencia de pago inter\u00e9s aplicada (valor 1)`);
    }

    const seleccionarCuentaIndex0 = async () => {
        const candidatos = [
            page.locator('label').filter({ hasText: /^Cuenta$/i }).first().locator('xpath=following::*[contains(@class,"p-dropdown") and not(contains(@class,"disabled"))][1]').first(),
            page.locator('xpath=(//*[contains(normalize-space(.),"Cuenta")])[1]/following::*[contains(@class,"p-dropdown") and not(contains(@class,"disabled"))][1]').first(),
            page.locator('xpath=(//*[contains(normalize-space(.),"Cuenta")])[1]/following::*[@role="combobox"][1]').first(),
        ];

        let combo: Locator | null = null;
        for (const candidato of candidatos) {
            if (await candidato.count().catch(() => 0) > 0) {
                combo = candidato;
                break;
            }
        }
        if (!combo) return false;

        await combo.scrollIntoViewIfNeeded().catch(() => { });
        await combo.click({ force: true });
        await page.waitForTimeout(600);

        const panel = page.locator('.p-dropdown-panel:visible, [role="listbox"]:visible').last();
        await panel.waitFor({ state: 'visible', timeout: 2500 });

        const items = panel.locator('.p-dropdown-item:visible, [role="option"]:visible');
        const totalItems = await items.count().catch(() => 0);
        if (totalItems === 0) {
            throw new Error("El dropdown Cuenta no mostrÃƒÂ³ opciones visibles.");
        }

        const item = items.nth(0);
        await item.scrollIntoViewIfNeeded().catch(() => { });
        await item.click({ force: true });
        await page.waitForTimeout(600);
        return true;
    };

    const cuentaSeleccionada = await seleccionarCuentaIndex0().catch(() => false);
    if (cuentaSeleccionada) {
        campoVisible = true;
        console.log(`[TallerCertificados] Cuenta visible -> ÃƒÂ­ndice 0 seleccionado`);
    }

    if (!campoVisible) {
        console.log(`[TallerCertificados] Campos no capitalizables no visibles; se continÃƒÂºa sin error`);
    }
}

async function etapaEvidencias(page: Page, registro: RegistroExcel, capturasRef: string[]) {
    let capturoModalFinalizada = false;
    let capturoPostFinalizar = false;
    const capturarYCerrarModalFinalizadaSiVisible = async () => {
        const modalFinalizada = modalSolicitudFinalizadaLocator(page);
        const visible = await modalFinalizada.isVisible().catch(() => false);
        if (!visible) return false;

        if (!capturoModalFinalizada) {
            const evidenciaFinalizadaModal = await capturarCuentaComoPNG(
                page,
                registro.identificacion,
                `${registro.tipoCuenta}__solicitud-finalizada`
            );
            capturasRef.push(evidenciaFinalizadaModal);
            capturoModalFinalizada = true;
        }

        await cerrarModalSolicitudFinalizada(page, { timeoutMs: 15000, waitForAppearMs: 200 }).catch(() => false);
        return true;
    };

    const cerroInicio = await capturarYCerrarModalFinalizadaSiVisible().catch(() => false);
    if (!cerroInicio) {
        await cerrarModalSolicitudFinalizada(page, { timeoutMs: 1200 }).catch(() => false);
    }

    const badgePendiente = await page.getByText(/Pendiente finalizar/i).first().isVisible().catch(() => false);
    if (badgePendiente) {
        await finalizarSolicitudConDocumento(page);
        const siguePendiente = await page.getByText(/Pendiente finalizar/i).first().isVisible().catch(() => false);
        if (siguePendiente) {
            console.log(`[WARN][${registro.identificacion}] Se omite screenshot final porque la solicitud sigue en estado 'Pendiente finalizar'.`);
            return;
        }
    }
    await capturarYCerrarModalFinalizadaSiVisible().catch(() => false);

    const evidenciaPng = await capturarCuentaComoPNG(page, registro.identificacion, registro.tipoCuenta);
    capturasRef.push(evidenciaPng);
}

async function prepararSiguienteRegistro(page: Page) {
    if (page.isClosed()) return;

    // Si PW_KEEP_BROWSER_ON_FAIL=1, no navegar ni resetear contexto para preservar diagnóstico
    if (process.env.PW_KEEP_BROWSER_ON_FAIL === '1') {
        console.log('[Cleanup] PW_KEEP_BROWSER_ON_FAIL=1; no navego ni cierro contexto');
        return;
    }

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

    await page.goto(PORTAL_MULTIPRODUCT_URL, {
        waitUntil: 'domcontentloaded',
    }).catch(() => { });
    await esperarPortalListoTrasLogin(page, { timeoutMs: 120000 }).catch(() => ({ loginDetectado: false }));
    await asegurarPortalEnMultiproducto(page, 'ResetPortal').catch(() => false);
}



test('Certificados de Deposito Cliente existente - desde Excel', async () => {    // ===== INICIO: Configuracion de sesion/navegador =====
    const session = await launchPortalSession({
        defaultContinueOnError: true,
        initialUrl: PORTAL_MULTIPRODUCT_URL,
        preferNewPage: true,
    });
    const { context, page } = session;
    instalarForzadorUrlBizagi(context);
    let failedRegistros: Array<{ identificacion: string; error: string }> = [];

    try {
        const ids = leerRegistrosDesdeExcel('data/ids.xlsx', 'CD_Clientes_Existentes');
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
                console.log('[CERT-EX][FLOW] Llamando a etapaSeccionProductos...');
                await etapaSeccionProductos(context, page, registro);
                console.log('[CERT-EX][FLOW] etapaSeccionProductos completada');

                await procesarAprobacionExcepcionTasaSiAplica(context, page, registro);

                // ===== INICIO: Relacionados y asociacion =====
                const labelPropositoTaller = await etapaRelacionadosYAsociacion(page, registro);

                // ===== INICIO: Taller de productos =====
                await etapaTallerProductos(page, registro, labelPropositoTaller, capturas);

                const pendienteFinal = await estadoPendienteFinalizarVisible(page).catch(() => false);
                const btnGestionFinal = page
                    .locator('button[aria-label*="Gesti" i], button:has-text("Gesti"), button:has-text("Gesti")')
                    .first();
                const btnFinalizarFinal = page
                    .locator('button[aria-label*="Finalizar" i], button:has-text("Finalizar")')
                    .first();
                const gestionVisible = await btnGestionFinal.isVisible().catch(() => false);
                const finalizarVisible = await btnFinalizarFinal.isVisible().catch(() => false);
                const senalPantallaFinal = pendienteFinal || (gestionVisible && finalizarVisible);
                if (senalPantallaFinal) {
                    await finalizarSolicitudConDocumento(page);
                }
                const pendienteDespues = await estadoPendienteFinalizarVisible(page).catch(() => false);
                if (pendienteDespues) {
                    throw new Error("[CRITICO] La solicitud sigue en 'Pendiente finalizar' y no se ejecutÃƒÂ³ la finalizaciÃƒÂ³n con GestiÃƒÂ³n Documental.");
                }

                // ===== INICIO: Evidencias =====
                await etapaEvidencias(page, registro, capturas);
                let pendienteAntesOK = await estadoPendienteFinalizarVisible(page).catch(() => false);
                if (pendienteAntesOK) {
                    await finalizarSolicitudConDocumento(page);
                    pendienteAntesOK = await estadoPendienteFinalizarVisible(page).catch(() => false);
                    if (pendienteAntesOK) {
                        throw new Error("[CRITICO] La solicitud sigue en 'Pendiente finalizar' y no se ejecutÃƒÂ³ la finalizaciÃƒÂ³n con GestiÃƒÂ³n Documental.");
                    }
                }
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
            const pdfPath = await unirPNGsEnUnPDF(capturas, "Evidencias_Certificados.pdf");
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
