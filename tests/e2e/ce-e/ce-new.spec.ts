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
    clickBotonPorLabel,
} from '../../helpers/uiHelpers';
import {
    leerRegistrosDesdeExcel,
    marcarCedulasProcesadasEnExcel,
    cancelarCasoEnBizagiDesdePortal,
    abrirSolicitudCumplimientoEnBizagiDesdePortal,
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
    seleccionarProductoCuentaEfectivoNuevo,
    type RegistroExcel,
} from '../../helpers/productos/cuenta-efectivo/nuevo/helpers';
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
const cumplimientoMpnsProcesados = new Set<string>();
const plaftsSinManejadorLogueadas = new Set<string>();
const productosConfirmadosPorRegistro = new Set<string>();
const PORTAL_BASE_URL = (process.env.PW_PORTAL_URL ?? process.env.PORTAL_URL ?? 'https://srvqacgowb01.local.bsc.com:5000').replace(/\/+$/, '');
const PORTAL_MULTIPRODUCT_URL = `${PORTAL_BASE_URL}/requests/create/multiproduct`;

const getBotonContinuar = (page: Page) =>
    page.locator('button:visible').filter({ hasText: /^Continuar\b/i }).last();

const getBotonContinuarPrincipal = async (page: Page) => {
    const candidatosRol = page.getByRole('button', { name: /^Continuar\b/i });
    const candidatosTexto = page.locator('button:visible').filter({ hasText: /^Continuar\b/i });
    const colecciones = [candidatosRol, candidatosTexto];

    for (const coleccion of colecciones) {
        const total = await coleccion.count().catch(() => 0);
        if (!total) continue;

        for (let i = total - 1; i >= 0; i--) {
            const btn = coleccion.nth(i);
            const visible = await btn.isVisible().catch(() => false);
            if (!visible) continue;

            const enDialogo = await btn
                .evaluate((el) => !!el.closest('.p-dialog, [role="dialog"]'))
                .catch(() => false);

            if (!enDialogo) {
                return btn;
            }
        }
    }

    return null;
};

async function clickContinuarDesdeVerificacionAprobada(
    page: Page,
    contexto: string,
    timeoutMs = FAST_UI ? 12000 : 22000
) {
    const inicio = Date.now();

    while (Date.now() - inicio < timeoutMs) {
        const modalSalidaAntesClick = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        if (modalSalidaAntesClick) {
            throw new Error(`[CRITICO] Se detecto el modal 'Salir del proceso de solicitud' antes de Continuar (${contexto}). Se cancelo el modal para evitar salir por error.`);
        }
        if (await estaEnPantallaProductos(page).catch(() => false) && await modalConfiguracionProductoVisible(page).catch(() => false)) {
            throw new Error('[Producto][CRITICO] Se intentó Continuar con modal de configuración abierto');
        }

        const candidatos = page.getByRole('button', { name: /^Continuar$/i });
        const total = await candidatos.count().catch(() => 0);
        let mejorIndice = -1;
        let mejorX = -1;

        for (let i = 0; i < total; i++) {
            const btn = candidatos.nth(i);
            const visible = await btn.isVisible().catch(() => false);
            const enabled = visible ? await btn.isEnabled().catch(() => false) : false;
            if (!visible || !enabled) continue;

            const enDialogo = await btn
                .evaluate((el) => !!el.closest('.p-dialog, [role="dialog"]'))
                .catch(() => false);
            if (enDialogo) continue;

            const box = await btn.boundingBox().catch(() => null);
            const x = box?.x ?? -1;
            if (x > mejorX) {
                mejorX = x;
                mejorIndice = i;
            }
        }

        if (mejorIndice < 0) {
            await page.waitForTimeout(FAST_UI ? 180 : 450);
            continue;
        }

        const btnContinuar = candidatos.nth(mejorIndice);
        await btnContinuar.scrollIntoViewIfNeeded().catch(() => { });
        await page.waitForTimeout(FAST_UI ? 150 : 300);

        let clicked = await btnContinuar.click({ timeout: 2500, noWaitAfter: true }).then(() => true).catch(() => false);
        if (!clicked) {
            clicked = await btnContinuar.click({ force: true, timeout: 2500, noWaitAfter: true }).then(() => true).catch(() => false);
        }
        if (!clicked) {
            await page.waitForTimeout(FAST_UI ? 180 : 450);
            continue;
        }

        await page.waitForTimeout(FAST_UI ? 220 : 500);
        const modalSalidaAbierto = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        if (modalSalidaAbierto) {
            throw new Error(`[CRITICO] Se abrio el modal 'Salir del proceso de solicitud' al intentar Continuar (${contexto}). Se cancelo el modal para evitar salir por error.`);
        }

        console.log(`[Cumplimiento] Click en Continuar (${contexto}).`);
        return true;
    }

    return false;
}

async function esperarPortalEstableDespuesOfac(page: Page, timeoutMs = FAST_UI ? 16000 : 26000) {
    const inicio = Date.now();
    const textosGuardado = page.getByText(/Guardando informaci[oó]n|Informaci[oó]n guardada|Actualizando solicitud/i).first();
    const overlays = page.locator(
        '.p-blockui:visible, [data-pc-name="blockui"]:visible, .p-progressspinner:visible, .p-progress-spinner:visible'
    );

    while (Date.now() - inicio < timeoutMs) {
        const modalSalidaAbierto = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        if (modalSalidaAbierto) {
            throw new Error("[CRITICO] Se detecto el modal 'Salir del proceso de solicitud' durante la estabilizacion post-OFAC.");
        }

        const guardandoVisible = await textosGuardado.isVisible().catch(() => false);
        const overlayCount = await overlays.count().catch(() => 0);
        if (!guardandoVisible && overlayCount === 0) return true;

        await page.waitForTimeout(FAST_UI ? 250 : 500);
    }

    return false;
}

async function esperarFinGuardandoSolicitudRapido(page: Page, timeoutMs = FAST_UI ? 8000 : 10000) {
    const inicio = Date.now();
    const textosGuardado = page.getByText(/Guardando solicitud|Guardando informaci[oó]n|Actualizando solicitud/i).first();
    const overlays = page.locator(
        '.p-blockui:visible, [data-pc-name="blockui"]:visible, .p-progressspinner:visible, .p-progress-spinner:visible, .p-component-overlay-enter:visible'
    );

    console.log('[Verificaciones] Esperando fin de Guardando solicitud');
    while (Date.now() - inicio < timeoutMs) {
        const guardandoVisible = await textosGuardado.isVisible().catch(() => false);
        const overlayCount = await overlays.count().catch(() => 0);
        if (!guardandoVisible && overlayCount === 0) return true;
        await page.waitForTimeout(FAST_UI ? 200 : 300);
    }
    return false;
}

async function leerAprobacionesVerificacionesVisibles(page: Page) {
    const ofacVisible = await page.getByText(/Listas\s+OFAC|\bOFAC\b/i).first().isVisible().catch(() => false);
    const plaftVisible = await page.getByText(/\bPLAFT\b|Prevenci(?:o|ó)n.*Lavado|Lavado de Activos/i).first().isVisible().catch(() => false);
    const bloqueVerificaciones = page
        .locator('section, fieldset, .p-panel, .p-card, div')
        .filter({ hasText: /Verificaciones/i })
        .filter({ hasText: /Aprobado|OFAC|PLAFT/i })
        .first();
    const scopeAprobados = await bloqueVerificaciones.isVisible().catch(() => false) ? bloqueVerificaciones : page.locator('body');
    const tagValuesAprobado = scopeAprobados.locator('span.p-tag-value:visible').filter({ hasText: /^\s*Aprobado\s*$/i });
    const tagsSuccessAprobado = scopeAprobados.locator('.p-tag-success:visible, .p-tag:visible').filter({ hasText: /^\s*Aprobado\s*$/i });
    const tagValueCount = await tagValuesAprobado.count().catch(() => 0);
    const tagSuccessCount = await tagsSuccessAprobado.count().catch(() => 0);
    const tagsAprobadoVisibles = tagValueCount || tagSuccessCount;

    const aprobadoEnContenedor = async (texto: RegExp) => {
        const contenedor = scopeAprobados
            .locator('tr, li, .p-panel, .p-card, .p-datatable-row, div[class*="row"], div[class*="flex"], div')
            .filter({ hasText: texto })
            .filter({ hasText: /^.*Aprobado.*$/i })
            .first();
        const tagValue = contenedor.locator('span.p-tag-value').filter({ hasText: /^Aprobado$/i }).first();
        const tagSuccess = contenedor.locator('.p-tag-success, .p-tag').filter({ hasText: /^\s*Aprobado\s*$/i }).first();
        return await tagValue.isVisible().catch(() => false) || await tagSuccess.isVisible().catch(() => false);
    };

    const ofacAprobadoPorContenedor = ofacVisible ? await aprobadoEnContenedor(/Listas\s+OFAC|\bOFAC\b/i) : false;
    const plaftAprobado = plaftVisible ? await aprobadoEnContenedor(/\bPLAFT\b|Prevenci(?:o|ó)n.*Lavado|Lavado de Activos/i) : false;
    const ofacAprobado = ofacAprobadoPorContenedor || (ofacVisible && tagsAprobadoVisibles > 0) || tagsAprobadoVisibles === 1;

    return { ofacVisible, ofacAprobado, plaftVisible, plaftAprobado, tagsAprobadoVisibles };
}

async function esperarAprobacionesVerificacionesVisibles(
    page: Page,
    requeridas: { ofac?: boolean; plaft?: boolean } = { ofac: true, plaft: true },
    timeoutMs = FAST_UI ? 8000 : 10000
) {
    const inicio = Date.now();
    console.log(`[Verificaciones] Esperando aprobaciones visibles requeridas: ${requeridas.ofac ? 'OFAC' : ''}${requeridas.ofac && requeridas.plaft ? ', ' : ''}${requeridas.plaft ? 'PLAFT' : ''}`);
    console.log(`[Verificaciones] Requeridas para aprobar: OFAC=${!!requeridas.ofac} PLAFT=${!!requeridas.plaft}`);
    let ultimoLog = 0;
    let ultimoEstado = '';

    while (Date.now() - inicio < timeoutMs) {
        const estado = await leerAprobacionesVerificacionesVisibles(page);
        const estadoLog = `${estado.tagsAprobadoVisibles}|${estado.ofacVisible}|${estado.ofacAprobado}|${estado.plaftVisible}|${estado.plaftAprobado}`;
        if (estadoLog !== ultimoEstado || Date.now() - ultimoLog >= 1000) {
            console.log(`[Verificaciones] Tags Aprobado visibles=${estado.tagsAprobadoVisibles}`);
            console.log(`[Verificaciones] OFAC requerido aprobado=${requeridas.ofac ? estado.ofacAprobado : 'omitido'}`);
            console.log(`[Verificaciones] PLAFT ${requeridas.plaft ? `requerido aprobado=${estado.plaftAprobado}` : 'omitido'}`);
            ultimoEstado = estadoLog;
            ultimoLog = Date.now();
        }

        const hayRequeridas = !!requeridas.ofac || !!requeridas.plaft;
        const todasAprobadas = hayRequeridas
            && (!requeridas.ofac || estado.ofacAprobado)
            && (!requeridas.plaft || estado.plaftAprobado);
        if (todasAprobadas) {
            console.log('[Verificaciones] Todas las verificaciones requeridas están Aprobadas');
            return true;
        }

        if (!hayRequeridas) return false;
        await page.waitForTimeout(FAST_UI ? 200 : 300);
    }

    return false;
}

async function clickContinuarTrasAprobadoRapido(page: Page) {
    await esperarFinGuardandoSolicitudRapido(page).catch(() => false);
    const inicio = Date.now();
    while (Date.now() - inicio < (FAST_UI ? 8000 : 10000)) {
        if (await estaEnPantallaProductos(page).catch(() => false) && await modalConfiguracionProductoVisible(page).catch(() => false)) {
            throw new Error('[Producto][CRITICO] Se intentó Continuar con modal de configuración abierto');
        }
        const btnContinuar = await getBotonContinuarPrincipal(page);
        const visible = btnContinuar ? await btnContinuar.isVisible().catch(() => false) : false;
        const enabled = visible && btnContinuar ? await btnContinuar.isEnabled().catch(() => false) : false;
        console.log(`[Verificaciones] Botón Continuar enabled=${enabled}`);
        if (btnContinuar && visible && enabled) {
            await btnContinuar.scrollIntoViewIfNeeded().catch(() => { });
            const clicked = await btnContinuar.click({ timeout: 1800, noWaitAfter: true }).then(() => true).catch(() => false)
                || await btnContinuar.click({ force: true, timeout: 1800, noWaitAfter: true }).then(() => true).catch(() => false);
            if (clicked) {
                console.log('[Verificaciones] Click inmediato en Continuar tras Aprobado');
                return true;
            }
        }
        await page.waitForTimeout(FAST_UI ? 200 : 300);
    }
    return false;
}

async function detectarPantallaPostProductoReal(page: Page) {
    const body = (await page.locator('body').innerText().catch(() => '')) || '';
    if (/Nivel de estudio|Reside actualmente en la Rep(?:u|ú)blica Dominicana|Posee otras identificaciones o nacionalidades|A(?:ñ|n)adir dirección|Direccion|Dirección|Referencia personal/i.test(body)) {
        return true;
    }

    const marcadores: Locator[] = [
        page.getByText(/Nivel de estudio/i).first(),
        page.getByText(/Reside actualmente en la Rep(?:u|ú)blica Dominicana/i).first(),
        page.getByText(/Posee otras identificaciones o nacionalidades/i).first(),
        page.getByRole('button', { name: /A(?:ñ|n)adir direcci(?:o|ó)n/i }).first(),
        page.getByText(/Referencia personal/i).first(),
    ];

    for (const marcador of marcadores) {
        if (await marcador.isVisible().catch(() => false)) return true;
        if (await marcador.waitFor({ state: 'attached', timeout: 250 }).then(() => true).catch(() => false)) return true;
    }

    return false;
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

    const exito = await enMultiproducto();
    console.log(`[${contexto}] URL tras forzado directo: ${page.url()} | exito=${exito}`);
    return exito;
}

async function esperarPendientesGestionDocumentalCero(page: Page, timeoutMs = FAST_UI ? 8000 : 15000) {
    const inicio = Date.now();
    let loops = 0;
    while (Date.now() - inicio < timeoutMs) {
        loops++;
        console.log(`[GestionDoc] esperarPendientes... loop ${loops}: leyendo pendientes...`);
        const pendientes = await leerPendientesGestionDocumentalEnBoton(page).catch(() => 0);
        
        console.log(`[GestionDoc] esperarPendientes... loop ${loops}: leyendo modalGestion...`);
        const modalGestion = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
            .first();
        const modalVisible = await modalGestion.isVisible().catch(() => false);
        
        console.log(`[GestionDoc] esperarPendientes... loop ${loops}: enviandoVisible...`);
        const enviandoVisible = modalVisible
            ? await modalGestion.getByText(/Enviando documento/i).first().isVisible().catch(() => false)
            : false;

        console.log(`[GestionDoc] esperarPendientes... loop ${loops}: pendientes=${pendientes}, modalVisible=${modalVisible}, enviando=${enviandoVisible}`);
        if ((pendientes <= 0 && !enviandoVisible) || !modalVisible) {
            console.log(`[GestionDoc] esperarPendientes... completado exitosamente.`);
            return true;
        }

        // Wait, NO cerrarPopupEdgeSyncSiVisible HERE if it's the one hanging.
        // It's possibly hanging, let's skip it to be safe, or just do a hard wait.
        console.log(`[GestionDoc] esperarPendientes... loop ${loops}: esperando 800ms...`);
        await page.waitForTimeout(FAST_UI ? 350 : 800);
    }
    console.log(`[GestionDoc] esperarPendientes... timeout agotado (${timeoutMs}ms).`);
    return false;
}
const getSeccionContactos = (page: Page) =>
    page
        .locator(
            'xpath=(//*[@role="group"][.//*[contains(normalize-space(.),"Contactos")]] | //fieldset[.//*[contains(normalize-space(.),"Contactos")]])[1]'
        )
        .first();
const getBotonAnadirContacto = (page: Page) =>
    getSeccionContactos(page)
        .locator('button:visible')
        .filter({ hasText: /^A(?:\u00f1|n)adir contacto$/i })
        .first();
const getBotonAnadirLaborales = (page: Page) =>
    page.locator('button:has-text("Añadir datos laborales"), button:has-text("Anadir datos laborales")').first();

async function clickContinuarRobusto(
    page: Page,
    options?: { timeoutMs?: number; contexto?: string; postWaitMs?: number }
) {
    const timeoutMs = options?.timeoutMs ?? (FAST_UI ? 18000 : 32000);
    const postWaitMs = options?.postWaitMs ?? (FAST_UI ? 300 : 800);
    const contexto = options?.contexto ?? 'flujo';
    const inicio = Date.now();

    while (Date.now() - inicio < timeoutMs) {
        await cerrarPopupEdgeSyncSiVisible(page).catch(() => false);
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 4000 : 7000).catch(() => false);
        if (await estaEnPantallaProductos(page).catch(() => false) && await modalConfiguracionProductoVisible(page).catch(() => false)) {
            throw new Error('[Producto][CRITICO] Se intentó Continuar con modal de configuración abierto');
        }

        const btnContinuar = await getBotonContinuarPrincipal(page);
        if (!btnContinuar) {
            await page.waitForTimeout(FAST_UI ? 180 : 450);
            continue;
        }

        const visible = await btnContinuar.isVisible().catch(() => false);
        const enabled = visible ? await btnContinuar.isEnabled().catch(() => false) : false;
        if (!visible || !enabled) {
            await page.waitForTimeout(FAST_UI ? 180 : 450);
            continue;
        }

        await btnContinuar.scrollIntoViewIfNeeded().catch(() => { });
        await page.waitForTimeout(FAST_UI ? 120 : 250);

        const ariaDisabled = ((await btnContinuar.getAttribute('aria-disabled').catch(() => '')) || '').toLowerCase();
        if (ariaDisabled === 'true') {
            await page.waitForTimeout(FAST_UI ? 180 : 450);
            continue;
        }

        let clicked = false;

        clicked = await btnContinuar.click({ timeout: 2500, noWaitAfter: true })
            .then(() => true)
            .catch(() => false);

        if (!clicked) {
            clicked = await btnContinuar.click({ force: true, timeout: 2500, noWaitAfter: true })
                .then(() => true)
                .catch(() => false);
        }

        if (!clicked) {
            const box = await btnContinuar.boundingBox().catch(() => null);
            if (box) {
                await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
                clicked = true;
            }
        }

        if (!clicked) {
            const handle = await btnContinuar.elementHandle().catch(() => null);
            if (handle) {
                clicked = await page.evaluate((el) => {
                    const h = el as HTMLButtonElement;
                    h.focus();
                    h.click();
                    return true;
                }, handle).catch(() => false);
            }
        }

        if (!clicked) {
            await page.waitForTimeout(FAST_UI ? 180 : 450);
            continue;
        }

        console.log(`[GestionDoc] Click en Continuar (${contexto}).`);

        await page.waitForTimeout(postWaitMs).catch(() => { });
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 12000 : 18000).catch(() => false);

        return true;
    }

    console.log(`[GestionDoc][WARN] No fue posible hacer click en Continuar (${contexto}) dentro del timeout.`);
    return false;
}

type ResultadoAvanceRelacionados = 'taller' | 'productos' | 'gestion-documental' | 'post-producto' | 'sin-cambio';

async function confirmarCumplimientoAprobadoYContinuar(page: Page, mpn?: string) {
    const timeoutMs = 90000;
    const inicio = Date.now();
    const estadoAprobado = page.getByText(/^Aprobado$/i).first();
    const btnRefrescarCumplimiento = page
        .locator('.p-panel .p-panel-icons button:visible, [data-pc-name="panel"] [data-pc-section="icons"] button:visible')
        .first();

    await page.bringToFront().catch(() => { });
    console.log('[Cumplimiento] Portal al frente después de OFAC');

    const aprobadoRapido = await esperarAprobacionesVerificacionesVisibles(
        page,
        { ofac: true, plaft: false },
        FAST_UI ? 8000 : 10000
    ).catch(() => false);
    if (aprobadoRapido) {
        const clicRapido = await clickContinuarTrasAprobadoRapido(page).catch(() => false);
        if (!clicRapido) {
            throw new Error("[CRITICO] No se pudo hacer click en 'Continuar' despues de aprobar verificaciones.");
        }
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 8000 : 12000).catch(() => false);
        await esperarPortalEstableDespuesOfac(page, FAST_UI ? 8000 : 12000).catch((e) => { throw e; });

        const pantallaPostProductoReal = await detectarPantallaPostProductoReal(page).catch(() => false);
        console.log(`[Cumplimiento] Pantalla post-producto detectada=${pantallaPostProductoReal}`);
        if (pantallaPostProductoReal) {
            await asegurarPantallaPostProductoAntesDeTaller(page).catch(() => false);
            console.log('[Cumplimiento] Post-OFAC dejó el portal en pantalla post-producto');
            return 'post-producto' as const;
        }
    } else {
        console.log('[Cumplimiento] Recargando solicitud después de OFAC porque Aprobado no apareció rápido');
        await page.reload({ waitUntil: 'domcontentloaded', timeout: FAST_UI ? 20000 : 40000 }).catch(() => { });
        console.log('[Cumplimiento] Esperando portal estable después de reload');
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 12000 : 18000).catch(() => false);
        await esperarPortalEstableDespuesOfac(page).catch((e) => { throw e; });
    }

    const modalSalidaTrasReloadInicial = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    if (modalSalidaTrasReloadInicial) {
        throw new Error("[CRITICO] Se detecto el modal 'Salir del proceso de solicitud' despues de recargar post-OFAC.");
    }

    while (Date.now() - inicio < timeoutMs) {
        const aprobadoVisible = await estadoAprobado.isVisible().catch(() => false);
        if (aprobadoVisible) {
            console.log(`[Cumplimiento] Estado Aprobado confirmado${mpn ? ` para ${mpn}` : ''}.`);

            const aprobadoTrasRecarga = await page
                .waitForFunction(
                    () => /Aprobado/i.test(document.body?.innerText || ''),
                    { timeout: FAST_UI ? 1800 : 3000 }
                )
                .then(() => true)
                .catch(() => false);

            if (!aprobadoTrasRecarga) {
                console.log(`[Cumplimiento][WARN] Tras recargar el portal no se confirmo nuevamente el estado Aprobado${mpn ? ` para ${mpn}` : ''}. Reintentando...`);
                await page.waitForTimeout(FAST_UI ? 800 : 1600);
                continue;
            }
            console.log('[Cumplimiento] Estado Aprobado confirmado');

            console.log('[Cumplimiento] Buscando botón Continuar correcto');
            const clicContinuar = await clickContinuarDesdeVerificacionAprobada(
                page,
                'post-OFAC',
                FAST_UI ? 5000 : 8000
            );
            if (!clicContinuar) {
                throw new Error("[CRITICO] No se pudo hacer click en 'Continuar' despues de aprobar Cumplimiento.");
            }
            console.log('[Cumplimiento] Click en Continuar post-OFAC');
            await esperarFinActualizandoSolicitud(page, FAST_UI ? 12000 : 18000).catch(() => false);
            await esperarPortalEstableDespuesOfac(page).catch((e) => { throw e; });

            const modalSalidaTrasContinuar = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
            if (modalSalidaTrasContinuar) {
                throw new Error("[CRITICO] Se abrio el modal 'Salir del proceso de solicitud' despues de Continuar post-OFAC. Se cancelo para evitar salir por error.");
            }

            console.log('[Cumplimiento] Esperando pantalla post-producto real');
            const pantallaPostProductoDetectada = await page.waitForFunction(
                () => /Nivel de estudio|Reside actualmente en la Rep(?:u|ú)blica Dominicana|Posee otras identificaciones o nacionalidades|A(?:ñ|n)adir dirección|Direccion|Dirección|Referencia personal/i.test(document.body?.innerText || ''),
                { timeout: FAST_UI ? 18000 : 30000 }
            ).then(() => true).catch(() => false);
            const pantallaPostProductoReal = pantallaPostProductoDetectada || await detectarPantallaPostProductoReal(page).catch(() => false);
            console.log(`[Cumplimiento] Pantalla post-producto detectada=${pantallaPostProductoReal}`);

            if (pantallaPostProductoReal) {
                await asegurarPantallaPostProductoAntesDeTaller(page).catch(() => false);
                console.log('[Cumplimiento] Post-OFAC dejó el portal en pantalla post-producto');
                return 'post-producto' as const;
            }

            console.log(`[Cumplimiento][WARN] Continuar ejecutado pero no se detecto pantalla post-producto real${mpn ? ` para ${mpn}` : ''}. Reintentando...`);
        }

        const refreshVisible = await btnRefrescarCumplimiento.isVisible().catch(() => false);
        if (refreshVisible) {
            await btnRefrescarCumplimiento.click({ force: true }).catch(() => { });
        }

        await page.waitForTimeout(FAST_UI ? 1200 : 2500);
    }

    throw new Error(`[CRITICO] No se confirmo el estado 'Aprobado' en portal luego de gestionar Cumplimiento${mpn ? ` para ${mpn}` : ''}.`);
}

async function procesarVerificacionesEspeciales(page: Page) {
    const badgeCumplimientoVisible = await page
        .getByText(/Verificaciones\s*-\s*Cumplimiento/i)
        .first()
        .isVisible()
        .catch(() => false);
    const plaftVisible = await page
        .getByText(/PLAFT|Prevenci(?:o|\u00f3)n.*Lavado|Lavado de Activos/i)
        .first()
        .isVisible()
        .catch(() => false);
    const verificacionesDetectadas = badgeCumplimientoVisible || plaftVisible;

    console.log(`[Verificaciones] Pantalla/badge de verificaciones detectado=${verificacionesDetectadas}`);
    console.log(`[Verificaciones] OFAC detectado=${badgeCumplimientoVisible}`);
    console.log(`[Verificaciones] PLAFT detectado=${plaftVisible}`);
    console.log(`[Cumplimiento] badgeVisible=${badgeCumplimientoVisible} url=${page.url()}`);

    if (plaftVisible && !badgeCumplimientoVisible) {
        const etapaKey = page.url().split('/').pop() || 'unknown';
        if (!plaftsSinManejadorLogueadas.has(etapaKey)) {
            plaftsSinManejadorLogueadas.add(etapaKey);
            console.log('[Verificaciones][WARN] PLAFT detectado pero no existe manejador implementado; se continuará flujo normal');
        }
    }

    if (!badgeCumplimientoVisible) {
        const verificacionBpm = await abrirBpmSiVerificacionConoceCliente(page).catch((e) => {
            console.log(`[Verificacion][WARN] ${e instanceof Error ? e.message : String(e)}`);
            return null;
        });
        if (verificacionBpm?.mpn) {
            console.log(`[Verificacion] Solicitud para BPM: ${verificacionBpm.mpn}`);
            return { tipo: 'conoce-cliente' as const, mpn: verificacionBpm.mpn };
        }

        console.log('[Verificaciones] No hay verificaciones pendientes, continuando flujo normal');
        return null;
    }

    console.log('[Verificaciones] Resolviendo OFAC');

    const mpnVisible = await page
        .locator('span.p-tag-value')
        .filter({ hasText: /^MPN-\d+$/i })
        .first()
        .textContent()
        .then((v) => (v || '').trim().toUpperCase())
        .catch(() => '');

    console.log(`[Cumplimiento] mpnVisible='${mpnVisible}'`);

    if (mpnVisible && cumplimientoMpnsProcesados.has(mpnVisible)) {
        const estado = await confirmarCumplimientoAprobadoYContinuar(page, mpnVisible);
        console.log('[Verificaciones] Todas las verificaciones condicionales aprobadas');
        return { tipo: 'cumplimiento' as const, mpn: mpnVisible, estado };
    }

    const verificacionCumplimiento = await abrirSolicitudCumplimientoEnBizagiDesdePortal(page, {
        password: BIZAGI_PASSWORD,
    }).catch((e) => {
        console.log(`[Cumplimiento][WARN] ${e instanceof Error ? e.message : String(e)}`);
        return null;
    });
    if (verificacionCumplimiento?.mpn) {
        cumplimientoMpnsProcesados.add(verificacionCumplimiento.mpn.toUpperCase());
        const estado = await confirmarCumplimientoAprobadoYContinuar(page, verificacionCumplimiento.mpn);
        console.log(`[Cumplimiento] Solicitud abierta en Bizagi para: ${verificacionCumplimiento.mpn}`);
        console.log('[Verificaciones] Todas las verificaciones condicionales aprobadas');
        return { tipo: 'cumplimiento' as const, mpn: verificacionCumplimiento.mpn, estado };
    }

    return null;
}

async function intentarAvanceRealHaciaTaller(
    page: Page,
    contexto: string,
    options?: { maxClicks?: number; tipoCuenta?: string }
): Promise<ResultadoAvanceRelacionados> {
    const maxClicks = options?.maxClicks ?? 2;

    for (let intento = 1; intento <= maxClicks; intento++) {
        if (options?.tipoCuenta && await estaEnPantallaProductos(page).catch(() => false)) {
            if (await modalConfiguracionProductoVisible(page).catch(() => false)) {
                throw new Error('[Producto][CRITICO] Se intentó Continuar con modal de configuración abierto');
            }
            const productoVisible = await productoAgregadoComoTarjetaVisible(page, options.tipoCuenta).catch(() => false);
            console.log(`[Producto] Producto agregado como tarjeta final visible=${productoVisible}`);
            if (!productoVisible) {
                throw new Error('[Producto][CRITICO] No se confirmo producto agregado; no se puede continuar.');
            }
            console.log('[Producto] Producto ya estaba agregado, no se volverá a agregar');
            console.log('[Producto] Evitando duplicar producto/cuenta');
            console.log('[Producto] Continuando con producto ya confirmado');
            console.log('[Producto] Producto confirmado, haciendo Continuar');
        }

        const verificacionEspecialAntes = await procesarVerificacionesEspeciales(page).catch(() => null);
        if (verificacionEspecialAntes?.tipo === 'cumplimiento') {
            if (verificacionEspecialAntes.estado === 'post-producto') {
                console.log(`[Continuar] Cumplimiento dejo pantalla post-producto (${contexto}).`);
                return 'post-producto';
            }
            console.log(`[Continuar] Cumplimiento ya gestionado antes del click (${contexto}). Paso a validacion de pantalla siguiente.`);
            return 'sin-cambio';
        }

        const clicOk = await clickContinuarRobusto(page, {
            timeoutMs: FAST_UI ? 14000 : 24000,
            contexto: `${contexto} - click ${intento}/${maxClicks}`,
            postWaitMs: FAST_UI ? 180 : 500,
        });

        if (!clicOk) {
            await page.waitForTimeout(FAST_UI ? 180 : 500);
            continue;
        }

        const urlAntesClic = page.url();
        console.log(`[Continuar] esperarFinActualizando post-click (${contexto})...`);
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 12000 : 18000).catch(() => false);
        const urlDespuesClic = page.url();
        console.log(`[Continuar] esperarFinActualizando completado. url=${urlDespuesClic}`);

        const verificacionEspecialDespuesClick = await procesarVerificacionesEspeciales(page).catch(() => null);
        if (verificacionEspecialDespuesClick?.tipo === 'cumplimiento') {
            if (verificacionEspecialDespuesClick.estado === 'post-producto') {
                console.log(`[Continuar] Cumplimiento dejo pantalla post-producto tras el click (${contexto}).`);
                return 'post-producto';
            }
            console.log(`[Continuar] Cumplimiento gestionado tras el click (${contexto}). Paso a validacion de pantalla siguiente.`);
            return 'sin-cambio';
        }

        // Si el URL no cambio, el portal rechazo con error de validacion (ej. Nivel de estudio vacio).
        // PERO: en portales SPA con wizard el URL NO cambia entre pasos. Solo intentar completar
        // campos si realmente seguimos en la pantalla pre-productos (hay "Gestión Documental" visible
        // u otros marcadores del formulario). Si ya estamos en la pantalla de productos, no tocar nada.
        if (urlDespuesClic === urlAntesClic) {
            const verificacionEspecialConUrlIgual = await procesarVerificacionesEspeciales(page).catch(() => null);
            if (verificacionEspecialConUrlIgual) {
                if (verificacionEspecialConUrlIgual.tipo === 'cumplimiento' && verificacionEspecialConUrlIgual.estado === 'post-producto') {
                    console.log(`[Continuar] Cumplimiento dejo pantalla post-producto con URL sin cambio (${contexto}).`);
                    return 'post-producto';
                }
                console.log(`[Continuar] Verificacion especial detectada (${verificacionEspecialConUrlIgual.tipo}) con URL sin cambio (${contexto}).`);
                return 'sin-cambio';
            }

            const enPantallaPrevia = await estaEnPantallaPreProductos(page).catch(() => false);
            if (enPantallaPrevia) {
                console.log(`[Continuar] URL no cambio y seguimos en pantalla previa. Intentando completar campos faltantes...`);
                const completoCampos = await completarInformacionClientePostProductoAntesDeTaller(page, { required: false });
                if (completoCampos) {
                    console.log(`[Continuar] Campos completados tras rechazo. Scrolleando a seccion productos y reintentando Continuar...`);
                    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight)).catch(() => { });
                    await page.waitForTimeout(FAST_UI ? 350 : 800);
                    continue;
                }
            } else {
                console.log(`[Continuar] URL no cambio pero ya NO estamos en pantalla previa (wizard SPA). Verificando destino...`);
            }
        }

        const llegoATaller = await esperarPantallaTallerProductos(page, {
            timeoutMs: FAST_UI ? 1000 : 2500,
        }).catch(() => false);

        if (llegoATaller) {
            console.log(`[Continuar] Avance real detectado hacia Taller (${contexto}).`);
            return 'taller';
        }

        const tituloProductos = page.getByText(/Productos/i).first();
        const tieneDropdowns = await page.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').count().catch(() => 0) >= 2;
        const enProductos = (await tituloProductos.isVisible().catch(() => false)) && tieneDropdowns;

        if (enProductos) {
            console.log(`[Continuar] Avance detectado hacia Productos (${contexto}).`);
            if (options?.tipoCuenta && productoConfirmadoEnRegistro({ tipoCuenta: options.tipoCuenta, identificacion: '' } as unknown as RegistroExcel)) {
                console.log(`[Producto] Producto ya confirmado para tipo=${options.tipoCuenta}, continuando...`);
                return 'productos';
            }
            console.log(`[Continuar] Productos detectado, retornando 'productos' para procesamiento posterior`);
            return 'productos';
        }

        const pendientes = await leerPendientesGestionDocumentalEnBoton(page).catch(() => 0);
        let requiereAdjunto = await requiereAdjuntarDocumentoIdentidad(page).catch(() => false);
        // Si el badge indica pendientes pero el modal aun no abrio, esperar hasta 2s y re-verificar.
        if (!requiereAdjunto && pendientes > 0) {
            await page.waitForTimeout(FAST_UI ? 1000 : 2000);
            requiereAdjunto = await requiereAdjuntarDocumentoIdentidad(page).catch(() => false);
        }
        console.log(`[Continuar] No llego a Taller. pendientes=${pendientes} requiereAdjunto=${requiereAdjunto} url=${page.url()}`);

        if (requiereAdjunto) {
            console.log(
                `[Continuar] Click abrio o mantuvo Gestion Documental (${contexto}). pendientes=${pendientes}`
            );
            return 'gestion-documental';
        }

        console.log(
            `[Continuar][WARN] Se hizo click en Continuar pero no hubo avance real ni requerimiento documental (${contexto}).`
        );

        await page.waitForTimeout(FAST_UI ? 180 : 500);
    }

    return 'sin-cambio';
}

const randomInt = (min: number, max: number) =>
    Math.floor(Math.random() * (max - min + 1)) + min;

const randomTelefono = () =>
    `809-${randomInt(200, 999)}-${randomInt(1000, 9999)}`;

const CORREO_FIJO_CLIENTE_NUEVO = 'calidad@hotmail.com';
const FAST_UI = (process.env.PW_FAST_UI ?? '1') === '1';
const CONTACT_MODAL_ATTEMPTS = FAST_UI ? 3 : 4;
const CONTACT_MODAL_BTN_TIMEOUT = FAST_UI ? 9000 : 15000;
const CONTACT_MODAL_OPEN_TIMEOUT_1 = FAST_UI ? 2200 : 4500;
const CONTACT_MODAL_OPEN_TIMEOUT_2 = FAST_UI ? 1400 : 2500;
const CONTACT_SHORT_WAIT_MS = FAST_UI ? 120 : 350;
const CONTACT_VALIDATE_ATTEMPTS = FAST_UI ? 14 : 30;
const CONTACT_VALIDATE_POLL_MS = FAST_UI ? 160 : 400;
const CONTACT_CORREO_RETRY_WAIT_MS = FAST_UI ? 180 : 500;
const LIST_PROBE_TIMEOUT = FAST_UI ? 700 : 2500;
const LIST_MAX_RETRIES = FAST_UI ? 3 : 6;
const LIST_RETRY_CLICK_TIMEOUT = FAST_UI ? 300 : 1200;
const LIST_RETRY_WAIT_MS = FAST_UI ? 120 : 600;
const LIST_LABEL_TIMEOUT = FAST_UI ? 3500 : 6000;
const LIST_PANEL_TIMEOUT = FAST_UI ? 4000 : 7000;
const LIST_SCOPE_RETRY_WAIT_MS = FAST_UI ? 350 : 600;
const LIST_PANEL_QUICK_TIMEOUT = FAST_UI ? 2200 : 4000;
const ADDRESS_MODAL_OPEN_BTN_TIMEOUT = FAST_UI ? 6000 : 12000;
const ADDRESS_MODAL_OPEN_TIMEOUT = FAST_UI ? 3200 : 6000;
const ADDRESS_MODAL_INPUT_TIMEOUT = FAST_UI ? 2800 : 7000;
const ADDRESS_MODAL_OPTIONAL_INPUT_TIMEOUT = FAST_UI ? 900 : 2500;
const ADDRESS_MODAL_LABEL_TIMEOUT = FAST_UI ? 1700 : 4200;
const ADDRESS_MODAL_PANEL_TIMEOUT = FAST_UI ? 2300 : 4500;
const ADDRESS_MODAL_MAX_RETRIES = FAST_UI ? 3 : 5;
const ADDRESS_MODAL_DEPENDENT_WAIT_MS = FAST_UI ? 40 : 180;
const ADDRESS_MODAL_VALUE_READ_TIMEOUT = FAST_UI ? 700 : 1800;
const PRODUCTO_MAX_INTENTOS = FAST_UI ? 6 : 8;
const PRODUCTO_WAIT_CONFIRM_MS = FAST_UI ? 1400 : 2500;
const PRODUCTO_RETRY_WAIT_MS = FAST_UI ? 600 : 1200;
const FINALIZACION_MAX_INTENTOS = FAST_UI ? 5 : 7;
const FINALIZACION_TIMEOUT_MS = FAST_UI ? 90000 : 120000;
const BIZAGI_PASSWORD = "H0la1234.";

const randomTexto = (prefijo: string) =>
    `${prefijo} ${randomInt(100, 9999)}`;

const resolverArchivoCedula = () => {
    const base = path.resolve(process.cwd(), 'data', 'Documentos');
    const candidatos = ['cedula.pdf', 'Cedula.pdf', 'copia-cedula.pdf', 'evidencia.pdf'];
    for (const nombre of candidatos) {
        const full = path.join(base, nombre);
        if (fs.existsSync(full)) return full;
    }
    return path.join(base, 'evidencia.pdf');
};

const asPattern = (labelText: string | RegExp) =>
    labelText instanceof RegExp
        ? labelText
        : new RegExp(String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

async function cerrarPopupEdgeSyncSiVisible(page: Page) {
    const popupSync = page
        .locator('[role="dialog"]:visible, .p-dialog:visible, div:visible')
        .filter({ hasText: /syncing your browsing data|Microsoft Edge|Got it|Customize/i })
        .first();
    const visible = await popupSync.isVisible().catch(() => false);
    if (!visible) return false;

    const btnGotIt = popupSync.getByRole('button', { name: /^Got it$/i }).first();
    const btnCerrar = popupSync.getByRole('button', { name: /^(Close|Cerrar)$/i }).first();
    if (await btnGotIt.isVisible().catch(() => false)) {
        await btnGotIt.click({ force: true }).catch(() => { });
    } else if (await btnCerrar.isVisible().catch(() => false)) {
        await btnCerrar.click({ force: true }).catch(() => { });
    } else {
        await page.keyboard.press('Escape').catch(() => { });
    }
    await page.waitForTimeout(200);
    return true;
}

async function requiereAdjuntarDocumentoIdentidad(page: Page) {
    const enPantallaProductos = await page
        .getByText(/Categoría de producto/i)
        .first()
        .isVisible()
        .catch(() => false);
    if (enPantallaProductos) {
        return false;
    }

    const modalGestionVisible = await page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
        .first()
        .isVisible()
        .catch(() => false);
    if (modalGestionVisible) return true;

    const alertas = [
        page.getByText(/Debe adjuntar|Debe cargar|Documento(?:s)? requerido(?:s)?|Documento(?:s)? pendiente(?:s)?|adjunte los documentos/i).first(),
        page.locator('.p-inline-message-error:visible, .p-message-error:visible, .p-toast-message-error:visible, .p-inline-message-warn:visible, .p-message-warn:visible, .p-toast-message-warn:visible, .p-toast-message:visible')
            .filter({ hasText: /adjuntar|adjunte|cargar|cargue|documento requerido|obligatorio|documento(?:s)? pendiente(?:s)?/i })
            .first(),
    ];
    for (const alerta of alertas) {
        if (await alerta.isVisible().catch(() => false)) return true;
    }

    return false;
}

async function cerrarModalGestionDocumentalSiVisible(page: Page) {
    const modal = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
        .first();

    const visible = await modal.isVisible().catch(() => false);
    if (!visible) return true;

    // Drain secondary "Cargar documento" modal before trying to close the primary one.
    // If the secondary modal is still open, the X button on the primary modal may be blocked.
    const dialogCargaDrain = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Cargar documento/i })
        .first();
    const cargaAbierta = await dialogCargaDrain.isVisible().catch(() => false);
    if (cargaAbierta) {
        console.log('[GestionDoc] Modal secundario "Cargar documento" aun abierto; esperando que cierre...');
        await dialogCargaDrain.waitFor({ state: 'hidden', timeout: 15000 }).catch(() => {
            console.log('[GestionDoc][WARN] Modal "Cargar documento" no cerro en 15s; continuando de todas formas.');
        });
    }

    const closeBtn = modal.locator(
        '[data-pc-section="closebutton"], [data-pc-group-section="headericon"], button.p-dialog-header-close, button[aria-label="Cerrar"], button[aria-label="Close"]'
    ).first();

    const txtEnviando = modal.getByText(/Enviando documento/i).first();
    const maxEsperaMs = FAST_UI ? 45000 : 70000;
    const inicio = Date.now();

    while (Date.now() - inicio < maxEsperaMs) {
        const sigueVisible = await modal.isVisible().catch(() => false);
        if (!sigueVisible) return true;

        const enviandoVisible = await txtEnviando.isVisible().catch(() => false);
        if (enviandoVisible) {
            await page.waitForTimeout(FAST_UI ? 400 : 900);
            continue;
        }

        const closeVisible = await closeBtn.isVisible().catch(() => false);
        if (closeVisible) {
            await closeBtn.scrollIntoViewIfNeeded().catch(() => { });

            let clicked = await closeBtn.click({ timeout: 2000 }).then(() => true).catch(() => false);
            if (!clicked) {
                clicked = await closeBtn.click({ force: true, timeout: 2000 }).then(() => true).catch(() => false);
            }

            if (!clicked) {
                const box = await closeBtn.boundingBox().catch(() => null);
                if (box) {
                    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
                    clicked = true;
                }
            }

            if (clicked) {
                const hidden = await modal.waitFor({ state: 'hidden', timeout: FAST_UI ? 5000 : 9000 })
                    .then(() => true)
                    .catch(() => false);
                if (hidden) {
                    await page.waitForTimeout(FAST_UI ? 250 : 700);
                    return true;
                }
            }
        }

        await page.keyboard.press('Escape').catch(() => { });
        const hiddenByEsc = await modal.waitFor({ state: 'hidden', timeout: FAST_UI ? 2500 : 4500 })
            .then(() => true)
            .catch(() => false);
        if (hiddenByEsc) {
            await page.waitForTimeout(FAST_UI ? 250 : 700);
            return true;
        }

        await page.waitForTimeout(FAST_UI ? 250 : 600);
    }

    console.log('[GestionDoc][WARN] El modal documental no cerro de forma limpia dentro del timeout.');
    return false;
}
async function continuarResolviendoGestionDocumentalSiPide(page: Page, options?: { maxIntentos?: number }) {
    const maxIntentos = options?.maxIntentos ?? 3;

    // Si ya estamos en la pantalla de productos, no hacer nada
    const yaEnProductosInicial = await page.getByText(/Categor[ií]a de producto/i).first().isVisible().catch(() => false);
    if (yaEnProductosInicial) {
        console.log('[GestionDoc] Ya en pantalla de productos al inicio. Omitiendo Continuar.');
        return true;
    }

    for (let intento = 1; intento <= maxIntentos; intento++) {
        const verificacionEspecialAntes = await procesarVerificacionesEspeciales(page).catch(() => null);
        if (verificacionEspecialAntes) {
            console.log(`[GestionDoc] Verificacion especial detectada (${verificacionEspecialAntes.tipo}) antes de reintentar Continuar.`);
            return true;
        }

        const clicContinuar = await clickContinuarRobusto(page, {
            timeoutMs: FAST_UI ? 12000 : 22000,
            contexto: `reintento ${intento}/${maxIntentos}`,
        });
        if (!clicContinuar) {
            const verificacionEspecialSinClick = await procesarVerificacionesEspeciales(page).catch(() => null);
            if (verificacionEspecialSinClick) {
                console.log(`[GestionDoc] Verificacion especial detectada (${verificacionEspecialSinClick.tipo}) al fallar el click en Continuar.`);
                return true;
            }
            if (intento < maxIntentos) {
                await page.waitForTimeout(FAST_UI ? 180 : 500);
                continue;
            }
            throw new Error("[CRITICO] Boton 'Continuar' no visible/habilitado para avanzar.");
        }

        // Si ya navego a la pantalla de productos, no hay nada mas que hacer
        const yaEnProductos = await page.getByText(/Categor[ií]a de producto/i).first()
            .waitFor({ state: 'visible', timeout: 10000 })
            .then(() => true)
            .catch(() => false);
        if (yaEnProductos) {
            console.log('[GestionDoc] Pantalla de productos detectada. No se requiere adjuntar documento.');
            return true;
        }

        await cerrarModalGestionDocumentalSiVisible(page).catch(() => false);
        const pendientesEnBadge = await leerPendientesGestionDocumentalEnBoton(page).catch(() => 0);
        let requiereAdjunto = await requiereAdjuntarDocumentoIdentidad(page).catch(() => false);

        // Verificar de nuevo si ya llegamos a productos (puede que haya cargado mientras revisabamos)
        const yaEnProductos2 = await page.getByText(/Categor[ií]a de producto/i).first().isVisible().catch(() => false);
        if (yaEnProductos2) {
            console.log('[GestionDoc] Pantalla de productos detectada (check2). No se requiere adjuntar documento.');
            return true;
        }

        // Si el badge indica pendientes pero el modal aun no abrio, esperar y re-verificar.
        if (!requiereAdjunto && pendientesEnBadge > 0) {
            await page.waitForTimeout(FAST_UI ? 1000 : 2000);
            requiereAdjunto = await requiereAdjuntarDocumentoIdentidad(page).catch(() => false);
        }

        // Si el badge sigue en > 0 y aun no hay señal explícita, abrir el modal proactivamente
        // para verificar si hay botones de upload pendientes (el toast puede haber desaparecido ya).
        // IMPORTANTE: el contenido del modal se renderiza via Angular CDK Portal FUERA de .p-dialog,
        // por lo que puede tardar varios segundos en aparecer tras abrir el modal (especialmente en
        // staging). Usar un polling loop en vez de un wait fijo.
        if (!requiereAdjunto && pendientesEnBadge > 0) {
            const btnGestionProbe = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
            const btnProbeVisible = await btnGestionProbe.isVisible().catch(() => false);
            if (btnProbeVisible) {
                console.log('[GestionDoc] Badge>0 sin modal/alerta. Abriendo GD modal para verificar uploads pendientes...');
                await btnGestionProbe.click({ force: true }).catch(() => { });

                // Esperar a que el modal este realmente visible (no solo que el click haya ocurrido)
                const modalProbe = page
                    .locator('.p-dialog:visible, [role="dialog"]:visible')
                    .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
                    .first();
                const modalAbrio = await modalProbe.waitFor({ state: 'visible', timeout: 6000 })
                    .then(() => true)
                    .catch(() => false);

                let uploadCount = 0;
                if (modalAbrio) {
                    // CDK Portal puede tardar varios segundos en renderizar el contenido.
                    // Polling hasta 8 segundos.
                    const portalDeadline = Date.now() + 8000;
                    while (Date.now() < portalDeadline) {
                        uploadCount = await page
                            .locator('button:visible:has(i.pi-upload), button:visible:has(.pi-upload), button:visible:has([class*="pi-upload"]), button:visible:has([class*="ph-upload"]), button:visible:has([class*="ph--upload"]), button:visible[aria-label*="cargar" i], button:visible[aria-label*="subir" i], button:visible[aria-label*="Cargar" i], button[aria-label*="adjuntar" i]:visible')
                            .count()
                            .catch(() => 0);
                        console.log(`[GestionDoc][Debug] probe: modalAbrio=true uploadCount=${uploadCount} (${portalDeadline - Date.now()}ms restantes)`);
                        if (uploadCount > 0) break;
                        await page.waitForTimeout(500);
                    }
                } else {
                    console.log('[GestionDoc][WARN] GD modal no abrio en 6s. Continuando sin verificar uploads.');
                }

                if (uploadCount > 0) {
                    console.log(`[GestionDoc] ${uploadCount} uploads pendientes confirmados. Cerrando modal y marcando requiereAdjunto=true.`);
                    // Cerrar limpiamente — el flujo de adjunto lo re-abrira
                    await cerrarModalGestionDocumentalSiVisible(page).catch(() => { });
                    await page.waitForTimeout(FAST_UI ? 300 : 600);
                    requiereAdjunto = true;
                } else {
                    console.log('[GestionDoc] Sin uploads pendientes en modal tras espera. Cerrando y continuando.');
                    await page.keyboard.press('Escape').catch(() => { });
                    await page.waitForTimeout(300);
                }
            }
        }

        // Si no hay requisito de documento real (modal visible o alerta), avanzar.
        // El badge puede quedar en 1 aunque el documento ya este cargado (contador de persona).
        if (!requiereAdjunto) return true;

        console.log(`[GestionDoc] Se detecto requerimiento de documento (cedula). Pendientes badge=${pendientesEnBadge}. Adjuntando y reintentando Continuar...`);
        let adjunto = await validarGestionDocumentalSiRequerido(page).catch((e) => {
            console.log(`[GestionDoc][WARN] validarGestionDocumentalSiRequerido fallo: ${String(e)}`);
            return false;
        });
        if (!adjunto) {
            const evidenciaPath = resolverArchivoCedula();
            const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
            const modalGestion = page
                .locator('.p-dialog:visible, [role="dialog"]:visible')
                .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
                .first();

            const modalYaVisible = await modalGestion.isVisible().catch(() => false);
            if (!modalYaVisible) {
                const btnGestionVisible = await btnGestion.isVisible().catch(() => false);
                if (btnGestionVisible) {
                    await btnGestion.scrollIntoViewIfNeeded().catch(() => { });
                    await btnGestion.click({ force: true }).catch(() => { });
                }
            }

            adjunto = await cargarDocumentoEnGestionDocumental(page, evidenciaPath)
                .then(() => true)
                .catch((e) => {
                    console.log(`[GestionDoc][WARN] carga directa en modal fallo: ${String(e)}`);
                    return false;
                });
        }
        if (adjunto) {
            const maxEsperaS = 30;
            for (let s = 1; s <= maxEsperaS; s++) {
                const modalGestionEnviando = page
                    .locator('.p-dialog:visible, [role="dialog"]:visible')
                    .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
                    .first();
                const enviandoVisible = await modalGestionEnviando.getByText(/Enviando documento/i).first().isVisible().catch(() => false);
                const modalVisible = await modalGestionEnviando.isVisible().catch(() => false);
                
                if (!enviandoVisible && !modalVisible) {
                    console.log(`[GestionDoc] Modal y texto desaparecieron en el seg ${s}.`);
                    break;
                } else if (!enviandoVisible && modalVisible) {
                    console.log(`[GestionDoc] Texto "Enviando" desaparecio en el seg ${s}, pero modal sigue visible.`);
                    break;
                }
                
                if (s % 5 === 0) console.log(`[GestionDoc] Esperando cierre de subida... seg ${s}/${maxEsperaS}. modal=${modalVisible} enviando=${enviandoVisible}`);
                await page.waitForTimeout(1000);
            }

            console.log('[GestionDoc] Procediendo a forzar cierre del modal...');
            const modalCerrado = await cerrarModalGestionDocumentalSiVisible(page).catch(() => false);
            if (!modalCerrado) {
                console.log("[GestionDoc][WARN] El modal de Gestion documental no pudo forzarse a cerrar luego de adjuntar.");
            } else {
                console.log("[GestionDoc] El modal de Gestion documental se ha cerrado.");
            }

            console.log('[GestionDoc] Esperando 1200ms tras cerrado de modal...');
            await page.waitForTimeout(FAST_UI ? 500 : 1200);

            console.log('[GestionDoc] Esperando Fin Actualizando Solicitud...');
            await esperarFinActualizandoSolicitud(page, FAST_UI ? 12000 : 18000).catch(() => false);

            console.log('[GestionDoc] Checkeando si modalGestionSigueVisible...');
            const modalGestionSigueVisible = await page
                .locator('.p-dialog:visible, [role="dialog"]:visible')
                .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
                .first()
                .isVisible()
                .catch(() => false);
            if (modalGestionSigueVisible) {
                console.log('[GestionDoc] modalGestionSigueVisible=true, reintentando cerrar...');
                await cerrarModalGestionDocumentalSiVisible(page).catch(() => false);
            }

            console.log('[GestionDoc] Esperando a que el badge de subidas pendientes este en cero...');
            const badgeCero = await esperarPendientesGestionDocumentalCero(page).catch(() => false);

            console.log('[GestionDoc] Leyendo pendientes del badge POST adjunto...');
            const pendientesPostAdjunto = await leerPendientesGestionDocumentalEnBoton(page).catch(() => -1);
            console.log(`[GestionDoc] Estado post-adjunto: badge=${pendientesPostAdjunto}, listo=${badgeCero}.`);

            // Drain any lingering dialog overlays before clicking Continuar.
            // A p-dialog-mask still in the DOM (even hidden) can intercept pointer events.
            console.log('[GestionDoc] Drenando overlays antes de Continuar...');
            await page.waitForFunction(
                () => !document.querySelector('.p-dialog-mask:not([style*="display: none"]), .p-component-overlay-enter'),
                { timeout: 5000 }
            ).catch(() => { });
            await page.waitForTimeout(300);

            console.log('[GestionDoc] Intentando click Continuar hacia taller...');
            const resultadoAvance = await intentarAvanceRealHaciaTaller(
                page,
                `post-adjunto ${intento}/${maxIntentos}`,
                { maxClicks: 2 }
            );

            console.log(`[GestionDoc] resultadoAvance = ${resultadoAvance}`);
            console.log(`[GestionDoc] resultadoAvance = ${resultadoAvance}`);

            // Helper seguro: estaEnPantallaProductos con timeout para evitar bloqueos indefinidos
            const estaEnPantallaProductosSeguro = async (pageRef: Page): Promise<boolean> => {
                return Promise.race([
                    estaEnPantallaProductos(pageRef),
                    pageRef.waitForTimeout(1500).then(() => false),
                ]).catch(() => false);
            };

            // Helper fallback para seleccionar Productos por labels directos o por posición de dropdowns
            const procesarProductosPorFallbackDirecto = async (pageRef: Page, tipoCuenta: string) => {
                console.log('[Producto][FallbackDirecto] Procesando Productos por labels visibles');
                try {
                    // Diagnóstico inicial
                    const productosTextVisible = await pageRef.getByText(/^Productos$/i).first().isVisible({ timeout: 1000 }).catch(() => false);
                    console.log(`[Producto][FallbackDirecto][Diag] texto Productos visible=${productosTextVisible}`);

                    // Intentar buscar labels
                    const categoriaLabel = pageRef.getByText(/Categor[ií]a de producto/i).first();
                    const productoLabel = pageRef.getByText(/^Producto$/i).first();
                    const categoriaLabelVisible = await categoriaLabel.isVisible({ timeout: 800 }).catch(() => false);
                    const productoLabelVisible = await productoLabel.isVisible({ timeout: 800 }).catch(() => false);

                    console.log(`[Producto][FallbackDirecto][Diag] label Categoria visible=${categoriaLabelVisible}`);
                    console.log(`[Producto][FallbackDirecto][Diag] label Producto visible=${productoLabelVisible}`);

                    // Si no encontramos labels por visibilidad, usar dropdowns por posición
                    let categoriDropdown: Locator;
                    let productoDropdown: Locator;

                    if (categoriaLabelVisible) {
                        // Usar labels si están visibles
                        categoriDropdown = categoriaLabel.locator('..').locator('div.p-dropdown, [data-pc-name="dropdown"]').first();
                        console.log('[Producto][FallbackDirecto][Diag] usando dropdown categoria por label');
                    } else {
                        // Fallback: buscar por posición de dropdowns
                        console.log('[Producto][FallbackDirecto][WARN] Label Categoría no visible; usando dropdowns visibles por posición');
                        const todosDropdowns = pageRef.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible, select:visible');
                        const countDropdowns = await todosDropdowns.count().catch(() => 0);
                        console.log(`[Producto][FallbackDirecto][Diag] dropdowns visibles en pantalla=${countDropdowns}`);

                        if (countDropdowns < 2) {
                            throw new Error(`[Producto][CRITICO] Pantalla Productos detectada pero no hay dropdowns suficientes para Categoría/Producto. count=${countDropdowns}`);
                        }

                        categoriDropdown = todosDropdowns.nth(0);
                        console.log('[Producto][FallbackDirecto][Diag] usando dropdown categoria index=0');

                        // leer aria-controls para debugging
                        const combobox = categoriDropdown.locator('[role="combobox"]').first();
                        const ariaControls = await combobox.getAttribute('aria-controls').catch(() => null);
                        console.log(`[Producto][FallbackDirecto][Aria] Categoría aria-controls='${ariaControls || 'null'}'`);
                    }

                    //multiple estrategias para abrir dropdown
                    let panelAbierto = false;
                    const estrategiasAbrir = [
                        async () => {
                            const trigger = categoriDropdown.locator('.p-dropdown-trigger, [data-pc-section="trigger"]').first();
                            if(await trigger.isVisible().catch(() => false)) { await trigger.click({ force: true }).catch(() => {}); return true; }
                            return false;
                        },
                        async () => {
                            const combo = categoriDropdown.locator('[role="combobox"]').first();
                            if(await combo.isVisible().catch(() => false)) { await combo.click({ force: true }).catch(() => {}); return true; }
                            return false;
                        },
                        async () => {
                            if(await categoriDropdown.isVisible().catch(() => false)) { await categoriDropdown.click({ force: true }).catch(() => {}); return true; }
                            return false;
                        },
                        async () => {
                            const input = categoriDropdown.locator('input').first();
                            if(await input.isVisible().catch(() => false)) { await input.focus().catch(() => {}); await pageRef.keyboard.press('ArrowDown').catch(() => {}); return true; }
                            return false;
                        },
                    ];

                    for (const estrategia of estrategiasAbrir) {
                        if (panelAbierto) break;
                        panelAbierto = await estrategia();
                        if (panelAbierto) await pageRef.waitForTimeout(300);
                    }

                    // Buscar panel por aria-controls primero, luego fallback
                    let panelCategoria = pageRef.locator('.p-dropdown-panel:visible, [role="listbox"]:visible').first();
                    const combobox = categoriDropdown.locator('[role="combobox"]').first();
                    const ariaControls = await combobox.getAttribute('aria-controls').catch(() => null);
                    if (ariaControls) {
                        const panelById = pageRef.locator(`#${ariaControls}`).first();
                        const panelByIdVisible = await panelById.isVisible().catch(() => false);
                        if (panelByIdVisible) panelCategoria = panelById;
                        console.log(`[Producto][FallbackDirecto][Aria] Categoría panel encontrado=${panelByIdVisible}`);
                    }

                    // validar que panel abriu y tiene opciones
                    const opcionesPanel = panelCategoria.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
                    const countOpciones = await opcionesPanel.count().catch(() => 0);
                    console.log(`[Producto][FallbackDirecto][Aria] Categoría opciones visibles=${countOpciones}`);

                    // Seleccionar "Cuentas de Efectivo" del panel correcto
                    console.log('[Producto][FallbackDirecto] Seleccionando Cuentas de Efectivo');
                    const categoriaOption = panelCategoria.getByText(/Cuentas de Efectivo/i).first();
                    const categoriaOptionVisible = await categoriaOption.isVisible({ timeout: 1200 }).catch(() => false);

                    // imprimir opciones para debug si no encuentra
                    if (!categoriaOptionVisible && countOpciones > 0) {
                        const primerasOptions: string[] = [];
                        for (let i = 0; i < Math.min(countOpciones, 10); i++) {
                            const texto = await opcionesPanel.nth(i).innerText().catch(() => '').then(t => t.trim()).catch(() => '');
                            if (texto) primerasOptions.push(texto);
                        }
                        console.log(`[Producto][FallbackDirecto][Diag] Opciones visibles categoría=[${primerasOptions.join(', ')}]`);
                    }

                    if (categoriaOptionVisible) {
                        await categoriaOption.click().catch(() => { });
                        console.log('[Producto][FallbackDirecto] Categoría seleccionada');
                        await pageRef.waitForTimeout(FAST_UI ? 800 : 1500);
                    } else {
                        console.log('[Producto][FallbackDirecto][WARN] Opción Cuentas de Efectivo no visible');
                        return false;
                    }

                    // IMPORTANTE: Recargar dropdowns despues de seleccionar Categoría porque el DOM cambió
                    const todosDropdownsDespues = pageRef.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible, select:visible');
                    const countDespues = await todosDropdownsDespues.count().catch(() => 0);
                    console.log(`[Producto][FallbackDirecto][Diag] dropdowns visibles despues de seleccionar categoria=${countDespues}`);

                    // Buscar dropdown Producto (ahora debe ser el segundo dropdown visible)
                    if (productoLabelVisible) {
                        productoDropdown = productoLabel.locator('..').locator('div.p-dropdown, [data-pc-name="dropdown"]').first();
                        console.log('[Producto][FallbackDirecto][Diag] usando dropdown producto por label');
                    } else {
                        console.log('[Producto][FallbackDirecto][WARN] Label Producto no visible; usando dropdowns visibles por posición');
                        if (countDespues < 2) {
                            throw new Error(`[Producto][CRITICO] Pantalla Productos detectada pero no hay dropdown Producto. count=${countDespues}`);
                        }
                        productoDropdown = todosDropdownsDespues.nth(1);
                        console.log('[Producto][FallbackDirecto][Diag] usando dropdown producto index=1');
                    }

                    // Obtener aria-controls para Producto
                    const productoCombobox = productoDropdown.locator('[role="combobox"]').first();
                    const productoAriaControls = await productoCombobox.getAttribute('aria-controls').catch(() => null);
                    console.log(`[Producto][FallbackDirecto][Aria] Producto aria-controls='${productoAriaControls || 'null'}'`);

                    // Abrir dropdown producto con estrategias multiples
                    let panelProductoAbierto = false;
                    for (const estrategia of estrategiasAbrir) {
                        if (panelProductoAbierto) break;
                        panelProductoAbierto = await estrategia();
                        if (panelProductoAbierto) await pageRef.waitForTimeout(300);
                    }

                    // Buscar panel producto
                    let panelProducto = pageRef.locator('.p-dropdown-panel:visible, [role="listbox"]:visible').first();
                    if (productoAriaControls) {
                        const panelProductoById = pageRef.locator(`#${productoAriaControls}`).first();
                        const panelProductoByIdVisible = await panelProductoById.isVisible().catch(() => false);
                        if (panelProductoByIdVisible) panelProducto = panelProductoById;
                    }

                    const opcionesProducto = panelProducto.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
                    const countOpcionesProducto = await opcionesProducto.count().catch(() => 0);
                    console.log(`[Producto][FallbackDirecto][Aria] Producto opciones visibles=${countOpcionesProducto}`);

                    // Seleccionar producto
                    console.log(`[Producto][FallbackDirecto] Seleccionando producto ${tipoCuenta}`);
                    const productoOption = pageRef.getByText(new RegExp(escapeRegexText(tipoCuenta), 'i')).first();
                    const productoOptionVisible = await productoOption.isVisible({ timeout: 1200 }).catch(() => false);
                    if (productoOptionVisible) {
                        await productoOption.click().catch(() => { });
                        console.log('[Producto][FallbackDirecto] Producto seleccionado');
                        await pageRef.waitForTimeout(400);
                    } else {
                        console.log('[Producto][FallbackDirecto][WARN] Opción de Producto no visible');
                        return false;
                    }

                    // Ejecutar etapa post-selección
                    const seccionProductos = await localizarSeccionProductos(pageRef).catch(() => pageRef.locator('body'));
                    await etapaSeccionProductosPostSeleccion(pageRef, { tipoCuenta, identificacion: 'TMP' } as unknown as RegistroExcel, seccionProductos).catch(() => { });

                    console.log('[Producto][FallbackDirecto] Etapa post-selección completada');
                    return true;
                } catch (e) {
                    console.log(`[Producto][FallbackDirecto][ERROR] ${e}`);
                    return false;
                }
            };

            if (resultadoAvance === 'productos') {
                console.log('[Producto] resultadoAvance=productos recibido; procesando pantalla Productos inmediatamente');
                console.log('[Producto][Trace] entrada ruta inmediata productos');

                console.log('[Producto][Trace] antes de estaEnPantallaProductosSeguro');
                const yaEnProductos = await estaEnPantallaProductosSeguro(page);
                const timedOut = !yaEnProductos;
                console.log(`[Producto][Trace] despues de estaEnPantallaProductosSeguro valor=${yaEnProductos} timeout=${timedOut}`);

                // Check visual signals independently
                const categoriaVisible = await page.getByText(/Categor[ií]a de producto/i).first().isVisible().catch(() => false);
                const productoVisible = await page.getByText(/^Producto$/i).first().isVisible().catch(() => false);
                const agregarRelacionadoVisible = await page.getByRole('button', { name: /Agregar relacionado/i }).isVisible().catch(() => false);
                const productosTextVisible = await page.getByText(/^Productos$/i).isVisible().catch(() => false);

                const visualSignalsCount = [categoriaVisible, productoVisible, agregarRelacionadoVisible, productosTextVisible].filter(Boolean).length;

                if (!yaEnProductos && visualSignalsCount > 0) {
                    console.log(`[Producto][DIAG] estaEnPantallaProductos=false pero hay señales visuales (${visualSignalsCount}/4): categoría=${categoriaVisible}, producto=${productoVisible}, agregarRelacionado=${agregarRelacionadoVisible}, productosText=${productosTextVisible}`);
                }

                if (yaEnProductos || visualSignalsCount >= 3 || timedOut) {
                    if (!yaEnProductos && timedOut) {
                        console.log('[Producto] Forzando fallback directo porque resultadoAvance=productos');
                    } else if (!yaEnProductos) {
                        console.log('[Producto] Forzando procesamiento basado en señales visuales confirmadas');
                    }

                    // Validar si producto ya está agregado
                    const tipoDesdeExcel = await page.locator('body').innerText({ timeout: 3000 }).catch(() => '').then(t => {
                        const match = t.match(/200\s*-\s*Cuentas?\s+de?\s+Ahorros?\s*(?:con\s+Libreta)?/i);
                        return match ? match[0].trim() : '200 - Cuentas de Ahorros con Libreta DOP';
                    }).catch(() => '200 - Cuentas de Ahorros con Libreta DOP');

                    const tipoACargar = await tipoDesdeExcel;
                    console.log(`[Producto] Tipo de cuenta a cargar: ${tipoACargar}`);

                    // Verificar si producto ya está agregado
                    console.log('[Producto][Trace] verificando si producto ya está agregado');
                    const productoYaAgregado = await productoAgregadoVisible(page, tipoACargar).catch(() => false);
if (productoYaAgregado) {
                        console.log('[Producto] Producto ya está agregado, saltando agregación');
                    } else {
                        // Usar estrategia legacy de localización que funciona en Certificado-ex.spec.ts
                        console.log('[Producto][Legacy] resultadoAvance=productos; usando localizador legacy de Certificado como referencia técnica');
                        console.log('[Producto][Legacy] Categoria objetivo=\'Cuentas de Efectivo\'');
                        console.log(`[Producto][Legacy] Producto objetivo='${tipoACargar}'`);

                        // Wrapper seguro con timeout y fallback
                        const localizeWithTimeout = async () => {
                            return await Promise.race([
                                localizarSeccionProductosLegacy(page).then((locator) => locator),
                                new Promise<null>((resolve) => setTimeout(() => resolve(null), 1500))
                            ]);
                        };

                        console.log('[Producto][Legacy][Trace] antes de localizeSeccionProductosLegacySeguro');
                        let seccion: Locator | null = null;
                        let tiempoAgotado = false;
                        try {
                            seccion = await localizeWithTimeout();
                            tiempoAgotado = seccion === null;
                        } catch (e) {
                            tiempoAgotado = true;
                            console.log(`[Producto][Legacy][Trace] error en localizeWithTimeout: ${e}`);
                        }
                        const despuesOk = !!seccion && !(seccion === null);
                        console.log(`[Producto][Legacy][Trace] despues de localizeSeccionProductosLegacySeguro ok=${despuesOk} timeout=${tiempoAgotado}`);

                        // Fallback: usar body como scope temporal si no se localizó
                        if (!seccion) {
                            console.log('[Producto][Legacy][FallbackScope] no se localizó seccionProductos, usando fallback');
                            const todosDropdowns = page.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
                            const dropdownsGlobales = await todosDropdowns.count().catch(() => 0);
                            console.log(`[Producto][Legacy][FallbackScope] dropdownsGlobales=${dropdownsGlobales}`);
                            if (dropdownsGlobales >= 2) {
                                seccion = page.locator('body');
                                console.log('[Producto][Legacy][FallbackScope] usando body como scope temporal=true');
                            }
                        }

                        if (!seccion) {
                            // Fallback adicional: buscar dropdowns visibles y usar un ancestro común
                            const primerDropdown = page.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').first();
                            const padre = primerDropdown.locator('xpath=ancestor::*[self::fieldset or self::section or self::div][1]').first();
                            const visible = await padre.isVisible().catch(() => false);
                            if (visible) {
                                seccion = padre;
                                console.log('[Producto][Legacy][FallbackScope] usando ancestro común de dropdown como scope');
                            }
                        }

                        if (!seccion) {
                            throw new Error('[Producto][Legacy][CRITICO] No se pudo localizar sección Productos con estrategia legacy ni fallback. seccion=null');
                        }

                        const seccionVisible = await seccion.isVisible({ timeout: 1500 }).catch(() => false);
                        console.log(`[Producto][Legacy] seccionProductos visible=${seccionVisible}`);

                        const dropdownsVisibles = await seccion
                            .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
                            .count()
                            .catch(() => 0);
                        console.log(`[Producto][Legacy] dropdowns visibles en seccion=${dropdownsVisibles}`);

                        let scopeNombre = 'legacy';
                        let usarBandaVisual = false;
                        let bandaDropdowns: { categoria: Locator; producto: Locator } | null = null;

                        if (!seccionVisible || dropdownsVisibles < 2) {
                            console.log('[Producto][Legacy][WARN] fallback no tiene dropdowns suficientes, intentando scope específico por labels');

                            const scopoEspecifico = await localizarScopeProductosPorLabelsYDropdowns(page);
                            if (scopoEspecifico) {
                                seccion = scopoEspecifico;
                                scopeNombre = 'productos';
                            } else {
                                // Tercer nivel: polling con banda visual (hasta 8s)
                                bandaDropdowns = await esperarDropdownsProductosListos(page);
                                usarBandaVisual = true;
                                scopeNombre = 'banda';
                            }
                        }

                        // Guard: detectar si aparece modal Cancelar antes de seleccionar
                        const modalCancelarAntes = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
                        if (modalCancelarAntes) {
                            throw new Error('[Producto][CRITICO] Se abrió modal Cancelar proceso de solicitud antes de seleccionar; scope incorrecto');
                        }

                        if (usarBandaVisual && bandaDropdowns) {
                            // Ruta banda visual: selección directa por dropdown locator
                            console.log(`[Producto][Banda] Seleccionando Categoría: Cuentas de Efectivo`);
                            const valorCategoria = await seleccionarOpcionEnDropdownDirecto(
                                page,
                                bandaDropdowns.categoria,
                                /Cuentas de Efectivo/i,
                                'Producto][Banda][Categoria'
                            );
                            console.log(`[Producto][Banda] Categoría seleccionada valor='${valorCategoria}'`);
                            await page.waitForTimeout(FAST_UI ? 800 : 1500);

                            const modalCat = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
                            if (modalCat) throw new Error('[Producto][CRITICO] Se abrió modal Cancelar proceso de solicitud durante selección de categoría; click fuera de banda Productos');

                            const urlDespuesCat = page.url();
                            console.log(`[Producto][Guard] despues de seleccionar categoria url=${urlDespuesCat}`);
                            if (!urlDespuesCat.includes('/requests/') || !urlDespuesCat.includes('/edit')) {
                                throw new Error(`[Producto][CRITICO] URL cambió después de seleccionar categoría. url=${urlDespuesCat}`);
                            }

                            // Reconsultar dropdowns después de Categoría (Producto puede habilitarse/cambiar)
                            console.log('[Producto][Banda] reconsultando dropdowns después de Categoría');
                            const bandaReconsulta = await esperarDropdownsProductosListos(page);
                            console.log('[Producto][Banda] usando segundo dropdown como Producto');

                            console.log(`[Producto][Banda] Seleccionando Producto: ${tipoACargar}`);
                            const valorProducto = await seleccionarOpcionEnDropdownDirecto(
                                page,
                                bandaReconsulta.producto,
                                /200\s*-\s*Cuentas\s+de\s+Ahorros\s+con\s+Libreta\s+DOP/i,
                                'Producto][Banda][Producto'
                            );
                            console.log(`[Producto][Banda] Producto seleccionado valor='${valorProducto}'`);

                            const modalProd = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
                            if (modalProd) throw new Error('[Producto][CRITICO] Se abrió modal Cancelar proceso de solicitud durante selección de producto; click fuera de banda Productos');

                            const urlDespuesProd = page.url();
                            console.log(`[Producto][Guard] despues de seleccionar producto url=${urlDespuesProd}`);
                            if (!urlDespuesProd.includes('/requests/') || !urlDespuesProd.includes('/edit')) {
                                throw new Error(`[Producto][CRITICO] URL cambió después de seleccionar producto. url=${urlDespuesProd}`);
                            }
                        } else {
                            // Ruta legacy/scope: selección por funciones de sección
                            console.log(`[Producto][Legacy] llamando seleccionarCategoria Cuentas de Efectivo scope=${scopeNombre}`);
                            await seleccionarCategoriaEnSeccionProductosLegacy(page, seccion);
                            await page.waitForTimeout(FAST_UI ? 800 : 1500);

                            const modalCat = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
                            if (modalCat) throw new Error('[Producto][CRITICO] Se abrió modal Cancelar proceso de solicitud durante selección de categoría; scope incorrecto');

                            const urlDespuesCat = page.url();
                            console.log(`[Producto][Guard] despues de seleccionar categoria url=${urlDespuesCat}`);
                            if (!urlDespuesCat.includes('/requests/') || !urlDespuesCat.includes('/edit')) {
                                throw new Error(`[Producto][CRITICO] URL cambió después de seleccionar categoría. url=${urlDespuesCat}`);
                            }

                            console.log('[Producto][Legacy] llamando seleccionarProductoEnSeccionProductos');
                            await seleccionarProductoEnSeccionProductos(page, seccion, tipoACargar);

                            const modalProd = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
                            if (modalProd) throw new Error('[Producto][CRITICO] Se abrió modal Cancelar proceso de solicitud durante selección de producto; scope incorrecto');

                            const urlDespuesProd = page.url();
                            console.log(`[Producto][Guard] despues de seleccionar producto url=${urlDespuesProd}`);
                            if (!urlDespuesProd.includes('/requests/') || !urlDespuesProd.includes('/edit')) {
                                throw new Error(`[Producto][CRITICO] URL cambió después de seleccionar producto. url=${urlDespuesProd}`);
                            }
                        }

                        // Post-selección (común a ambas rutas)
                        await etapaSeccionProductosPostSeleccion(page, { tipoCuenta: tipoACargar, identificacion: 'TMP' } as unknown as RegistroExcel, seccion);

                        const urlDespuesPostSeleccion = page.url();
                        console.log(`[Producto][Guard] despues de post-seleccion url=${urlDespuesPostSeleccion}`);
                        if (!urlDespuesPostSeleccion.includes('/requests/') || !urlDespuesPostSeleccion.includes('/edit')) {
                            throw new Error(`[Producto][CRITICO] URL cambió después de post-selección. url=${urlDespuesPostSeleccion}`);
                        }

                        // Validar que se agregó
                        await page.waitForTimeout(300);
                        const productoConfirmado = await productoAgregadoVisible(page, tipoACargar);
                        if (!productoConfirmado) {
                            throw new Error(`[Producto][Legacy][CRITICO] No se pudo confirmar producto agregado después de selección. producto='${tipoACargar}'`);
                        }
                        console.log('[Producto][Legacy] producto agregado visible=true');
                    }

                    // Confirmación final
                    const productoConfirmadoFinal = await productoAgregadoVisible(page, tipoACargar).catch(() => false);
                    console.log(`[Producto] Producto agregado visible=${productoConfirmadoFinal}`);
                    if (!productoConfirmadoFinal) {
                        console.log('[Producto][CRITICO] Pantalla Productos visible pero no se pudo seleccionar Categoría/Producto');
                    }
                } else {
                    console.log('[Producto] Productos no detectado como pantalla activa y sin señales visuales confirmadas. No procesando.');
                }
            }
            if (resultadoAvance === 'taller' || resultadoAvance === 'productos') {
                return true;
            }

            const verificacionEspecialDespues = await procesarVerificacionesEspeciales(page).catch(() => null);
            if (verificacionEspecialDespues) {
                console.log(`[GestionDoc] Verificacion especial detectada (${verificacionEspecialDespues.tipo}) despues del intento post-adjunto.`);
                return true;
            }
        }
    }

    console.log('[GestionDoc] For loop completado o exahustado. Verificando estado final...');
    await cerrarModalGestionDocumentalSiVisible(page).catch(() => false);
    const sigueRequerido = await requiereAdjuntarDocumentoIdentidad(page).catch(() => false);
    const pendientesFinal = await leerPendientesGestionDocumentalEnBoton(page).catch(() => 0);
    if (sigueRequerido) {
        throw new Error("[CRITICO] Gestion documental sigue requerida tras reintentos; no se puede continuar.");
    }
    return true;
}

async function leerPendientesGestionDocumentalEnBoton(page: Page) {
    const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
    const visible = await btnGestion.isVisible().catch(() => false);
    if (!visible) return 0;

    const badge = btnGestion.locator('[data-pc-name="badge"], .p-badge, span[class*="badge"]').first();
    const badgeText = ((await badge.textContent({ timeout: 500 }).catch(() => "")) || "").trim();
    const ariaLabel = ((await btnGestion.getAttribute('aria-label', { timeout: 500 }).catch(() => "")) || "").trim();
    const textoBoton = ((await btnGestion.textContent({ timeout: 500 }).catch(() => "")) || "").trim();
    const sourceText = [badgeText, ariaLabel, textoBoton].join(" ");
    const matchPendientes = sourceText.match(/\b(\d+)\b/);
    const pendientes = matchPendientes ? Number.parseInt(matchPendientes[1], 10) : 0;
    return Number.isFinite(pendientes) ? pendientes : 0;
}

async function resolverGestionDocumentalPostContinuarPrimeraPantalla(
    page: Page,
    options?: { maxIntentos?: number }
) {
    const maxIntentos = options?.maxIntentos ?? 3;
    const evidenciaPath = resolverArchivoCedula();
    let huboGestionDocumental = false;

    for (let intento = 1; intento <= maxIntentos; intento++) {
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 12000 : 18000).catch(() => false);
        await page.waitForTimeout(FAST_UI ? 180 : 500);

        const pendientes = await leerPendientesGestionDocumentalEnBoton(page).catch(() => 0);
        const requiereAdjunto = await requiereAdjuntarDocumentoIdentidad(page).catch(() => false);
        // El badge puede quedar en 1 aunque el documento este cargado (contador de persona).
        // Solo considerar requerido si el modal esta abierto o hay alertas reales.
        if (!requiereAdjunto) {
            return huboGestionDocumental;
        }

        console.log(
            `[GestionDoc][PostContinuar] Pendientes detectados=${pendientes}. Intento ${intento}/${maxIntentos}.`
        );
        huboGestionDocumental = true;

        const modalGestion = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
            .first();

        const modalVisible = await modalGestion.isVisible().catch(() => false);
        if (!modalVisible) {
            const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
            const btnVisible = await btnGestion.isVisible().catch(() => false);

            if (!btnVisible) {
                throw new Error(
                    "[CRITICO] Se detecto requerimiento de Gestion documental pero no se encontro el boton para abrirla."
                );
            }

            await btnGestion.scrollIntoViewIfNeeded().catch(() => { });
            await btnGestion.click({ force: true }).catch(() => { });
        }

        const adjunto = await cargarDocumentoEnGestionDocumental(page, evidenciaPath)
            .then(() => true)
            .catch((e) => {
                console.log(`[GestionDoc][PostContinuar][WARN] carga documental fallo: ${String(e)}`);
                return false;
            });

        if (!adjunto) {
            throw new Error("[CRITICO] No se pudo adjuntar el documento requerido en Gestion documental.");
        }

        const modalCerrado = await cerrarModalGestionDocumentalSiVisible(page).catch(() => false);
        if (!modalCerrado) {
            throw new Error("[CRITICO] El modal documental no cerro limpiamente luego de adjuntar.");
        }

        await page.waitForTimeout(FAST_UI ? 500 : 1200);
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 12000 : 18000).catch(() => false);

        const badgeCero = await esperarPendientesGestionDocumentalCero(page).catch(() => false);
        const pendientesPostAdjunto = await leerPendientesGestionDocumentalEnBoton(page).catch(() => -1);

        console.log(
            `[GestionDoc][PostContinuar] Estado post-adjunto: badge=${pendientesPostAdjunto}, listo=${badgeCero}, modalCerrado=${modalCerrado}.`
        );


        const resultadoAvance = await intentarAvanceRealHaciaTaller(
            page,
            `post-continuar primera pantalla ${intento}/${maxIntentos}`,
            { maxClicks: 2 }
        );

        if (resultadoAvance === 'taller') {
            console.log("[GestionDoc][PostContinuar] Avance correcto hacia Taller despues de adjuntar documento.");
            return true;
        }

        const pendientesDespues = await leerPendientesGestionDocumentalEnBoton(page).catch(() => 0);
        const sigueRequerido = await requiereAdjuntarDocumentoIdentidad(page).catch(() => false);

        if (resultadoAvance === 'gestion-documental' || sigueRequerido || pendientesDespues > 0) {
            console.log(
                `[GestionDoc][PostContinuar] La gestion documental sigue apareciendo. Reintentando ciclo. pendientes=${pendientesDespues}`
            );
            continue;
        }

        console.log(
            "[GestionDoc][PostContinuar] Documentos OK, pero aun no hubo avance real. Devuelvo control al flujo principal para nuevo intento."
        );
        return huboGestionDocumental;
    }

    throw new Error("[CRITICO] Gestion documental sigue requerida luego de adjuntar y cerrar el modal.");
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

    // Fallback: cualquier fieldset o div con >= 2 dropdowns visibles que contenga
    // la palabra "Producto" (sin acentos exigidos, amplio).
    const fallbackCandidatos = page.locator(
        'fieldset:visible, form > div:visible, .p-panel-content:visible, [data-pc-section="content"]:visible'
    );
    const total = await fallbackCandidatos.count().catch(() => 0);
    for (let i = 0; i < total; i++) {
        const el = fallbackCandidatos.nth(i);
        const txt = (await el.textContent().catch(() => '')) || '';
        if (!/product/i.test(txt) && !/categor/i.test(txt)) continue;
        const drops = await el.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').count().catch(() => 0);
        if (drops >= 2) {
            console.log(`[localizarSeccionProductos] Fallback match con ${drops} dropdowns (texto: "${txt.substring(0, 60).replace(/\s+/g,' ')}")`);
            return el;
        }
    }

    // Ultimo recurso: retornar el ancestro comun de los primeros 2 dropdowns visibles de la pagina
    const primerDropdown = page.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').first();
    const ancestro = primerDropdown.locator('xpath=ancestor::fieldset[1] | xpath=ancestor::form[1] | xpath=ancestor::div[@class][1]').first();
    const ancestroVisible = await ancestro.isVisible().catch(() => false);
    if (ancestroVisible) {
        console.log('[localizarSeccionProductos] Ultimo recurso: ancestro del primer dropdown.');
        return ancestro;
    }

    return candidatos[0];
}

async function localizarScopeProductosPorLabelsYDropdowns(page: Page): Promise<Locator | null> {
    console.log('[Producto][ScopeProductos] Buscando scope por labels Categoría/Producto');

    // Buscar label "Categoría de producto"
    const labelCategoria = page
        .locator('xpath=//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Categoria de producto")]')
        .first();
    const labelCategoriaVisible = await labelCategoria.isVisible().catch(() => false);
    console.log(`[Producto][ScopeProductos] label Categoria visible=${labelCategoriaVisible}`);

    // Buscar label "Producto"
    const labelProducto = page
        .locator('xpath=//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Producto")] and not(contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Categoria"))')
        .first();
    const labelProductoVisible = await labelProducto.isVisible().catch(() => false);
    console.log(`[Producto][ScopeProductos] label Producto visible=${labelProductoVisible}`);

    if (!labelCategoriaVisible || !labelProductoVisible) {
        console.log('[Producto][ScopeProductos] scope encontrado=false');
        return null;
    }

    // Desde el label Categoría, buscar ancestros en orden de preferencia (más específicos primero)
    let scopeCandidate: Locator | null = null;
    const ancestorCandidatos = [
        labelCategoria.locator('xpath=ancestor::fieldset[1]').first(),
        labelCategoria.locator('xpath=ancestor::section[1]').first(),
        labelCategoria.locator('xpath=ancestor::div[@class="p-card"][1]').first(),
        labelCategoria.locator('xpath=ancestor::div[@data-pc-section="content"][1]').first(),
    ];

    for (const candidato of ancestorCandidatos) {
        const visible = await candidato.isVisible().catch(() => false);
        if (!visible) continue;

        const dropdownsEnCandidato = await candidato
            .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
            .count()
            .catch(() => 0);

        if (dropdownsEnCandidato >= 2) {
            scopeCandidate = candidato;
            break;
        }
    }

    if (!scopeCandidate) {
        console.log('[Producto][ScopeProductos] scope encontrado=false');
        return null;
    }

    const scopeVisible = await scopeCandidate.isVisible().catch(() => false);
    if (!scopeVisible) {
        console.log('[Producto][ScopeProductos] scope encontrado=false');
        return null;
    }

    // Contar dropdowns en el scope candidato
    const dropdownsEnScope = await scopeCandidate
        .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
        .count()
        .catch(() => 0);

    console.log(`[Producto][ScopeProductos] dropdowns visibles en scope=${dropdownsEnScope}`);

    if (dropdownsEnScope < 2) {
        console.log('[Producto][ScopeProductos] scope encontrado=false');
        return null;
    }

    // Verificar que contiene realmente "Producto"
    const scopeText = (await scopeCandidate.textContent().catch(() => '')) || '';
    const tieneProducto = /product/i.test(scopeText);
    const tieneCat = /categor/i.test(scopeText);

    if (!tieneProducto && !tieneCat) {
        console.log('[Producto][ScopeProductos] scope encontrado=false');
        return null;
    }

    // CRÍTICO: Validar que el scope NO contiene botones globales peligrosos
    const textoBotonesGlobales = /\b(Salir|Cancelar|Regresar|Gestión\s+Documental|Continuar)\b/i;
    if (textoBotonesGlobales.test(scopeText)) {
        console.log('[Producto][ScopeProductos][CRITICO] Scope incluye botones globales; scope rechazado');
        console.log('[Producto][ScopeProductos][Diag] scope contiene palabras prohibidas');
        return null;
    }

    console.log('[Producto][ScopeProductos] scope encontrado=true');
    console.log('[Producto][ScopeProductos] usando scope Productos específico');
    return scopeCandidate;
}

async function esperarDropdownsProductosListos(
    page: Page
): Promise<{ categoria: Locator; producto: Locator } | null> {
    const POLL_TIMEOUT = 8000;
    const POLL_INTERVAL = 300;
    const inicio = Date.now();
    let intento = 0;
    let dropdownsGlobalesMax = 0;

    console.log('[Producto][Banda] Esperando render de dropdowns Productos hasta 8s');

    while (Date.now() - inicio < POLL_TIMEOUT) {
        intento++;

        // Encabezado "Productos"
        const headProductos = page
            .locator('xpath=//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::span or self::div or self::p][normalize-space(.)="Productos" or normalize-space(.)="Producto"]')
            .first();
        const boxProductos = await headProductos.boundingBox().catch(() => null);

        // Encabezado "Relacionados" como límite inferior (opcional)
        const headRelacionados = page
            .locator('xpath=//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::span or self::div or self::p or self::button][contains(normalize-space(.),"Relacionados") or contains(normalize-space(.),"relacionado")]')
            .first();
        const boxRelacionados = await headRelacionados.boundingBox().catch(() => null);

        // Selectores amplios: incluye div.p-dropdown, span.p-dropdown, [data-pc-name], [role=combobox]
        const candidatosRaw = page.locator(
            '.p-dropdown:visible, [data-pc-name="dropdown"]:visible, [role="combobox"]:visible, span.p-dropdown-label:visible'
        );
        const totalRaw = await candidatosRaw.count().catch(() => 0);

        // Normalizar cada candidato al ancestro .p-dropdown más cercano via evaluate
        const dropdownsNormalizados: Array<{ locator: Locator; box: { y: number; x: number } }> = [];
        const seenHandles = new Set<string>();

        for (let i = 0; i < totalRaw; i++) {
            const el = candidatosRaw.nth(i);
            // Subir al wrapper .p-dropdown si el elemento es un descendiente (combobox, label)
            const ancestroHandle = await el.evaluate((node) => {
                const wrapper = node.closest('.p-dropdown, [data-pc-name="dropdown"]');
                if (!wrapper) return null;
                // Retornar un identificador único basado en posición DOM
                const rect = wrapper.getBoundingClientRect();
                return `${Math.round(rect.x)}_${Math.round(rect.y)}_${Math.round(rect.width)}`;
            }).catch(() => null);

            const key = ancestroHandle ?? `raw_${i}`;
            if (seenHandles.has(key)) continue;
            seenHandles.add(key);

            const box = await el.boundingBox().catch(() => null);
            if (!box) continue;
            dropdownsNormalizados.push({ locator: el, box: { y: box.y, x: box.x } });
        }

        const totalGlobales = dropdownsNormalizados.length;
        if (totalGlobales > dropdownsGlobalesMax) dropdownsGlobalesMax = totalGlobales;

        // Filtrar por banda si tenemos bounding boxes de encabezados
        let candidatosBanda = dropdownsNormalizados;
        if (boxProductos) {
            const yMin = boxProductos.y;
            const yMax = boxRelacionados ? boxRelacionados.y : yMin + 600;
            const enBanda = dropdownsNormalizados.filter(d => d.box.y >= yMin && d.box.y < yMax);
            candidatosBanda = enBanda.length >= 2 ? enBanda : dropdownsNormalizados;
        }

        // Ordenar por Y luego X
        candidatosBanda.sort((a, b) => a.box.y !== b.box.y ? a.box.y - b.box.y : a.box.x - b.box.x);
        const dropdownsBanda = candidatosBanda.length;

        console.log(`[Producto][Banda] intento ${intento}: dropdownsGlobales=${totalGlobales} dropdownsBanda=${dropdownsBanda}${dropdownsBanda < 2 ? '; esperando render de Productos' : ''}`);

        if (dropdownsBanda >= 2) {
            const catLoc = candidatosBanda[0].locator;
            const prodLoc = candidatosBanda[1].locator;
            const catBox = await catLoc.boundingBox().catch(() => null);
            const prodBox = await prodLoc.boundingBox().catch(() => null);
            console.log(`[Producto][Banda] dropdowns Productos listos=true count=${dropdownsBanda}`);
            console.log(`[Producto][Banda] categoria bbox=${JSON.stringify(catBox)}`);
            console.log(`[Producto][Banda] producto bbox=${JSON.stringify(prodBox)}`);
            console.log('[Producto][Banda] usando primer dropdown como Categoría');
            console.log('[Producto][Banda] usando segundo dropdown como Producto');
            return { categoria: catLoc, producto: prodLoc };
        }

        await page.waitForTimeout(POLL_INTERVAL);
    }

    const urlActual = page.url();
    throw new Error(`[Producto][CRITICO] Productos visible pero no aparecieron 2 dropdowns luego de 8s. url=${urlActual} dropdownsGlobalesMax=${dropdownsGlobalesMax}`);
}

async function seleccionarOpcionEnDropdownDirecto(
    page: Page,
    dropdown: Locator,
    opcionRegex: RegExp,
    nombreLog: string
): Promise<string> {
    const LIST_PANEL_TIMEOUT_LOCAL = 8000;
    const LIST_PANEL_QUICK_LOCAL = 3000;

    await dropdown.scrollIntoViewIfNeeded().catch(() => {});

    // Abrir dropdown con varias estrategias
    let panelAbierto = false;
    const estrategias = [
        async () => {
            const trigger = dropdown.locator('.p-dropdown-trigger, [data-pc-section="trigger"]').first();
            if (await trigger.isVisible().catch(() => false)) {
                await trigger.click({ force: true });
                return true;
            }
            return false;
        },
        async () => {
            const combo = dropdown.locator('[role="combobox"]').first();
            if (await combo.isVisible().catch(() => false)) {
                await combo.click({ force: true });
                return true;
            }
            return false;
        },
        async () => {
            if (await dropdown.isVisible().catch(() => false)) {
                await dropdown.click({ force: true });
                return true;
            }
            return false;
        },
        async () => {
            const input = dropdown.locator('input').first();
            if (await input.isVisible().catch(() => false)) {
                await input.focus();
                await page.keyboard.press('ArrowDown');
                return true;
            }
            return false;
        },
    ];

    for (const estrategia of estrategias) {
        if (panelAbierto) break;
        panelAbierto = await estrategia().catch(() => false);
        if (panelAbierto) await page.waitForTimeout(300);
    }

    // Localizar panel por aria-controls o fallback
    const combobox = dropdown.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
    let panel: Locator | null = null;

    if (panelId) {
        const byId = page.locator(`#${panelId}`);
        const visible = await byId.waitFor({ state: 'visible', timeout: LIST_PANEL_QUICK_LOCAL }).then(() => true).catch(() => false);
        if (visible) panel = byId;
    }
    if (!panel) {
        const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
        const visible = await fallback.waitFor({ state: 'visible', timeout: LIST_PANEL_QUICK_LOCAL }).then(() => true).catch(() => false);
        if (visible) panel = fallback;
    }

    if (!panel) {
        throw new Error(`[${nombreLog}][CRITICO] No se pudo abrir panel del dropdown`);
    }

    const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
    await items.first().waitFor({ state: 'visible', timeout: LIST_PANEL_TIMEOUT_LOCAL }).catch(() => {});

    const itemTarget = items.filter({ hasText: opcionRegex }).first();
    const itemVisible = await itemTarget.isVisible().catch(() => false);

    if (!itemVisible) {
        const countItems = await items.count().catch(() => 0);
        const primeras: string[] = [];
        for (let i = 0; i < Math.min(countItems, 8); i++) {
            const txt = (await items.nth(i).innerText().catch(() => '')).trim();
            if (txt) primeras.push(txt);
        }
        throw new Error(`[${nombreLog}][CRITICO] Opción no encontrada en panel. regex=${opcionRegex} opciones=[${primeras.join(', ')}]`);
    }

    await itemTarget.click({ force: true });
    await page.waitForTimeout(120);

    const labelEl = dropdown.locator('.p-dropdown-label, [data-pc-section="label"]').first();
    const valorSeleccionado = ((await labelEl.textContent().catch(() => '')) || '').trim();
    return valorSeleccionado;
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
        const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
        let panel: Locator | null = null;
        if (panelId) {
            const byId = page.locator(`#${panelId}`);
            const visible = await byId.waitFor({ state: 'visible', timeout: LIST_PANEL_QUICK_TIMEOUT })
                .then(() => true)
                .catch(() => false);
            if (visible) panel = byId;
        }
        if (!panel) {
            const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
            const visible = await fallback.waitFor({ state: 'visible', timeout: LIST_PANEL_QUICK_TIMEOUT })
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
            const listo = await items.first().waitFor({ state: 'visible', timeout: LIST_PANEL_TIMEOUT })
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
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
    }

    throw new Error("[CRITICO] No se pudo seleccionar 'Categoria de producto' en la seccion de Productos.");
}

async function esperarFinActualizandoSolicitud(page: Page, timeoutMs = 120000) {
    const inicio = Date.now();
    const txtActualizando = page.getByText(/Actualizando solicitud/i).first();
    const txtDepurando = page.getByText(/Depurando solicitante|Consultando datos del solicitante/i).first();
    const overlaysReales = page.locator(
        '.p-blockui:visible, [data-pc-name="blockui"]:visible, .p-progressspinner:visible, .p-progress-spinner:visible'
    );
    const progressbars = page.locator('[role="progressbar"]:visible');

    while (Date.now() - inicio < timeoutMs) {
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        const actualizandoVisible = await txtActualizando.isVisible().catch(() => false);
        const depurandoVisible = await txtDepurando.isVisible().catch(() => false);
        const overlayCount = await overlaysReales.count().catch(() => 0);
        // progressbar solo cuenta como bloqueo si hay texto de carga activo (evita falso-positivo con progressbar permanente)
        const progressbarCount = (actualizandoVisible || depurandoVisible)
            ? await progressbars.count().catch(() => 0)
            : 0;

        if (!actualizandoVisible && !depurandoVisible && overlayCount === 0 && progressbarCount === 0) {
            return true;
        }

        await page.waitForTimeout(350);
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
        // "Gestión Documental" button is only visible on the datos-laborales/financieros form page,
        // not on the products/taller page — reliable marker that we're still pre-products.
        page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first(),
    ];

    for (const m of marcadores) {
        if (await m.isVisible().catch(() => false)) {
            return true;
        }
    }
    return false;
}

async function existeCorreoEnContactos(page: Page, correo?: string) {
    const seccion = getSeccionContactos(page);
    await seccion.waitFor({ state: 'visible', timeout: 15000 });

    const cardCorreo = seccion
        .getByText(/^\s*Correo\s+Electr[oó]nico\s*$/i)
        .first();
    const visibleCardCorreo = await cardCorreo.isVisible().catch(() => false);
    if (visibleCardCorreo) return true;

    if (correo) {
        const visibleCorreoFijo = await seccion
            .getByText(new RegExp(String(correo).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), 'i'))
            .first()
            .isVisible()
            .catch(() => false);
        if (visibleCorreoFijo) return true;
    }
    return false;
}

async function existeCorreoEnPagina(page: Page, correo: string) {
    const safe = String(correo).trim().toLowerCase();
    if (!safe) return false;
    const bodyTxt = ((await page.locator('body').innerText().catch(() => '')) || '').toLowerCase();
    if (bodyTxt.includes(safe)) return true;
    const html = ((await page.content().catch(() => '')) || '').toLowerCase();
    return html.includes(safe);
}

async function leerEstadoDropdown(
    page: Page,
    labelText: string | RegExp,
    options?: { timeoutMs?: number }
) {
    const timeoutMs = options?.timeoutMs ?? 15000;
    const pattern = asPattern(labelText);
    let label = page.locator('label').filter({ hasText: pattern }).first();
    if ((await label.count().catch(() => 0)) === 0) {
        label = page.getByText(pattern).first();
    }
    const labelVisible = await label.waitFor({ state: 'visible', timeout: timeoutMs }).then(() => true).catch(() => false);
    if (!labelVisible) {
        return { vacio: true, texto: "" };
    }

    let dropdown = label
        .locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"][1]')
        .first();
    if (!(await dropdown.isVisible().catch(() => false))) {
        dropdown = label
            .locator('xpath=following::div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"][1]')
            .first();
    }
    const dropdownVisible = await dropdown.waitFor({ state: 'visible', timeout: timeoutMs }).then(() => true).catch(() => false);
    if (!dropdownVisible) {
        return { vacio: true, texto: "" };
    }

    const labelEl = dropdown.locator('.p-dropdown-label, [data-pc-section="label"]').first();
    const texto = ((await labelEl.textContent().catch(() => "")) || "").trim();
    const clases = ((await labelEl.getAttribute("class").catch(() => "")) || "");

    const vacio = !texto || /^seleccione/i.test(texto) || clases.includes("p-placeholder");
    return { vacio, texto };
}

async function asegurarDropdownConIndexSiVacio(
    page: Page,
    labelText: string,
    index = 3
) {
    // Fast-path: si ya tiene valor, salir sin reintentos.
    const estadoInicial = await leerEstadoDropdown(page, labelText, { timeoutMs: LIST_PROBE_TIMEOUT }).catch(() => ({ vacio: true, texto: "" }));
    if (!estadoInicial.vacio) return;

    for (let intento = 1; intento <= LIST_MAX_RETRIES; intento++) {
        await clickReintentarListaSiVisible(page, labelText, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
        await seleccionarDefaultSiVacio(page, labelText, index).catch(() => { });

        let estado = await leerEstadoDropdown(page, labelText, { timeoutMs: LIST_PROBE_TIMEOUT }).catch(() => ({ vacio: true, texto: "" }));
        if (!estado.vacio) return;

        // Fallback fuerte: abrir y seleccionar index directo.
        await seleccionarDropdownFiltrableConReintentar(page, labelText, { index }, { maxIntentos: FAST_UI ? 1 : 2 }).catch(() => { });
        estado = await leerEstadoDropdown(page, labelText, { timeoutMs: LIST_PROBE_TIMEOUT }).catch(() => ({ vacio: true, texto: "" }));
        if (!estado.vacio) return;

        await page.waitForTimeout(LIST_RETRY_WAIT_MS);
    }

    throw new Error(`[CRITICO] '${labelText}' quedo vacio luego de reintentos.`);
}

async function asegurarOficialPromotor(page: Page) {
    await asegurarDropdownConIndexSiVacio(page, 'Oficial', 3);
    await asegurarDropdownConIndexSiVacio(page, 'Promotor', 3);
}

async function resolverInputNumeroIdentificacion(page: Page, timeoutMs = 20000) {
    const candidatos = [
        page
            .locator('label')
            .filter({ hasText: /N(?:u|\u00fa)mero de identificaci(?:o|\u00f3)n/i })
            .first()
            .locator('xpath=following::input[1]')
            .first(),
        page
            .locator('xpath=//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Numero de identificacion")]/following::input[1]')
            .first(),
        page.locator('input[placeholder*="identificacion" i], input[placeholder*="identificación" i]').first(),
    ];

    const timeoutPorIntento = Math.max(3000, Math.floor(timeoutMs / candidatos.length));
    for (const input of candidatos) {
        const visible = await input
            .waitFor({ state: 'visible', timeout: timeoutPorIntento })
            .then(() => true)
            .catch(() => false);
        if (visible) return input;
    }

    throw new Error("No se pudo localizar el input de 'Numero de identificacion'.");
}

async function resolverInputFechaExpiracion(page: Page, timeoutMs = 15000) {
    const candidatos = [
        page
            .locator('xpath=//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Fecha de expiracion")]/following::input[contains(@class,"p-inputtext")][1]')
            .first(),
        page.locator('input[placeholder*="expir" i]').first(),
    ];

    const timeoutPorIntento = Math.max(3000, Math.floor(timeoutMs / candidatos.length));
    for (const input of candidatos) {
        const visible = await input
            .waitFor({ state: 'visible', timeout: timeoutPorIntento })
            .then(() => true)
            .catch(() => false);
        if (visible) return input;
    }
    throw new Error("No se pudo localizar el input 'Fecha de expiracion'.");
}

function extraerFechaLimiteDesdeTexto(texto: string) {
    const match = String(texto || '').match(/(m(?:i|í)nima|m(?:a|á)xima)\s+permitida\s+es\s*(\d{2}-\d{2}-\d{4})/i);
    if (!match) return null;
    return {
        tipo: match[1].toLowerCase(),
        fecha: match[2],
    };
}

function sumarDiasAFechaDDMMYYYY(fecha: string, dias: number) {
    const [dd, mm, yyyy] = fecha.split('-').map(Number);
    const utc = new Date(Date.UTC(yyyy, mm - 1, dd));
    utc.setUTCDate(utc.getUTCDate() + dias);

    const ddOut = String(utc.getUTCDate()).padStart(2, '0');
    const mmOut = String(utc.getUTCMonth() + 1).padStart(2, '0');
    const yyyyOut = utc.getUTCFullYear();
    return `${ddOut}-${mmOut}-${yyyyOut}`;
}

async function leerFechaLimiteExpiracion(page: Page) {
    const mensajes = [
        page.getByText(/La fecha m(?:i|\u00ed)nima permitida es|La fecha m(?:a|\u00e1)xima permitida es/i).first(),
        page.locator('xpath=//*[contains(normalize-space(.),"La fecha minima permitida es") or contains(normalize-space(.),"La fecha máxima permitida es") or contains(normalize-space(.),"La fecha maxima permitida es")]').first(),
    ];

    for (const locator of mensajes) {
        const visible = await locator.isVisible().catch(() => false);
        if (!visible) continue;

        const texto = await locator.innerText().catch(() => '');
        const limite = extraerFechaLimiteDesdeTexto(texto);
        if (limite) return limite;
    }

    return null;
}

async function hayMensajeFechaExpiracionInvalida(page: Page) {
    return page
        .getByText(/La fecha m(?:i|\u00ed)nima permitida es|La fecha m(?:a|\u00e1)xima permitida es/i)
        .first()
        .isVisible()
        .catch(() => false);
}

function normalizarFechaADigitos(fecha: string) {
    return String(fecha || '').replace(/\D/g, '');
}

function parsearFechaDDMMYYYY(fecha: string) {
    const [dd, mm, yyyy] = String(fecha || '').split('-').map(Number);
    return { dd, mm, yyyy };
}

function nombreMesEspanol(monthIndex: number) {
    return [
        'enero',
        'febrero',
        'marzo',
        'abril',
        'mayo',
        'junio',
        'julio',
        'agosto',
        'septiembre',
        'octubre',
        'noviembre',
        'diciembre',
    ][monthIndex] ?? '';
}

async function seleccionarFechaDesdeCalendario(page: Page, fecha: string) {
    const { dd, mm, yyyy } = parsearFechaDDMMYYYY(fecha);
    const panel = page.locator('.p-datepicker:visible, [data-pc-name="datepicker"]:visible').first();
    const visible = await panel.isVisible().catch(() => false);
    if (!visible) return false;

    const prevBtn = panel.locator('button').filter({ has: page.locator('.pi-chevron-left, .pi-angle-left') }).first();
    const nextBtn = panel.locator('button').filter({ has: page.locator('.pi-chevron-right, .pi-angle-right') }).first();
    const header = panel.locator('.p-datepicker-title, .p-datepicker-header, [class*="datepicker-title"]').first();

    const objetivoMes = nombreMesEspanol(mm - 1);
    const objetivoHeader = `${objetivoMes} ${yyyy}`;

    for (let i = 0; i < 240; i++) {
        const textoHeader = (await header.innerText().catch(() => '')).toLowerCase().replace(/\s+/g, ' ').trim();
        if (textoHeader.includes(objetivoMes) && textoHeader.includes(String(yyyy))) {
            break;
        }

        const match = textoHeader.match(/([a-záéíóúñ]+)\s+(\d{4})/i);
        if (!match) break;

        const anioActual = Number(match[2]);
        const mesActualIdx = [
            'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
            'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre',
        ].findIndex((mes) => match[1].toLowerCase().includes(mes));

        const diffMeses = (yyyy - anioActual) * 12 + ((mm - 1) - Math.max(mesActualIdx, 0));
        if (diffMeses === 0) break;

        const boton = diffMeses > 0 ? nextBtn : prevBtn;
        const puedeClick = await boton.isVisible().catch(() => false);
        if (!puedeClick) break;
        await boton.click({ force: true }).catch(() => { });
        await page.waitForTimeout(80);
    }

    const celdasDia = panel.locator('td >> text=' + String(dd));
    const total = await celdasDia.count().catch(() => 0);
    for (let i = 0; i < total; i++) {
        const dia = celdasDia.nth(i);
        const celda = dia.locator('xpath=ancestor::td[1]').first();
        const disabled = await celda.getAttribute('class').catch(() => '');
        if (String(disabled || '').includes('disabled') || String(disabled || '').includes('other-month')) continue;
        await dia.click({ force: true }).catch(() => { });
        await page.waitForTimeout(120);
        return true;
    }

    const botonesDia = panel.getByRole('button', { name: new RegExp(`^${dd}$`) });
    const totalBotones = await botonesDia.count().catch(() => 0);
    for (let i = 0; i < totalBotones; i++) {
        const boton = botonesDia.nth(i);
        if (!(await boton.isVisible().catch(() => false))) continue;
        await boton.click({ force: true }).catch(() => { });
        await page.waitForTimeout(120);
        return true;
    }

    console.log(`[ClienteNuevo][WARN] No se pudo seleccionar '${objetivoHeader}' dia ${dd} desde el calendario.`);
    return false;
}

async function escribirFechaEnInput(input: Locator, page: Page, fecha: string) {
    await input.scrollIntoViewIfNeeded().catch(() => { });
    await input.click({ force: true }).catch(() => { });
    await page.keyboard.press('Control+A').catch(() => { });
    await page.keyboard.press('Delete').catch(() => { });
    await input.fill('').catch(() => { });
    await input.pressSequentially(fecha, { delay: 60 }).catch(() => { });
    await input.dispatchEvent('input').catch(() => { });
    await input.dispatchEvent('change').catch(() => { });
    await input.blur().catch(() => { });
    await page.waitForTimeout(180);

    let valor = (await input.inputValue().catch(() => '')).trim();
    if (normalizarFechaADigitos(valor) === normalizarFechaADigitos(fecha)) {
        return valor;
    }

    const seleccionadaDesdeCalendario = await seleccionarFechaDesdeCalendario(page, fecha).catch(() => false);
    if (seleccionadaDesdeCalendario) {
        valor = (await input.inputValue().catch(() => '')).trim();
        if (normalizarFechaADigitos(valor) === normalizarFechaADigitos(fecha)) {
            return valor;
        }
    }

    await input.click({ force: true }).catch(() => { });
    await page.keyboard.press('Control+A').catch(() => { });
    await input.evaluate((el, v) => {
        const inp = el as HTMLInputElement;
        const nativeSetter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype,
            'value'
        )?.set;

        inp.focus();
        nativeSetter?.call(inp, String(v));
        inp.setAttribute('value', String(v));
        inp.dispatchEvent(new InputEvent('input', { bubbles: true, data: String(v), inputType: 'insertText' }));
        inp.dispatchEvent(new Event('change', { bubbles: true }));
        inp.dispatchEvent(new Event('blur', { bubbles: true }));
    }, fecha).catch(() => { });
    await page.waitForTimeout(150);

    return (await input.inputValue().catch(() => '')).trim();
}

async function setFechaExpiracionRobusta(page: Page, fecha = '09-01-2038') {
    const input = await resolverInputFechaExpiracion(page, 15000);
    let fechaObjetivo = fecha;
    let ultimoValor = '';
    const digitosObjetivoInicial = normalizarFechaADigitos(fecha);

    for (let i = 1; i <= 2; i++) {
        let valor = await escribirFechaEnInput(input, page, fechaObjetivo);
        ultimoValor = valor;
        if (normalizarFechaADigitos(valor) === normalizarFechaADigitos(fechaObjetivo)) return true;

        const sigueInvalidaTrasTecleo = await hayMensajeFechaExpiracionInvalida(page);
        if (!sigueInvalidaTrasTecleo && valor) return true;

        const limite = await leerFechaLimiteExpiracion(page);
        if (limite?.fecha && limite.fecha !== fechaObjetivo) {
            fechaObjetivo = limite.tipo.includes('min')
                ? sumarDiasAFechaDDMMYYYY(limite.fecha, 1)
                : limite.fecha;
        }

        valor = await escribirFechaEnInput(input, page, fechaObjetivo);
        ultimoValor = valor;
        if (normalizarFechaADigitos(valor) === normalizarFechaADigitos(fechaObjetivo)) return true;

        const sigueInvalidaTrasSetDirecto = await hayMensajeFechaExpiracionInvalida(page);
        if (!sigueInvalidaTrasSetDirecto && valor) return true;

        await page.waitForTimeout(120);
    }

    const valorFinalAceptable = normalizarFechaADigitos(ultimoValor) === digitosObjetivoInicial;
    if (valorFinalAceptable) {
        return true;
    }

    console.log(`[ClienteNuevo][WARN] No se pudo confirmar fecha de expiracion valida. objetivo='${fechaObjetivo}' valorFinal='${ultimoValor}'. Se continuara con Depurar.`);
    return false;
}

async function esperarFinConsultaSolicitante(page: Page, timeoutMs = 120000) {
    const spinnerRegex = /Consultando datos del solicitante|Depurando solicitante|Creando solicitud/i;
    const modalCasosActivos = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Casos?\s+activos?|caso activo|solicitud activa|caso ya existe|ya existe un caso/i })
        .first();

    await page.getByText(spinnerRegex).first().waitFor({ state: 'visible', timeout: 2500 }).catch(() => { });

    const inicio = Date.now();
    while (Date.now() - inicio < timeoutMs) {
        const modalVisible = await modalCasosActivos.isVisible().catch(() => false);
        if (modalVisible) {
            console.log('[ClienteNuevo] Modal de casos activos detectado durante la espera post-Depurar.');
            return;
        }

        const spinnerVisible = await page.getByText(spinnerRegex).first().isVisible().catch(() => false);
        if (!spinnerVisible) return;

        await page.waitForTimeout(250).catch(() => { });
    }
}

async function clickDepurarRobusto(page: Page, motivo = "") {
    const btnDepurar = page.getByRole('button', { name: /Depurar/i }).first();
    await btnDepurar.waitFor({ state: 'visible', timeout: 20000 });
    await btnDepurar.scrollIntoViewIfNeeded().catch(() => { });

    const enabled = await btnDepurar.isEnabled().catch(() => true);
    if (!enabled) {
        throw new Error("[CRITICO] El boton 'Depurar' esta deshabilitado.");
    }

    console.log(`[ClienteNuevo] Click en Depurar${motivo ? ` (${motivo})` : ''}...`);
    const urlAntes = page.url();
    await btnDepurar.click({ force: true, noWaitAfter: true }).catch(async () => {
        await btnDepurar.click({ noWaitAfter: true }).catch(() => { });
    });
    
    await esperarFinConsultaSolicitante(page, 120000);
    await esperarFinActualizandoSolicitud(page, 45000).catch(() => {});

    // After spinner clears, the portal may navigate to /requests/{id}/edit.
    // Wait up to 4s for a URL change; if none happens, we stayed on the same page (e.g. validation error).
    const urlTrasSpinner = page.url();
    if (urlTrasSpinner === urlAntes) {
        await page.waitForFunction(
            (antes: string) => location.href !== antes,
            urlAntes,
            { timeout: 4000 }
        ).catch(() => { });
    }
    await page.waitForLoadState('domcontentloaded').catch(() => { });
    // If we navigated to an edit page, wait for it to fully settle before the caller proceeds
    const urlFinal = page.url();
    if (/\/requests\/\d+/.test(urlFinal)) {
        console.log(`[ClienteNuevo] Portal navego a ${urlFinal} tras Depurar (${motivo}). Esperando estabilidad...`);
        await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => { });
    }
}

async function asegurarSeccionContactosDisponible(page: Page) {
    const btnAnadirContacto = getBotonAnadirContacto(page);
    const seccionContactos = getSeccionContactos(page);
    for (let intento = 1; intento <= 4; intento++) {
        const visibleContacto = await btnAnadirContacto
            .waitFor({ state: 'visible', timeout: 10000 })
            .then(() => true)
            .catch(() => false);
        if (visibleContacto) return true;

        const seccionVisible = await seccionContactos
            .waitFor({ state: 'visible', timeout: 2500 })
            .then(() => true)
            .catch(() => false);
        if (seccionVisible) {
            const indicadorContenido = seccionContactos
                .getByText(/No se han encontrado resultados|Correo\s*Electr[oó]nico|Tel[eé]fono|A(?:\u00f1|n)adir contacto/i)
                .first();
            const contenidoVisible = await indicadorContenido.isVisible().catch(() => false);
            if (contenidoVisible) return true;
        }

        const btnContinuar = getBotonContinuar(page);
        const puedeContinuar = await btnContinuar.isVisible().catch(() => false);
        if (puedeContinuar) {
            console.log('[ClienteNuevo] Contactos no visible; intentando Continuar para abrir etapa de datos del cliente...');
            try {
                await continuarResolviendoGestionDocumentalSiPide(page, { maxIntentos: 2 });
            } catch (e) {
                const msg = e instanceof Error ? e.message : String(e);
                if (/\[CRITICO\].*Gestion documental/i.test(msg)) throw e;
                await btnContinuar.click({ force: true }).catch(() => { });
                await page.waitForTimeout(1200);
            }
        }

        const paso2 = page.locator(
            'li.p-steps-item:has-text("2"), [data-pc-section="menuitem"]:has-text("2"), .p-steps-item:has(.p-steps-number:has-text("2"))'
        ).first();
        const paso2Visible = await paso2.isVisible().catch(() => false);
        if (paso2Visible) {
            await paso2.click({ force: true }).catch(() => { });
            await page.waitForTimeout(900);
        }

        if (intento === 3) {
            console.log('[ClienteNuevo][WARN] Contactos sigue sin aparecer en intento 3. Recargando pagina...');
            await page.reload({ waitUntil: 'domcontentloaded' }).catch(() => {});
            await esperarFinActualizandoSolicitud(page, 30000).catch(() => {});
        }
    }

    return false;
}

async function asegurarDropdownConIndex0SiVacio(page: Page, labelText: string | RegExp) {
    const estadoInicial = await leerEstadoDropdown(page, labelText, { timeoutMs: LIST_PROBE_TIMEOUT }).catch(() => ({ vacio: true, texto: "" }));
    if (!estadoInicial.vacio) return;

    const scopeRaiz = page.locator('body').first();
    const rapido = await seleccionarDropdownIndexRapidoEnScope(page, scopeRaiz, labelText, 0, {
        timeoutMs: FAST_UI ? 900 : 1400,
        panelTimeoutMs: FAST_UI ? 1400 : 2200,
        maxIntentos: FAST_UI ? 2 : 3,
    }).catch(() => false);
    if (rapido) {
        const estadoRapido = await leerEstadoDropdown(page, labelText, { timeoutMs: LIST_PROBE_TIMEOUT }).catch(() => ({ vacio: true, texto: "" }));
        if (!estadoRapido.vacio) {
            console.log(`[DropdownRapido] '${String(labelText)}' resuelto por via rapida con valor='${estadoRapido.texto}'.`);
            return;
        }
    }

    for (let intento = 1; intento <= LIST_MAX_RETRIES; intento++) {
        await clickReintentarListaSiVisible(page, labelText, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
        await seleccionarDefaultSiVacio(page, labelText, 0).catch(() => { });

        let estado = await leerEstadoDropdown(page, labelText, { timeoutMs: LIST_PROBE_TIMEOUT }).catch(() => ({ vacio: true, texto: "" }));
        if (!estado.vacio) return;

        const pattern = labelText instanceof RegExp ? labelText : new RegExp(labelText, "i");
        await seleccionarDropdownPorCampo(page, pattern, 0).catch(() => { });
        estado = await leerEstadoDropdown(page, labelText, { timeoutMs: LIST_PROBE_TIMEOUT }).catch(() => ({ vacio: true, texto: "" }));
        if (!estado.vacio) return;

        await page.waitForTimeout(LIST_RETRY_WAIT_MS);
    }
    throw new Error(`[CRITICO] '${String(labelText)}' quedo vacio en Cliente Nuevo.`);
}

async function seleccionarDropdownIndexEnScope(
    page: Page,
    scope: Locator,
    labelText: string | RegExp,
    index = 0,
    options?: {
        maxIntentos?: number;
        timeoutMs?: number;
        panelTimeoutMs?: number;
    }
) {
    const maxIntentos = options?.maxIntentos ?? LIST_MAX_RETRIES;
    const timeoutMs = options?.timeoutMs ?? LIST_LABEL_TIMEOUT;
    const panelTimeoutMs = options?.panelTimeoutMs ?? LIST_PANEL_TIMEOUT;

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

    for (let intento = 1; intento <= maxIntentos; intento++) {
        await cerrarPopupEdgeSyncSiVisible(page).catch(() => false);
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        await clickReintentarScope().catch(() => false);

        let label = scope.locator('label').filter({ hasText: pattern }).first();
        if ((await label.count().catch(() => 0)) === 0) {
            label = scope.getByText(pattern).first();
        }
        const labelVisible = await label.waitFor({ state: 'visible', timeout: timeoutMs })
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
        const fieldVisible = await field.waitFor({ state: 'visible', timeout: timeoutMs })
            .then(() => true)
            .catch(() => false);
        if (!fieldVisible) continue;

        await field.scrollIntoViewIfNeeded().catch(() => { });
        await field.click({ force: true }).catch(() => { });
        const cerroTrasClickField = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        if (cerroTrasClickField) {
            await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
            continue;
        }

        const combobox = field.locator('[role="combobox"]').first();
        const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
        let panel: Locator | null = null;
        if (panelId) {
            const byId = page.locator(`#${panelId}`);
            const byIdVisible = await byId.waitFor({ state: 'visible', timeout: Math.min(panelTimeoutMs, LIST_PANEL_QUICK_TIMEOUT) })
                .then(() => true)
                .catch(() => false);
            if (byIdVisible) panel = byId;
        }
        if (!panel) {
            const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
            const fallbackVisible = await fallback.waitFor({ state: 'visible', timeout: Math.min(panelTimeoutMs, LIST_PANEL_QUICK_TIMEOUT) })
                .then(() => true)
                .catch(() => false);
            if (fallbackVisible) panel = fallback;
        }
        if (!panel) {
            await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
            await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
            continue;
        }

        const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
        const hayItems = await items.first().waitFor({ state: 'visible', timeout: Math.min(3500, panelTimeoutMs) })
            .then(() => true)
            .catch(() => false);
        if (!hayItems) {
            await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
            await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
            continue;
        }

        const count = await items.count().catch(() => 0);
        if (count <= index || count === 0) {
            await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
            await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
            continue;
        }

        const item = items.nth(index);
        await item.scrollIntoViewIfNeeded().catch(() => { });
        await item.click({ force: true });

        const valueText = ((await field.locator('.p-dropdown-label, [data-pc-section="label"]').first().textContent().catch(() => '')) || '').trim();
        const sinSeleccion = !valueText || /^seleccione/i.test(valueText) || /Reintentar/i.test(valueText);
        if (!sinSeleccion) return;
    }

    throw new Error(`No se pudo seleccionar index ${index} para '${String(labelText)}' tras reintentos.`);
}

async function seleccionarDropdownIndexRapidoEnScope(
    page: Page,
    scope: Locator,
    labelText: string | RegExp,
    index = 0,
    options?: {
        timeoutMs?: number;
        panelTimeoutMs?: number;
        maxIntentos?: number;
    }
) {
    const timeoutMs = options?.timeoutMs ?? 1200;
    const panelTimeoutMs = options?.panelTimeoutMs ?? 1800;
    const maxIntentos = options?.maxIntentos ?? 2;
    const pattern = labelText instanceof RegExp
        ? labelText
        : new RegExp(String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

    const leerValorActual = async () => {
        const valor = await leerValorDropdownEnScope(scope, pattern, { timeoutMs: Math.min(timeoutMs, 900) }).catch(() => '');
        return valor.trim();
    };

    const tieneValorSeleccionado = async () => {
        const valor = await leerValorActual();
        return !!valor && !/^seleccione/i.test(valor) && !/reintentar/i.test(valor);
    };

    if (await tieneValorSeleccionado().catch(() => false)) return true;

    for (let intento = 1; intento <= maxIntentos; intento++) {
        let label = scope.locator('label').filter({ hasText: pattern }).first();
        if ((await label.count().catch(() => 0)) === 0) {
            label = scope.getByText(pattern).first();
        }
        const labelVisible = await label.waitFor({ state: 'visible', timeout: timeoutMs })
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
        const fieldVisible = await field.waitFor({ state: 'visible', timeout: timeoutMs })
            .then(() => true)
            .catch(() => false);
        if (!fieldVisible) continue;

        let dropdown = field.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();
        if (!(await dropdown.isVisible().catch(() => false))) {
            dropdown = field;
        }
        await dropdown.scrollIntoViewIfNeeded().catch(() => { });
        await dropdown.click({ force: true }).catch(() => { });

        const combobox = dropdown.locator('[role="combobox"]').first();
        const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
        let panel: Locator | null = null;
        if (panelId) {
            const byId = page.locator(`#${panelId}`);
            const byIdVisible = await byId.waitFor({ state: 'visible', timeout: panelTimeoutMs })
                .then(() => true)
                .catch(() => false);
            if (byIdVisible) panel = byId;
        }
        if (!panel) {
            const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
            const fallbackVisible = await fallback.waitFor({ state: 'visible', timeout: panelTimeoutMs })
                .then(() => true)
                .catch(() => false);
            if (fallbackVisible) panel = fallback;
        }
        if (!panel) {
            await page.keyboard.press('Escape').catch(() => { });
            await page.waitForTimeout(80);
            continue;
        }

        const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
        const itemVisible = await items.nth(index).waitFor({ state: 'visible', timeout: panelTimeoutMs })
            .then(() => true)
            .catch(() => false);

        if (itemVisible) {
            await items.nth(index).click({ force: true }).catch(() => { });
        } else {
            await dropdown.click({ force: true }).catch(() => { });
            for (let i = 0; i <= index; i++) {
                await page.keyboard.press('ArrowDown').catch(() => { });
                await page.waitForTimeout(40);
            }
            await page.keyboard.press('Enter').catch(() => { });
        }

        await page.waitForTimeout(90);
        if (await tieneValorSeleccionado().catch(() => false)) return true;

        await page.keyboard.press('Escape').catch(() => { });
        await page.waitForTimeout(90);
    }

    return false;
}

async function seleccionarDropdownDependienteConEspera(
    page: Page,
    scope: Locator,
    labelText: string | RegExp,
    index = 0,
    options?: {
        maxIntentos?: number;
        timeoutMs?: number;
        panelTimeoutMs?: number;
        esperaEntreIntentosMs?: number;
    }
) {
    const maxIntentos = options?.maxIntentos ?? (FAST_UI ? 7 : 10);
    const timeoutMs = options?.timeoutMs ?? (FAST_UI ? 4200 : 7000);
    const panelTimeoutMs = options?.panelTimeoutMs ?? (FAST_UI ? 4500 : 8000);
    const esperaEntreIntentosMs = options?.esperaEntreIntentosMs ?? (FAST_UI ? 280 : 900);

    let ultimoError: unknown = null;
    for (let intento = 1; intento <= maxIntentos; intento++) {
        try {
            await seleccionarDropdownIndexEnScope(page, scope, labelText, index, {
                maxIntentos: FAST_UI ? 3 : 5,
                timeoutMs,
                panelTimeoutMs,
            });
            return;
        } catch (e) {
            ultimoError = e;
            await clickReintentarListaSiVisible(page, labelText, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
            await page.waitForTimeout(esperaEntreIntentosMs);
        }
    }

    const msg = ultimoError instanceof Error ? ultimoError.message : String(ultimoError);
    throw new Error(`No se pudo seleccionar index ${index} para '${String(labelText)}' tras espera dependiente. ${msg}`);
}

async function asegurarEmpleoActualSi(modalLaboral: Locator) {
    const bloque = modalLaboral
        .locator('xpath=//*[contains(normalize-space(.),"El empleo es actual")]/ancestor::*[self::div or self::fieldset][1]')
        .first();
    const radioSi = bloque.getByRole('radio', { name: /^S(?:i|\u00ed)$/i }).first();
    const radioNo = bloque.getByRole('radio', { name: /^No$/i }).first();
    const radioSiFallback = bloque
        .locator('[role="radio"][aria-label*="Sí" i], [role="radio"][aria-label*="Si" i], [aria-label="Sí"], [aria-label="Si"]')
        .first();

    const siVisible = await radioSi.waitFor({ state: 'visible', timeout: 4000 }).then(() => true).catch(() => false);
    const targetSi = siVisible ? radioSi : radioSiFallback;
    await targetSi.waitFor({ state: 'visible', timeout: 12000 });

    for (let intento = 1; intento <= 5; intento++) {
        await targetSi.scrollIntoViewIfNeeded().catch(() => { });
        await targetSi.click({ force: true }).catch(async () => {
            await targetSi.click().catch(() => { });
        });
        await modalLaboral.page().waitForTimeout(150);

        const checkedSi = ((await targetSi.getAttribute('aria-checked').catch(() => '')) || '').toLowerCase() === 'true';
        const checkedNo = ((await radioNo.getAttribute('aria-checked').catch(() => '')) || '').toLowerCase() === 'true';

        const claseSi = ((await targetSi.getAttribute('class').catch(() => '')) || '').toLowerCase();
        const claseNo = ((await radioNo.getAttribute('class').catch(() => '')) || '').toLowerCase();
        const claseSiOn = /active|highlight|selected|primary/.test(claseSi);
        const claseNoOn = /active|highlight|selected|primary/.test(claseNo);

        if ((checkedSi && !checkedNo) || (claseSiOn && !claseNoOn)) return;
    }

    throw new Error("[CRITICO] No se pudo seleccionar 'Si' en 'El empleo es actual'.");
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
        await page.waitForTimeout(350);
        return true;
    };

    const maxIntentos = LIST_MAX_RETRIES;
    const labelTimeout = LIST_LABEL_TIMEOUT;
    const fieldTimeout = LIST_LABEL_TIMEOUT;
    const panelTimeout = LIST_PANEL_TIMEOUT;
    for (let intento = 1; intento <= maxIntentos; intento++) {
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        await clickReintentarScope().catch(() => false);
        await clickReintentarListaSiVisible(page, labelText, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);

        let label = scope.locator('label').filter({ hasText: pattern }).first();
        if ((await label.count().catch(() => 0)) === 0) {
            label = scope.getByText(pattern).first();
        }
        const labelVisible = await label.waitFor({ state: 'visible', timeout: labelTimeout })
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
        const fieldVisible = await field.waitFor({ state: 'visible', timeout: fieldTimeout })
            .then(() => true)
            .catch(() => false);
        if (!fieldVisible) continue;

        await field.scrollIntoViewIfNeeded().catch(() => { });
        await field.click({ force: true }).catch(() => { });
        const cerroTrasClickField = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        if (cerroTrasClickField) {
            await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
            continue;
        }

        const combobox = field.locator('[role="combobox"]').first();
        const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
        let panel: Locator | null = null;
        if (panelId) {
            const byId = page.locator(`#${panelId}`);
            const byIdVisible = await byId.waitFor({ state: 'visible', timeout: Math.min(panelTimeout, LIST_PANEL_QUICK_TIMEOUT) })
                .then(() => true)
                .catch(() => false);
            if (byIdVisible) panel = byId;
        }
        if (!panel) {
            const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
            const fallbackVisible = await fallback.waitFor({ state: 'visible', timeout: Math.min(panelTimeout, LIST_PANEL_QUICK_TIMEOUT) })
                .then(() => true)
                .catch(() => false);
            if (fallbackVisible) panel = fallback;
        }
        if (!panel) {
            await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
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

        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
    }

    throw new Error(`No se pudo seleccionar opcion '${String(textoOpcion)}' en '${String(labelText)}'.`);
}

async function leerValorDropdownEnScopeLegacy(
    scope: Locator,
    labelText: string | RegExp,
    options?: {
        timeoutMs?: number;
    }
) {
    const timeoutMs = options?.timeoutMs ?? 5000;
    const pattern = labelText instanceof RegExp
        ? labelText
        : new RegExp(String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

    let label = scope.locator('label').filter({ hasText: pattern }).first();
    if ((await label.count().catch(() => 0)) === 0) {
        label = scope.getByText(pattern).first();
    }
    await label.waitFor({ state: 'visible', timeout: timeoutMs });

    let field = label
        .locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]')
        .first();
    if (!(await field.isVisible().catch(() => false))) {
        field = label
            .locator('xpath=following::*[self::div[contains(@class,"p-dropdown")] or @data-pc-name="dropdown"][1]')
            .first();
    }
    await field.waitFor({ state: 'visible', timeout: timeoutMs });
    const texto = ((await field.locator('.p-dropdown-label, [data-pc-section="label"]').first().textContent().catch(() => '')) || '').trim();
    return texto;
}

async function llenarInputEnScope(
    scope: Locator,
    labelText: string | RegExp,
    valor: string,
    options?: {
        timeoutMs?: number;
        required?: boolean;
    }
) {
    const timeoutMs = options?.timeoutMs ?? 15000;
    const required = options?.required ?? true;
    const pattern = labelText instanceof RegExp
        ? labelText
        : new RegExp(String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

    let label = scope.locator('label').filter({ hasText: pattern }).first();
    if ((await label.count().catch(() => 0)) === 0) {
        label = scope.getByText(pattern).first();
    }
    const labelVisible = await label.waitFor({ state: 'visible', timeout: timeoutMs })
        .then(() => true)
        .catch(() => false);
    if (!labelVisible) {
        if (required) throw new Error(`No se encontro label '${String(labelText)}' para llenar input.`);
        return false;
    }

    const input = label.locator('xpath=following::input[1]').first();
    const inputVisible = await input.waitFor({ state: 'visible', timeout: timeoutMs })
        .then(() => true)
        .catch(() => false);
    if (!inputVisible) {
        if (required) throw new Error(`No se encontro input visible para '${String(labelText)}'.`);
        return false;
    }
    await input.scrollIntoViewIfNeeded().catch(() => { });
    await input.click({ force: true });
    await input.fill('');
    await input.fill(valor);
    await input.blur().catch(() => { });
    return true;
}

async function abrirModalContacto(page: Page) {
    const btnAnadirContacto = getBotonAnadirContacto(page);
    const modalPorTexto = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Contacto|Categor(?:i|\u00ed)a/i })
        .first();
    const modalPorEstructura = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({
            has: page.locator('label').filter({ hasText: /Categor(?:i|\u00ed)a/i }),
        })
        .first();

    for (let intento = 1; intento <= CONTACT_MODAL_ATTEMPTS; intento++) {
        await cerrarPopupEdgeSyncSiVisible(page).catch(() => false);
        await btnAnadirContacto.waitFor({ state: 'visible', timeout: CONTACT_MODAL_BTN_TIMEOUT });
        await btnAnadirContacto.scrollIntoViewIfNeeded().catch(() => { });
        await btnAnadirContacto.click({ force: true }).catch(() => { });

        const porTexto = await modalPorTexto
            .waitFor({ state: 'visible', timeout: CONTACT_MODAL_OPEN_TIMEOUT_1 })
            .then(() => true)
            .catch(() => false);
        if (porTexto) return modalPorTexto;

        const porEstructura = await modalPorEstructura
            .waitFor({ state: 'visible', timeout: CONTACT_MODAL_OPEN_TIMEOUT_2 })
            .then(() => true)
            .catch(() => false);
        if (porEstructura) return modalPorEstructura;

        await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
    }

    throw new Error("[CRITICO] No se pudo abrir el modal de Contacto tras reintentos.");
}

async function abrirModalLaboral(page: Page) {
    const btnAnadirLaboral = getBotonAnadirLaborales(page);
    const modalLaboral = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Informaci(?:o|\u00f3)n laboral|Nombre de la empresa|El empleo es actual/i })
        .first();

    for (let intento = 1; intento <= 4; intento++) {
        await btnAnadirLaboral.waitFor({ state: 'visible', timeout: 15000 });
        await btnAnadirLaboral.scrollIntoViewIfNeeded().catch(() => { });
        await btnAnadirLaboral.click({ force: true }).catch(() => { });

        const abierto = await modalLaboral
            .waitFor({ state: 'visible', timeout: 7000 })
            .then(() => true)
            .catch(() => false);
        if (abierto) return modalLaboral;

        await page.mouse.wheel(0, 300).catch(() => { });
        await page.waitForTimeout(350);
    }

    throw new Error("[CRITICO] No se pudo abrir modal de 'Informacion laboral' tras reintentos.");
}

async function llenarYAceptarContactoEnModal(
    page: Page,
    modalContacto: Locator,
    categoria: 'telefono' | 'correo',
    contactoValor: string
): Promise<'ok' | 'duplicado'> {
    const expectedCategoria = categoria === 'correo'
        ? /Correo\s*Electr[oó]nico/i
        : /Tel[eé]fono/i;

    const seleccionarTipoContactoRobusto = async () => {
        await seleccionarDropdownDependienteConEspera(page, modalContacto, /^Tipo$/i, 0, {
            maxIntentos: FAST_UI ? 7 : 10,
            timeoutMs: FAST_UI ? 4200 : 7000,
            panelTimeoutMs: FAST_UI ? 4800 : 8500,
            esperaEntreIntentosMs: FAST_UI ? 300 : 900,
        }).catch(async () => {
            await seleccionarDropdownIndexEnScope(page, modalContacto, /^Tipo$/i, 0, {
                maxIntentos: FAST_UI ? 4 : 6,
                timeoutMs: FAST_UI ? 3200 : 6000,
                panelTimeoutMs: FAST_UI ? 3600 : 7000,
            });
        });
    };

    const seleccionarCategoriaContactoRobusta = async () => {
        // Primer intento rapido por indice para Telefono (suele estar en la primera opcion).
        if (categoria === 'telefono') {
            await seleccionarDropdownIndexEnScope(page, modalContacto, /Categor(?:i|\u00ed)a/i, 0, {
                maxIntentos: 2,
                timeoutMs: 2200,
                panelTimeoutMs: 2200,
            }).catch(() => { });
            const valorRapido = await leerValorDropdownEnScope(modalContacto, /Categor(?:i|\u00ed)a/i).catch(() => '');
            if (expectedCategoria.test(valorRapido)) return true;
        }

        // Fallback por texto (util para Correo y para entornos con orden distinto en el combo).
        if (categoria === 'telefono') {
            await seleccionarDropdownEnScopePorTexto(
                page,
                modalContacto,
                /Categor(?:i|\u00ed)a/i,
                /Tel[eé]fono/i,
                0
            ).catch(() => { });
        } else {
            await seleccionarDropdownEnScopePorTexto(
                page,
                modalContacto,
                /Categor(?:i|\u00ed)a/i,
                /Correo\s*Electr[oó]nico/i,
                1
            ).catch(() => { });
        }

        const valor = await leerValorDropdownEnScope(modalContacto, /Categor(?:i|\u00ed)a/i).catch(() => '');
        return expectedCategoria.test(valor);
    };

    for (let intentoCat = 1; intentoCat <= 4; intentoCat++) {
        const seleccionOk = await seleccionarCategoriaContactoRobusta().catch(() => false);
        const valorCategoria = await leerValorDropdownEnScope(modalContacto, /Categor(?:i|\u00ed)a/i).catch(() => '');
        if (seleccionOk || expectedCategoria.test(valorCategoria)) break;
        if (intentoCat >= 4) {
            throw new Error(`[CRITICO] La categoria de contacto no quedo seleccionada como '${categoria}'. Valor actual: '${valorCategoria}'.`);
        }
        await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
    }

    const lblTipo = modalContacto.locator('label').filter({ hasText: /^Tipo$/i }).first();
    if (await lblTipo.isVisible().catch(() => false)) {
        await seleccionarTipoContactoRobusto().catch(() => { });
    }
    await llenarInputEnScope(modalContacto, /^Contacto$/i, contactoValor);
    const lblUbicacion = modalContacto.locator('label').filter({ hasText: /Ubicaci(?:o|\u00f3)n/i }).first();
    if (await lblUbicacion.isVisible().catch(() => false)) {
        await seleccionarDropdownIndexEnScope(page, modalContacto, /Ubicaci(?:o|\u00f3)n/i, 0, {
            maxIntentos: 3,
            timeoutMs: 2500,
            panelTimeoutMs: 2500,
        });
    }

    const reforzarCamposRequeridos = async () => {
        const valorCategoria = await leerValorDropdownEnScope(modalContacto, /Categor(?:i|\u00ed)a/i).catch(() => '');
        if (!expectedCategoria.test(valorCategoria)) {
            await seleccionarCategoriaContactoRobusta().catch(() => { });
        }

        const lblTipoLocal = modalContacto.locator('label').filter({ hasText: /^Tipo$/i }).first();
        if (await lblTipoLocal.isVisible().catch(() => false)) {
            await seleccionarTipoContactoRobusto().catch(() => { });
        }

        await llenarInputEnScope(modalContacto, /^Contacto$/i, contactoValor).catch(() => { });

        const lblUbicacionLocal = modalContacto.locator('label').filter({ hasText: /Ubicaci(?:o|\u00f3)n/i }).first();
        if (await lblUbicacionLocal.isVisible().catch(() => false)) {
            await seleccionarDropdownIndexEnScope(page, modalContacto, /Ubicaci(?:o|\u00f3)n/i, 0, {
                maxIntentos: 2,
                timeoutMs: 1800,
                panelTimeoutMs: 1800,
            }).catch(() => { });
        }
    };

    // A veces el modal no cierra al primer click; forzamos cierre con reintentos.
    const btnAceptar = modalContacto
        .locator('button:visible')
        .filter({ hasText: /^Aceptar$/i })
        .last();
    await btnAceptar.waitFor({ state: 'visible', timeout: 10000 });
    for (let intento = 1; intento <= 5; intento++) {
        await reforzarCamposRequeridos();

        const inputContacto = modalContacto
            .locator('label')
            .filter({ hasText: /^Contacto$/i })
            .first()
            .locator('xpath=following::input[1]')
            .first();
        const valorContacto = (await inputContacto.inputValue().catch(() => '')).trim();
        if (!valorContacto || (categoria === 'correo' && !valorContacto.includes('@'))) {
            await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
            continue;
        }

        const enabled = await btnAceptar.isEnabled().catch(() => true);
        if (!enabled) {
            await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
            continue;
        }

        await btnAceptar.scrollIntoViewIfNeeded().catch(() => { });
        await btnAceptar.click().catch(async () => {
            await btnAceptar.click({ force: true }).catch(() => { });
        });

        const modalDuplicado = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Contacto duplicado/i })
            .first();
        const esDuplicado = await modalDuplicado.isVisible().catch(() => false);
        if (esDuplicado) {
            const btnCerrarDuplicado = modalDuplicado.getByRole('button', { name: /^Cerrar$/i }).first();
            if (await btnCerrarDuplicado.isVisible().catch(() => false)) {
                await btnCerrarDuplicado.click({ force: true }).catch(() => { });
                await modalDuplicado.waitFor({ state: 'hidden', timeout: 8000 }).catch(() => { });
            }
            const btnCancelarContacto = modalContacto.getByRole('button', { name: /^Cancelar$/i }).first();
            if (await btnCancelarContacto.isVisible().catch(() => false)) {
                await btnCancelarContacto.click({ force: true }).catch(() => { });
                await modalContacto.waitFor({ state: 'hidden', timeout: 8000 }).catch(() => { });
            }
            return 'duplicado';
        }

        const cerro = await modalContacto
            .waitFor({ state: 'hidden', timeout: 5000 })
            .then(() => true)
            .catch(() => false);
        if (cerro) return 'ok';

        const hayRequerido = await modalContacto
            .locator('.p-error:visible, small:visible, span:visible')
            .filter({ hasText: /es requerido|required/i })
            .first()
            .isVisible()
            .catch(() => false);
        if (hayRequerido) {
            await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
            continue;
        }

        await page.waitForTimeout(CONTACT_SHORT_WAIT_MS);
    }

    throw new Error("[CRITICO] El modal de Contacto quedo con campos requeridos pendientes al intentar Aceptar.");
}

async function agregarContactoClienteNuevo(
    page: Page,
    categoria: 'telefono' | 'correo',
    contactoValor: string
) {
    const modalContacto = await abrirModalContacto(page);
    return await llenarYAceptarContactoEnModal(page, modalContacto, categoria, contactoValor);
}

async function validarContactoAgregadoEnSeccion(
    page: Page,
    contactoValor: string,
    tipo: 'telefono' | 'correo',
    options?: { strict?: boolean }
) {
    const strict = options?.strict ?? true;
    const seccionContactos = getSeccionContactos(page);
    await seccionContactos.waitFor({ state: 'visible', timeout: 20000 });
    await seccionContactos.scrollIntoViewIfNeeded().catch(() => { });

    for (let intento = 1; intento <= CONTACT_VALIDATE_ATTEMPTS; intento++) {
        const sinResultados = seccionContactos.getByText(/No se han encontrado resultados/i).first();
        const vacio = await sinResultados.isVisible().catch(() => false);
        const textoSeccion = ((await seccionContactos.innerText().catch(() => '')) || '').toLowerCase();

        if (!vacio) {
            if (tipo === 'correo') {
                const cardCorreo = await seccionContactos
                    .getByText(/Correo\s*Electr[oó]nico/i)
                    .first()
                    .isVisible()
                    .catch(() => false);
                if (cardCorreo) return true;

                if (textoSeccion.includes(contactoValor.toLowerCase())) return true;
            } else {
                const cardTelefono = await seccionContactos
                    .getByText(/Tel[eé]fono/i)
                    .first()
                    .isVisible()
                    .catch(() => false);
                if (cardTelefono) return true;

                const digitosEsperados = contactoValor.replace(/\D/g, '');
                const digitosSeccion = textoSeccion.replace(/\D/g, '');
                const candidatoLargo = digitosEsperados.length >= 7 ? digitosEsperados.slice(-7) : digitosEsperados;
                if (digitosSeccion.includes(digitosEsperados) || digitosSeccion.includes(candidatoLargo)) return true;
            }
        }
        await page.waitForTimeout(CONTACT_VALIDATE_POLL_MS);
    }

    // Fallback global: si la seccion no refleja bien el card pero ya esta en pantalla.
    const textoGlobal = ((await page.locator('body').innerText().catch(() => '')) || '').toLowerCase();
    if (tipo === 'telefono') {
        const digitosEsperados = contactoValor.replace(/\D/g, '');
        const digitosGlobal = textoGlobal.replace(/\D/g, '');
        const candidatoLargo = digitosEsperados.length >= 7 ? digitosEsperados.slice(-7) : digitosEsperados;
        if (textoGlobal.includes('teléfono') || textoGlobal.includes('telefono')) {
            if (digitosGlobal.includes(digitosEsperados) || digitosGlobal.includes(candidatoLargo)) return true;
        }
    }

    if (!strict) {
        return false;
    }

    if (tipo === 'correo') {
        const existeCorreoGlobal = await existeCorreoEnPagina(page, contactoValor).catch(() => false);
        if (existeCorreoGlobal) return true;
        throw new Error(`[CRITICO] No se encontro el correo agregado (${contactoValor}) en la seccion Contactos.`);
    }
    throw new Error(`[CRITICO] No se encontro el telefono agregado (${contactoValor}) en la seccion Contactos.`);
}

async function completarDatosClienteNuevo(page: Page) {
    const btnAnadirContacto = getBotonAnadirContacto(page);
    const visibleContactos = await btnAnadirContacto.isVisible().catch(() => false);
    if (!visibleContactos) return false;

    console.log('[ClienteNuevo] Iniciando llenado obligatorio...');

    await btnAnadirContacto.scrollIntoViewIfNeeded().catch(() => { });

    // Secuencia obligatoria: Telefono -> Aceptar -> Correo -> Aceptar.
    const telefono = randomTelefono();
    await agregarContactoClienteNuevo(page, 'telefono', telefono);
    const telefonoOk = await validarContactoAgregadoEnSeccion(page, telefono, 'telefono', { strict: false });
    if (telefonoOk) {
        console.log(`[ClienteNuevo] Telefono agregado OK: ${telefono}`);
    } else {
        console.log(`[ClienteNuevo][WARN] No se pudo confirmar telefono por lectura UI, pero se continua al correo: ${telefono}`);
    }

    // Paso explicito: despues de aceptar telefono, volver a pulsar "Añadir contacto"
    // para abrir un NUEVO modal y agregar correo.
    const correo = CORREO_FIJO_CLIENTE_NUEVO;

    const yaHayCorreo = await existeCorreoEnContactos(page, undefined);
    const yaTieneCorreoFijo = await existeCorreoEnContactos(page, correo);
    if (yaHayCorreo || yaTieneCorreoFijo) {
        console.log(`[ClienteNuevo] Correo ya existente en seccion Contactos. Se reutiliza: ${correo}`);
    }

    let correoAgregado = false;
    if (yaHayCorreo || yaTieneCorreoFijo) {
        correoAgregado = true;
    } else {
        for (let intentoCorreo = 1; intentoCorreo <= 3; intentoCorreo++) {
            try {
                console.log('[ClienteNuevo] Reabriendo "Añadir contacto" para correo...');
                const modalCorreo = await abrirModalContacto(page);
                const estadoCorreo = await llenarYAceptarContactoEnModal(page, modalCorreo, 'correo', correo);
                if (estadoCorreo === 'duplicado') {
                    console.log(`[ClienteNuevo] Correo duplicado detectado; se asume existente: ${correo}`);
                    correoAgregado = true;
                    break;
                }
                const correoConfirmado = await validarContactoAgregadoEnSeccion(page, correo, 'correo', { strict: false });
                if (!correoConfirmado) {
                    const existeCorreoFallback = await existeCorreoEnPagina(page, correo);
                    if (!existeCorreoFallback) {
                        throw new Error(`[CRITICO] No se pudo confirmar el correo agregado (${correo}).`);
                    }
                }
                correoAgregado = true;
                break;
            } catch (e) {
                const correoYaExiste = await existeCorreoEnPagina(page, correo).catch(() => false);
                if (correoYaExiste) {
                    console.log(`[ClienteNuevo] Correo detectado tras validacion/reintento. Se reutiliza: ${correo}`);
                    correoAgregado = true;
                    break;
                }
                const modalAbierto = page
                    .locator('.p-dialog:visible, [role="dialog"]:visible')
                    .filter({ hasText: /Contacto|Categor(?:i|\u00ed)a/i })
                    .first();
                const sigueAbierto = await modalAbierto.isVisible().catch(() => false);
                if (sigueAbierto) {
                    const btnCancelar = modalAbierto.getByRole('button', { name: /^Cancelar$/i }).first();
                    if (await btnCancelar.isVisible().catch(() => false)) {
                        await btnCancelar.click({ force: true }).catch(() => { });
                        await modalAbierto.waitFor({ state: 'hidden', timeout: 6000 }).catch(() => { });
                    }
                }
                if (intentoCorreo >= 3) throw e;
                await page.waitForTimeout(CONTACT_CORREO_RETRY_WAIT_MS);
            }
        }
    }
    if (!correoAgregado) {
        throw new Error('[CRITICO] No se pudo agregar el correo en Contactos.');
    }
    console.log(`[ClienteNuevo] Correo agregado OK: ${correo}`);

    await asegurarInformacionIngresosClienteNuevo(page);

    const modalLaboral = await abrirModalLaboral(page);

    // Orden horizontal (izq->der, fila por fila) para mantener consistencia visual.
    await llenarInputEnScope(modalLaboral, /Nombre de la empresa/i, randomTexto('Empresa'));
    await llenarInputEnScope(modalLaboral, /Departamento/i, randomTexto('Area'));
    await llenarInputEnScope(modalLaboral, /Puesto que ocupa/i, randomTexto('Puesto'));
    await seleccionarDropdownDependienteConEspera(page, modalLaboral, /^Cargo$/i, 0, {
        maxIntentos: FAST_UI ? 6 : 9,
        timeoutMs: FAST_UI ? 3600 : 6500,
        panelTimeoutMs: FAST_UI ? 4200 : 7800,
        esperaEntreIntentosMs: FAST_UI ? 260 : 750,
    });

    await llenarInputEnScope(modalLaboral, /^Salario$/i, String(randomInt(25000, 75000)));
    await asegurarEmpleoActualSi(modalLaboral);
    await llenarInputEnScope(modalLaboral, /Fecha de ingreso/i, '01-02-2022');

    // Si el campo esta editable y vacio, llenarlo para evitar bloqueos por validacion.
    const lblTiempoPuesto = modalLaboral.locator('label').filter({ hasText: /Tiempo en puesto/i }).first();
    if (await lblTiempoPuesto.isVisible().catch(() => false)) {
        const inputTiempo = lblTiempoPuesto.locator('xpath=following::input[1]').first();
        const editableTiempo = await inputTiempo.isEditable().catch(() => false);
        const valTiempo = (await inputTiempo.inputValue().catch(() => '')).trim();
        if (editableTiempo && !valTiempo) {
            await inputTiempo.fill('24').catch(() => { });
            await inputTiempo.blur().catch(() => { });
        }
    }

    await llenarInputEnScope(modalLaboral, /Tel[eé]fono/i, randomTelefono()).catch(() => { });
    await llenarInputEnScope(modalLaboral, /Extensi(?:o|\u00f3)n/i, String(randomInt(100, 999))).catch(() => { });

    await seleccionarDropdownDependienteConEspera(page, modalLaboral, /^Tipo$/i, 0, {
        maxIntentos: FAST_UI ? 8 : 12,
        timeoutMs: FAST_UI ? 4200 : 7000,
        panelTimeoutMs: FAST_UI ? 4800 : 9000,
        esperaEntreIntentosMs: FAST_UI ? 300 : 900,
    });
    await seleccionarDropdownDependienteConEspera(page, modalLaboral, /^Pa(?:i|\u00ed)s$/i, 0, {
        maxIntentos: FAST_UI ? 6 : 10,
        timeoutMs: FAST_UI ? 3600 : 6500,
        panelTimeoutMs: FAST_UI ? 4200 : 7800,
        esperaEntreIntentosMs: FAST_UI ? 260 : 750,
    });
    await llenarInputEnScope(modalLaboral, /Calle.*Avenida.*Autopista/i, randomTexto('Calle'));
    await llenarInputEnScope(modalLaboral, /Nombre edificio/i, randomTexto('Edificio')).catch(() => { });

    await llenarInputEnScope(modalLaboral, /N[uú]mero casa.*edificio/i, String(randomInt(1, 999)));
    await llenarInputEnScope(modalLaboral, /N[uú]mero apartamento/i, String(randomInt(1, 80)));
    await seleccionarDropdownIndexEnScope(page, modalLaboral, /^Regi(?:o|\u00f3)n$/i, 0, { maxIntentos: 3, timeoutMs: 2500, panelTimeoutMs: 3000 });
    await seleccionarDropdownIndexEnScope(page, modalLaboral, /^Provincia$/i, 0, { maxIntentos: 3, timeoutMs: 2500, panelTimeoutMs: 3000 });

    await seleccionarDropdownDependienteConEspera(page, modalLaboral, /^Municipio$/i, 0, {
        maxIntentos: FAST_UI ? 8 : 12,
        timeoutMs: FAST_UI ? 4200 : 7000,
        panelTimeoutMs: FAST_UI ? 4800 : 9000,
        esperaEntreIntentosMs: FAST_UI ? 300 : 900,
    });
    await seleccionarDropdownDependienteConEspera(page, modalLaboral, /^Localidad$/i, 0, {
        maxIntentos: FAST_UI ? 8 : 12,
        timeoutMs: FAST_UI ? 4200 : 7000,
        panelTimeoutMs: FAST_UI ? 4800 : 9000,
        esperaEntreIntentosMs: FAST_UI ? 300 : 900,
    });
    await seleccionarDropdownDependienteConEspera(page, modalLaboral, /^Sector$/i, 0, {
        maxIntentos: FAST_UI ? 7 : 10,
        timeoutMs: FAST_UI ? 3800 : 6500,
        panelTimeoutMs: FAST_UI ? 4400 : 8000,
        esperaEntreIntentosMs: FAST_UI ? 260 : 800,
    });
    await llenarInputEnScope(modalLaboral, /^Referencia$/i, randomTexto('Referencia'));

    const btnAceptarLaboral = modalLaboral.getByRole('button', { name: /^Aceptar$/i }).first();
    await btnAceptarLaboral.waitFor({ state: 'visible', timeout: 15000 });
    await btnAceptarLaboral.click({ force: true });
    await modalLaboral.waitFor({ state: 'hidden', timeout: 20000 });
    await page.waitForTimeout(FAST_UI ? 220 : 500);

    // Verificacion tolerante: en algunos layouts "No se han encontrado resultados"
    // pertenece a la seccion vecina y no debe bloquear el flujo.
    const contenedorLaboral = getBotonAnadirLaborales(page)
        .locator('xpath=ancestor::*[self::fieldset or self::div][1]')
        .first();
    const sinResultadosLaboral = contenedorLaboral.getByText(/No se han encontrado resultados/i).first();
    const hayCardLaboral = await contenedorLaboral
        .locator('text=/Fecha de ingreso|Tiempo en puesto|Salario/i')
        .first()
        .isVisible()
        .catch(() => false);
    const hayBotonEditarLaboral = await contenedorLaboral
        .locator('button:has(i.pi-pencil), button:has(.pi-pencil)')
        .first()
        .isVisible()
        .catch(() => false);

    if (!hayCardLaboral && !hayBotonEditarLaboral && await sinResultadosLaboral.isVisible().catch(() => false)) {
        console.log('[ClienteNuevo][WARN] No se pudo confirmar tarjeta laboral por lectura UI; se continua con el flujo.');
    }

    console.log('[ClienteNuevo] Llenado obligatorio completado.');
    return true;
}

async function asegurarInformacionIngresosClienteNuevo(page: Page) {
    const tituloIngresos = page.getByText(/Informaci(?:o|\u00f3)n de ingresos/i).first();
    await tituloIngresos.waitFor({ state: 'visible', timeout: 15000 });

    const seccionIngresos = tituloIngresos
        .locator('xpath=ancestor::*[self::fieldset or self::div][1]')
        .first();
    await seccionIngresos.scrollIntoViewIfNeeded().catch(() => { });

    const asegurarCampo = async (labelRegex: RegExp, nombreCampo: string) => {
        for (let intento = 1; intento <= (FAST_UI ? 5 : 7); intento++) {
            await clickReintentarListaSiVisible(page, labelRegex, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
            await seleccionarDropdownIndexEnScope(page, seccionIngresos, labelRegex, 0, {
                maxIntentos: FAST_UI ? 3 : 5,
                timeoutMs: FAST_UI ? 2600 : 5000,
                panelTimeoutMs: FAST_UI ? 2800 : 5000,
            }).catch(() => { });

            const valor = await leerValorDropdownEnScope(seccionIngresos, labelRegex).catch(() => '');
            const vacio = !valor || /^seleccione/i.test(valor) || /Reintentar/i.test(valor);
            if (!vacio) return;

            await page.waitForTimeout(FAST_UI ? 120 : 350);
        }

        throw new Error(`[CRITICO] '${nombreCampo}' quedo vacio en Informacion de ingresos.`);
    };

    await asegurarCampo(/Actividad econ(?:o|\u00f3)mica/i, 'Actividad economica');
    await asegurarCampo(/Tipo de cliente/i, 'Tipo de cliente');
    await asegurarCampo(/Ocupaci(?:o|\u00f3)n/i, 'Ocupacion');
}

async function seleccionarOpcionSiNoEnGrupo(
    page: Page,
    scope: Locator,
    opcion: 'Si' | 'No',
    options?: { required?: boolean; timeoutMs?: number }
) {
    const required = options?.required ?? true;
    const timeoutMs = options?.timeoutMs ?? 12000;
    const target = opcion === 'Si' ? /^S(?:i|\u00ed)$/i : /^No$/i;

    const grupo = scope
        .locator('.p-selectbutton, [data-pc-name="selectbutton"], [role="radiogroup"]')
        .first();
    const grupoVisible = await grupo.waitFor({ state: 'visible', timeout: timeoutMs }).then(() => true).catch(() => false);
    if (!grupoVisible) {
        if (required) throw new Error(`[CRITICO] No se encontro grupo Si/No para seleccionar '${opcion}'.`);
        return false;
    }

    const candidatos = [
        grupo.getByRole('radio', { name: target }).first(),
        grupo.getByRole('button', { name: target }).first(),
        grupo
            .locator(
                opcion === 'Si'
                    ? '[role="radio"][aria-label*="Sí" i], [role="radio"][aria-label*="Si" i], button:has-text("Sí"), button:has-text("Si")'
                    : '[role="radio"][aria-label*="No" i], button:has-text("No")'
            )
            .first(),
    ];

    let boton: Locator | null = null;
    for (const c of candidatos) {
        const visible = await c.isVisible().catch(() => false);
        if (visible) {
            boton = c;
            break;
        }
    }

    if (!boton) {
        if (required) throw new Error(`[CRITICO] No se encontro opcion '${opcion}' en grupo Si/No.`);
        return false;
    }

    for (let intento = 1; intento <= 4; intento++) {
        await boton.scrollIntoViewIfNeeded().catch(() => { });
        await boton.click({ force: true }).catch(async () => {
            await boton?.click().catch(() => { });
        });
        await page.waitForTimeout(120);

        const checked = ((await boton.getAttribute('aria-checked').catch(() => '')) || '').toLowerCase() === 'true';
        const pressed = ((await boton.getAttribute('aria-pressed').catch(() => '')) || '').toLowerCase() === 'true';
        const clase = ((await boton.getAttribute('class').catch(() => '')) || '').toLowerCase();
        const on = checked || pressed || /highlight|active|selected|primary/.test(clase);
        if (on) return true;
    }

    if (required) throw new Error(`[CRITICO] No se pudo seleccionar '${opcion}' en grupo Si/No.`);
    return false;
}

async function seleccionarRespuestaPregunta(
    page: Page,
    preguntaRegex: RegExp,
    opcion: 'Si' | 'No',
    options?: { required?: boolean }
) {
    const required = options?.required ?? true;
    const pregunta = page.getByText(preguntaRegex).first();
    const visible = await pregunta.waitFor({ state: 'visible', timeout: 7000 }).then(() => true).catch(() => false);
    if (!visible) {
        if (required) throw new Error(`[CRITICO] No se encontro la pregunta '${preguntaRegex}'.`);
        return false;
    }

    const bloque = pregunta
        .locator('xpath=ancestor::*[self::div or self::fieldset][1]')
        .first();
    await bloque.scrollIntoViewIfNeeded().catch(() => { });
    return seleccionarOpcionSiNoEnGrupo(page, bloque, opcion, { required, timeoutMs: 10000 });
}

async function asegurarDropdownIndex0SiVacioEnScope(
    page: Page,
    scope: Locator,
    labelText: string | RegExp,
    options?: {
        required?: boolean;
        maxIntentos?: number;
        timeoutMs?: number;
        panelTimeoutMs?: number;
    }
) {
    const required = options?.required ?? false;
    const maxIntentos = options?.maxIntentos ?? (FAST_UI ? 4 : 6);
    const timeoutMs = options?.timeoutMs ?? LIST_LABEL_TIMEOUT;
    const panelTimeoutMs = options?.panelTimeoutMs ?? LIST_PANEL_TIMEOUT;
    const pattern = labelText instanceof RegExp
        ? labelText
        : new RegExp(String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

    const label = scope.locator('label').filter({ hasText: pattern }).first();
    const labelVisible = await label.waitFor({ state: 'visible', timeout: timeoutMs })
        .then(() => true)
        .catch(() => false);
    if (!labelVisible) {
        if (required) throw new Error(`[CRITICO] No se encontro label '${String(labelText)}' dentro del modal.`);
        return false;
    }

    const valorActual = await leerValorDropdownEnScope(scope, pattern, { timeoutMs }).catch(() => '');
    const vacio = !valorActual || /^seleccione/i.test(valorActual) || /Reintentar/i.test(valorActual);
    if (!vacio) return false;

    await seleccionarDropdownIndexEnScope(page, scope, pattern, 0, {
        maxIntentos,
        timeoutMs,
        panelTimeoutMs,
    });
    return true;
}

async function seleccionarIndex0SiVacioEnDropdown(page: Page, dropdown: Locator) {
    const labelEl = dropdown.locator('.p-dropdown-label, [data-pc-section="label"]').first();
    const texto = ((await labelEl.textContent().catch(() => '')) || '').trim();
    const clases = ((await labelEl.getAttribute('class').catch(() => '')) || '');
    const vacio = !texto || /^seleccione/i.test(texto) || /Reintentar/i.test(texto) || clases.includes('p-placeholder');
    if (!vacio) return false;

    await dropdown.scrollIntoViewIfNeeded().catch(() => { });
    await dropdown.click({ force: true }).catch(() => { });

    const combobox = dropdown.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
    let panel: Locator | null = null;
    if (panelId) {
        const byId = page.locator(`#${panelId}`);
        const byIdVisible = await byId.waitFor({ state: 'visible', timeout: LIST_PANEL_QUICK_TIMEOUT })
            .then(() => true)
            .catch(() => false);
        if (byIdVisible) panel = byId;
    }
    if (!panel) {
        const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
        const fallbackVisible = await fallback.waitFor({ state: 'visible', timeout: LIST_PANEL_QUICK_TIMEOUT })
            .then(() => true)
            .catch(() => false);
        if (fallbackVisible) panel = fallback;
    }
    if (!panel) return false;

    const btnRetryPanel = panel
        .locator('button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button.p-button-warning')
        .first();
    const retryVisible = await btnRetryPanel.isVisible().catch(() => false);
    if (retryVisible) {
        await btnRetryPanel.click({ force: true }).catch(() => { });
        await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
    }

    const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
    const hayItem = await items.first().waitFor({ state: 'visible', timeout: LIST_PANEL_TIMEOUT })
        .then(() => true)
        .catch(() => false);
    if (!hayItem) return false;

    await items.nth(0).click({ force: true }).catch(() => { });
    return true;
}

async function completarDropdownsVaciosIndex0EnSeccion(page: Page, tituloRegex: RegExp) {
    const titulo = page.getByText(tituloRegex).first();
    const visibleTitulo = await titulo.isVisible().catch(() => false);
    if (!visibleTitulo) return false;

    const seccion = titulo.locator('xpath=ancestor::*[self::fieldset or self::div][1]').first();
    await seccion.scrollIntoViewIfNeeded().catch(() => { });

    const dropdowns = seccion.locator('div.p-dropdown, [data-pc-name="dropdown"]');
    const count = await dropdowns.count().catch(() => 0);
    if (!count) return false;

    for (let i = 0; i < count; i++) {
        await seleccionarIndex0SiVacioEnDropdown(page, dropdowns.nth(i)).catch(() => { });
    }
    return true;
}

async function completarCamposTextoOtrasInformaciones(page: Page) {
    const titulo = page.getByText(/Otras informaciones/i).first();
    const visibleTitulo = await titulo.isVisible().catch(() => false);
    if (!visibleTitulo) return false;

    const seccion = titulo.locator('xpath=ancestor::*[self::fieldset or self::div][1]').first();
    await seccion.scrollIntoViewIfNeeded().catch(() => { });

    await llenarInputEnScope(seccion, /Tiempo en vivienda/i, '1').catch(() => { });
    await llenarInputEnScope(seccion, /Cantidad de dependientes/i, '1').catch(() => { });
    await llenarInputEnScope(seccion, /Cantidad de hijos/i, '1').catch(() => { });
    return true;
}

async function asegurarDireccionEnCorrespondencia(page: Page) {
    const titulo = page.getByText(/Correspondencia/i).first();
    const visibleTitulo = await titulo.isVisible().catch(() => false);
    if (!visibleTitulo) return false;

    const seccion = titulo.locator('xpath=ancestor::*[self::fieldset or self::div][1]').first();
    await seccion.scrollIntoViewIfNeeded().catch(() => { });

    await asegurarDropdownIndex0SiVacioEnScope(page, seccion, /Tipo de env(?:i|\u00ed)o/i, { required: false }).catch(() => { });

    let direccionVisibleAlgunaVez = false;
    for (let intento = 1; intento <= 6; intento++) {
        const labelDireccion = seccion.locator('label').filter({ hasText: /Direcci(?:o|\u00f3)n/i }).first();
        const direccionVisible = await labelDireccion.isVisible().catch(() => false);
        if (!direccionVisible) {
            await page.waitForTimeout(FAST_UI ? 120 : 350);
            continue;
        }

        direccionVisibleAlgunaVez = true;
        await clickReintentarListaSiVisible(page, /Direcci(?:o|\u00f3)n/i, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
        await seleccionarDropdownIndexEnScope(page, seccion, /Direcci(?:o|\u00f3)n/i, 0, {
            maxIntentos: FAST_UI ? 4 : 6,
            timeoutMs: LIST_LABEL_TIMEOUT,
            panelTimeoutMs: LIST_PANEL_TIMEOUT,
        }).catch(() => { });

        const valorDireccion = await leerValorDropdownEnScope(seccion, /Direcci(?:o|\u00f3)n/i).catch(() => '');
        const vacio = !valorDireccion || /^seleccione/i.test(valorDireccion) || /Reintentar/i.test(valorDireccion);
        if (!vacio) return true;

        await page.waitForTimeout(FAST_UI ? 120 : 350);
    }

    if (direccionVisibleAlgunaVez) {
        throw new Error("[CRITICO] En Correspondencia, el campo 'Dirección' quedo vacio.");
    }
    return false;
}

async function abrirModalDireccionCasa(page: Page) {
    const btnAnadirDireccion = page.getByRole('button', { name: /A(?:\u00f1|n)adir direcci(?:o|\u00f3)n/i }).first();
    const modalDireccion = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Direcci(?:o|\u00f3)n|Calle|Provincia|Municipio/i })
        .first();
    let ultimoTextoCandidato = '';
    let ultimosCandidatos = 0;

    const modalVisible = async () => {
        const visible = await modalDireccion.waitFor({ state: 'visible', timeout: ADDRESS_MODAL_OPEN_TIMEOUT }).then(() => true).catch(() => false);
        if (!visible) return false;
        const senales = await Promise.all([
            modalDireccion.getByText(/Direcci(?:o|\u00f3)n/i).first().isVisible().catch(() => false),
            modalDireccion.locator('label').filter({ hasText: /^Tipo$/i }).first().isVisible().catch(() => false),
            modalDireccion.locator('label').filter({ hasText: /^Pa(?:i|í)s$/i }).first().isVisible().catch(() => false),
            modalDireccion.locator('label').filter({ hasText: /Calle|Avenida|Autopista/i }).first().isVisible().catch(() => false),
            modalDireccion.getByRole('button', { name: /Aceptar/i }).first().isVisible().catch(() => false),
            modalDireccion.getByRole('button', { name: /Cancelar/i }).first().isVisible().catch(() => false),
        ]);
        return senales.filter(Boolean).length >= 2;
    };

    const clickAbrir = async (target: Locator) => {
        const visible = await target.isVisible().catch(() => false);
        if (!visible) return false;
        await target.scrollIntoViewIfNeeded().catch(() => { });
        let clicked = await target.click({ timeout: 1800 }).then(() => true).catch(() => false);
        if (!clicked) clicked = await target.click({ force: true, timeout: 1800 }).then(() => true).catch(() => false);
        if (!clicked) {
            await target.focus().catch(() => { });
            await page.keyboard.press('Enter').catch(() => { });
            clicked = await modalVisible();
            if (!clicked) {
                await page.keyboard.press('Space').catch(() => { });
                clicked = await modalVisible();
            }
        }
        return clicked;
    };

    for (let intento = 1; intento <= 4; intento++) {
        console.log('[DireccionCasa] Buscando bloque Dirección');
        await btnAnadirDireccion.waitFor({ state: 'visible', timeout: ADDRESS_MODAL_OPEN_BTN_TIMEOUT });
        await btnAnadirDireccion.scrollIntoViewIfNeeded().catch(() => { });
        await btnAnadirDireccion.click({ timeout: 1800 }).catch(async () => {
            await btnAnadirDireccion.click({ force: true, timeout: 1800 }).catch(() => { });
        });

        console.log('[DireccionCasa] Buscando opción Casa');
        const bloqueDireccion = btnAnadirDireccion
            .locator('xpath=ancestor::*[(self::fieldset or self::section or self::div) and .//*[contains(normalize-space(.),"Direcci")]][1]')
            .first();
        const opcionCasa = page.locator(
            'li[aria-label="Casa"]:visible, [role="menuitem"]:has-text("Casa"):visible, .p-menuitem:has-text("Casa"):visible, .p-menuitem-link:has-text("Casa"):visible, button:has-text("Casa"):visible'
        ).first();
        const contenedorCasa = page
            .locator('xpath=(//*[normalize-space(.)="Casa" or contains(normalize-space(.),"Casa")]/ancestor::*[(self::li or self::tr or self::div) and (.//button or .//*[@role="button"] or self::*[@role="menuitem"])][1])[1]')
            .first();
        const contenedorCasaEnDireccion = bloqueDireccion
            .locator('xpath=.//*[normalize-space(.)="Casa" or contains(normalize-space(.),"Casa")]/ancestor::*[(self::li or self::tr or self::div) and (.//button or .//*[@role="button"] or self::*[@role="menuitem"])][1]')
            .first();
        const candidatos: Locator[] = [
            opcionCasa,
            contenedorCasaEnDireccion,
            contenedorCasa.locator('button:visible, [role="button"]:visible, .p-button:visible, i:visible, [class*="pi-"]:visible, [class*="ph-"]:visible').first(),
            contenedorCasa,
        ];

        let candidatosVisibles = 0;
        for (const candidato of candidatos) {
            if (await candidato.isVisible().catch(() => false)) candidatosVisibles++;
        }
        ultimosCandidatos = candidatosVisibles;
        ultimoTextoCandidato = ((await contenedorCasa.innerText().catch(() => '')) || '').replace(/\s+/g, ' ').trim().slice(0, 240);
        console.log(`[DireccionCasa] Candidatos abrir modal encontrados=${candidatosVisibles}`);

        for (const candidato of candidatos) {
            if (!await candidato.isVisible().catch(() => false)) continue;
            console.log('[DireccionCasa] Click en abrir modal Dirección Casa');
            await clickAbrir(candidato);
            const visibleModal = await modalVisible();
            console.log(`[DireccionCasa] Modal Dirección Casa visible=${visibleModal}`);
            if (visibleModal) return modalDireccion;
        }

        const visibleModal = await modalVisible();
        console.log(`[DireccionCasa] Modal Dirección Casa visible=${visibleModal}`);
        if (visibleModal) return modalDireccion;

        await page.waitForTimeout(FAST_UI ? 120 : 250);
    }

    const urlActual = page.url();
    const direccionVisible = await page.getByText(/Direcci(?:o|ó)n/i).first().isVisible().catch(() => false);
    const casaVisible = await page.getByText(/Casa/i).first().isVisible().catch(() => false);
    console.log(`[DireccionCasa][Diag] url=${urlActual} direccionVisible=${direccionVisible} casaVisible=${casaVisible} candidatos=${ultimosCandidatos} texto='${ultimoTextoCandidato}'`);
    throw new Error(`[DireccionCasa][CRITICO] No se pudo abrir modal Dirección Casa. candidatos=${ultimosCandidatos} texto='${ultimoTextoCandidato}'`);
}

async function completarDireccionPostProducto(page: Page) {
    const btnAnadirDireccion = page.getByRole('button', { name: /A(?:\u00f1|n)adir direcci(?:o|\u00f3)n/i }).first();
    const visible = await btnAnadirDireccion.isVisible().catch(() => false);
    if (!visible) return false;

    // Si ya existe una direccion registrada dentro de la seccion Direccion, no agregar otra
    const seccionDireccion = page.locator(
        'xpath=//fieldset[.//*[contains(normalize-space(.),"Direcci") and (self::legend or self::label or self::span or self::h2 or self::h3)]]'
    ).first();
    const yaHayDireccion = await seccionDireccion.getByText(/Predeterminado/i).first().isVisible().catch(() => false);
    if (yaHayDireccion) {
        console.log('[Direccion] Ya existe una direccion registrada en seccion Direccion. Omitiendo agregar nueva.');
        return false;
    }

    const modalDireccion = await abrirModalDireccionCasa(page);
    console.log('[Direccion] Modal Casa abierto');

    const valorInvalido = (valor: string) => !valor || /^seleccione/i.test(valor) || /Reintentar/i.test(valor) || /^[-\s]+$/.test(valor.trim());

    const resolverCampoDireccion = async (label: RegExp) => {
        const labelLoc = modalDireccion.locator('label').filter({ hasText: label }).first();
        await labelLoc.waitFor({ state: 'visible', timeout: 1800 });
        let campo = labelLoc
            .locator('xpath=ancestor::div[(contains(@class,"field") or contains(@class,"p-field") or contains(@class,"col") or contains(@class,"form")) and (.//input or .//textarea or .//*[contains(@class,"p-dropdown") or @data-pc-name="dropdown" or @role="combobox"])][1]')
            .first();
        if (!(await campo.isVisible().catch(() => false))) {
            campo = labelLoc
                .locator('xpath=ancestor::div[.//input or .//textarea or .//*[contains(@class,"p-dropdown") or @data-pc-name="dropdown" or @role="combobox"]][1]')
                .first();
        }
        return { labelLoc, campo };
    };

    const leerValorRapido = async (label: RegExp) => {
        const campoResuelto = await resolverCampoDireccion(label).catch(() => null);
        if (!campoResuelto) return '';
        const { campo } = campoResuelto;

        const labelDropdown = campo.locator('.p-dropdown-label:visible, [data-pc-section="label"]:visible').first();
        const textoDropdown = ((await labelDropdown.innerText().catch(() => '')) || '').trim();
        if (textoDropdown) return textoDropdown;

        const input = campo.locator('input:visible, textarea:visible').first();
        const valorInput = ((await input.inputValue().catch(() => '')) || '').trim();
        if (valorInput) return valorInput;
        return '';
    };

    const llenarTextoRapido = async (label: RegExp, valor: string, required = true) => {
        const campoResuelto = await resolverCampoDireccion(label).catch(() => null);
        const labelVisible = !!campoResuelto;
        if (!labelVisible) {
            if (required) throw new Error(`[Direccion] No se encontro campo de texto '${label}'.`);
            return false;
        }

        const input = campoResuelto.campo.locator('input:visible, textarea:visible').first();
        await input.waitFor({ state: 'visible', timeout: 1500 });
        await input.fill(valor, { timeout: 1200 }).catch(async () => {
            await input.evaluate((el, v) => {
                const inputEl = el as HTMLInputElement | HTMLTextAreaElement;
                inputEl.value = String(v);
                inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            }, valor);
        });
        await input.blur().catch(() => { });
        return true;
    };

    const seleccionarComboRapido = async (label: RegExp, nombre: string, required = true, preferida?: RegExp) => {
        const valorActual = await leerValorRapido(label);
        if (!valorInvalido(valorActual)) {
            console.log(`[Direccion] ${nombre} ya tiene valor, se omite`);
            return { ok: true, valor: valorActual };
        }

        console.log(`[Direccion] Seleccionando ${nombre}`);
        const { labelLoc, campo } = await resolverCampoDireccion(label);
        const dropdown = campo.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible, [role="combobox"]:visible').first();
        const trigger = dropdown.locator('.p-dropdown-trigger, [data-pc-section="trigger"]').first();
        const input = campo.locator('input:visible').first();
        const boton = campo.locator('button:visible').first();
        if (await dropdown.isVisible().catch(() => false)) {
            await dropdown.click({ force: true, timeout: 1200 }).catch(() => { });
        } else if (await trigger.isVisible().catch(() => false)) {
            await trigger.click({ force: true, timeout: 1200 }).catch(() => { });
        } else if (await input.isVisible().catch(() => false)) {
            await input.click({ force: true, timeout: 1200 }).catch(() => { });
        } else if (await boton.isVisible().catch(() => false)) {
            await boton.click({ force: true, timeout: 1200 }).catch(() => { });
        } else {
            await labelLoc.click({ force: true }).catch(() => { });
        }

        let panel = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible, div.ui-select-dropdown.open').last();
        let panelVisible = await panel.waitFor({ state: 'visible', timeout: 2000 }).then(() => true).catch(() => false);
        if (!panelVisible) {
            await trigger.click({ force: true, timeout: 800 }).catch(() => { });
            await input.click({ force: true, timeout: 800 }).catch(() => { });
            await page.keyboard.press('ArrowDown').catch(() => { });
            panel = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible, div.ui-select-dropdown.open').last();
            panelVisible = await panel.waitFor({ state: 'visible', timeout: 1000 }).then(() => true).catch(() => false);
        }

        if (!panelVisible) {
            await page.keyboard.press('Enter').catch(() => { });
            await page.waitForTimeout(FAST_UI ? 150 : 250);
            const valorTrasTeclado = await leerValorRapido(label);
            const ok = !valorInvalido(valorTrasTeclado);
            if (ok) {
                console.log(`[Direccion] ${nombre} seleccionado rapido=true valor='${valorTrasTeclado}'`);
                return { ok: true, valor: valorTrasTeclado };
            }
            throw new Error(`[Direccion] No se abrio dropdown de ${nombre}. valorActual='${valorActual}'`);
        }

        const opciones = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
        const total = await opciones.count().catch(() => 0);
        const participio = /Regi(?:o|ó)n|Provincia|Localidad/i.test(nombre) ? 'seleccionada' : 'seleccionado';
        if (preferida) {
            const opcionPreferida = opciones.filter({ hasText: preferida }).first();
            if (await opcionPreferida.isVisible().catch(() => false)) {
                await opcionPreferida.click({ force: true, timeout: 1200 }).catch(() => { });
                await page.waitForTimeout(FAST_UI ? 180 : 350);
                const valorFinalPreferido = await leerValorRapido(label);
                const okPreferido = !valorInvalido(valorFinalPreferido);
                console.log(`[Direccion] ${nombre} ${participio} rapido=${okPreferido} valor='${valorFinalPreferido}'`);
                if (!okPreferido && required) throw new Error(`[Direccion] ${nombre} quedo vacio tras seleccionar opcion preferida.`);
                return { ok: okPreferido, valor: valorFinalPreferido };
            }
        }

        for (let i = 0; i < total; i++) {
            const opcion = opciones.nth(i);
            if (!(await opcion.isVisible().catch(() => false))) continue;
            const texto = ((await opcion.innerText().catch(() => '')) || '').trim();
            if (valorInvalido(texto)) continue;
            await opcion.click({ force: true, timeout: 1200 }).catch(() => { });
            await page.waitForTimeout(FAST_UI ? 120 : 250);
            const valorFinal = await leerValorRapido(label);
            const ok = !valorInvalido(valorFinal);
            console.log(`[Direccion] ${nombre} ${participio} rapido=${ok} valor='${valorFinal}'`);
            if (!ok && required) throw new Error(`[Direccion] ${nombre} quedo vacio tras seleccionar '${texto}'.`);
            return { ok, valor: valorFinal };
        }

        throw new Error(`[Direccion] No hay opciones validas para ${nombre}. total=${total}`);
    };

    await llenarTextoRapido(/Calle.*Avenida.*Autopista/i, randomTexto('Calle'));
    await llenarTextoRapido(/Nombre edificio/i, randomTexto('Edificio'), false).catch(() => false);
    await llenarTextoRapido(/N[uú]mero casa.*edificio/i, String(randomInt(1, 999)));
    await llenarTextoRapido(/N[uú]mero apartamento/i, String(randomInt(1, 80)));

    console.log('[Direccion] Tipo se omite; no es necesario seleccionarlo');
    console.log('[Direccion] País listo');

    // Usar estrategia de Datos Laborales que funcionan correctamente
    console.log('[Direccion] Seleccionando Región');
    await seleccionarDropdownEnScopePorTexto(page, modalDireccion, /^Regi(?:o|ó)n$/i, /Distrito Nacional/i, 0)
        .catch(async () => {
            console.log('[Direccion][Diag] Región: Distrito Nacional no encontrado, usando primer valor');
            return seleccionarDropdownIndexEnScope(page, modalDireccion, /^Regi(?:o|ó)n$/i, 0, { maxIntentos: 3, timeoutMs: 2500, panelTimeoutMs: 3000 });
        });
    let regionValor = await leerValorRapido(/^Regi(?:o|ó)n$/i);
    console.log(`[Direccion] Región seleccionada rapido=true valor='${regionValor}'`);
    if (valorInvalido(regionValor)) throw new Error(`[Direccion][CRITICO] No se pudo seleccionar Región`);
    await page.waitForTimeout(FAST_UI ? 300 : 700);

    console.log('[Direccion] Seleccionando Provincia');
    await seleccionarDropdownEnScopePorTexto(page, modalDireccion, /^Provincia$/i, /Distrito Nacional/i, 0)
        .catch(async () => {
            console.log('[Direccion][Diag] Provincia: Distrito Nacional no encontrado, usando primer valor');
            return seleccionarDropdownIndexEnScope(page, modalDireccion, /^Provincia$/i, 0, { maxIntentos: 3, timeoutMs: 2500, panelTimeoutMs: 3000 });
        });
    let provinciaValor = await leerValorRapido(/^Provincia$/i);
    console.log(`[Direccion] Provincia seleccionada rapido=true valor='${provinciaValor}'`);
    if (valorInvalido(provinciaValor)) throw new Error(`[Direccion][CRITICO] No se pudo seleccionar Provincia`);
    await page.waitForTimeout(FAST_UI ? 300 : 700);

    console.log('[Direccion] Seleccionando Municipio');
    await seleccionarDropdownDependienteConEspera(page, modalDireccion, /^Municipio$/i, 0, {
        maxIntentos: FAST_UI ? 8 : 12,
        timeoutMs: FAST_UI ? 4200 : 7000,
        panelTimeoutMs: FAST_UI ? 4800 : 9000,
        esperaEntreIntentosMs: FAST_UI ? 300 : 900,
    });
    let municipioValor = await leerValorRapido(/^Municipio$/i);
    console.log(`[Direccion] Municipio seleccionado rapido=true valor='${municipioValor}'`);
    if (valorInvalido(municipioValor)) throw new Error(`[Direccion][CRITICO] No se pudo seleccionar Municipio`);
    await page.waitForTimeout(FAST_UI ? 300 : 700);

    console.log('[Direccion] Seleccionando Localidad');
    await seleccionarDropdownDependienteConEspera(page, modalDireccion, /^Localidad$/i, 0, {
        maxIntentos: FAST_UI ? 8 : 12,
        timeoutMs: FAST_UI ? 4200 : 7000,
        panelTimeoutMs: FAST_UI ? 4800 : 9000,
        esperaEntreIntentosMs: FAST_UI ? 300 : 900,
    });
    let localidadValor = await leerValorRapido(/^Localidad$/i);
    console.log(`[Direccion] Localidad seleccionada rapido=true valor='${localidadValor}'`);
    if (valorInvalido(localidadValor)) throw new Error(`[Direccion][CRITICO] No se pudo seleccionar Localidad`);
    await page.waitForTimeout(FAST_UI ? 300 : 700);

    console.log('[Direccion] Seleccionando Sector');
    await seleccionarDropdownDependienteConEspera(page, modalDireccion, /^Sector$/i, 0, {
        maxIntentos: FAST_UI ? 7 : 10,
        timeoutMs: FAST_UI ? 3800 : 6500,
        panelTimeoutMs: FAST_UI ? 4400 : 8000,
        esperaEntreIntentosMs: FAST_UI ? 260 : 800,
    });
    let sectorValor = await leerValorRapido(/^Sector$/i);
    console.log(`[Direccion] Sector seleccionado rapido=true valor='${sectorValor}'`);
    if (valorInvalido(sectorValor)) throw new Error(`[Direccion][CRITICO] No se pudo seleccionar Sector`);
    await page.waitForTimeout(FAST_UI ? 300 : 700);

    console.log('[Direccion] Rellenando Referencia');
    console.log('[Direccion][Referencia] Buscando campo Referencia');

    let referenciaValor = '';
    {
        // Buscar campo por label dentro del modal
        const labelRef = modalDireccion.locator('label').filter({ hasText: /^Referencia$/i }).first();
        const labelVisible = await labelRef.isVisible({ timeout: 1500 }).catch(() => false);
        console.log(`[Direccion][Referencia] Campo por label encontrado=${labelVisible}`);

        let campoRef: Locator | null = null;

        if (labelVisible) {
            // Subir al ancestro contenedor del campo
            const contenedor = labelRef
                .locator('xpath=ancestor::div[.//input or .//textarea or .//*[@role="textbox"] or .//*[@contenteditable="true"]][1]')
                .first();
            const contenedorVisible = await contenedor.isVisible().catch(() => false);
            if (contenedorVisible) {
                campoRef = contenedor.locator('textarea, input, [role="textbox"], [contenteditable="true"]').first();
            }
        }

        // Fallback global dentro del modal si no encontró por label
        if (!campoRef) {
            const globalRef = modalDireccion.locator(
                'textarea:visible, [role="textbox"]:visible'
            ).first();
            const globalVisible = await globalRef.isVisible({ timeout: 1500 }).catch(() => false);
            if (globalVisible) campoRef = globalRef;
        }
        if (!campoRef) {
            // Buscar input visible con placeholder/name/aria-label que mencione referencia
            const inputs = modalDireccion.locator('input:visible');
            const totalInputs = await inputs.count().catch(() => 0);
            for (let i = 0; i < totalInputs; i++) {
                const inp = inputs.nth(i);
                const ph = (await inp.getAttribute('placeholder').catch(() => '')) || '';
                const nm = (await inp.getAttribute('name').catch(() => '')) || '';
                const al = (await inp.getAttribute('aria-label').catch(() => '')) || '';
                if (/referencia/i.test(ph + nm + al)) { campoRef = inp; break; }
            }
        }

        const campoVisible = await campoRef?.isVisible().catch(() => false) ?? false;
        console.log(`[Direccion][Referencia] Campo visible=${campoVisible}`);

        if (!campoRef || !campoVisible) {
            throw new Error('[Direccion][CRITICO] No se pudo llenar Referencia: campo no encontrado en modal');
        }

        console.log('[Direccion][Referencia] Fill directo');
        await campoRef.fill('Referencia automatizada', { timeout: 2000 }).catch(async () => {
            await campoRef!.evaluate((el, v) => {
                if ('value' in el) {
                    (el as HTMLInputElement).value = v;
                } else {
                    (el as HTMLElement).innerText = v;
                }
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }, 'Referencia automatizada');
        });
        await campoRef.blur().catch(() => {});

        // Leer valor final
        const esInput = await campoRef.evaluate(el => el.tagName === 'INPUT' || el.tagName === 'TEXTAREA').catch(() => true);
        referenciaValor = esInput
            ? ((await campoRef.inputValue().catch(() => '')) || '').trim()
            : ((await campoRef.innerText().catch(() => '')) || '').trim();
    }

    console.log(`[Direccion] Referencia llenada valor='${referenciaValor}'`);

    console.log('[Direccion] Validando campos antes de Aceptar');
    if (valorInvalido(regionValor) || valorInvalido(provinciaValor) || valorInvalido(municipioValor) || valorInvalido(localidadValor) || valorInvalido(sectorValor) || valorInvalido(referenciaValor)) {
        throw new Error(`[Direccion][CRITICO] Validacion incompleta. region='${regionValor}' provincia='${provinciaValor}' municipio='${municipioValor}' localidad='${localidadValor}' sector='${sectorValor}' referencia='${referenciaValor}'`);
    }

    const btnAceptar = modalDireccion.getByRole('button', { name: /^Aceptar$/i }).first();
    await btnAceptar.waitFor({ state: 'visible', timeout: 12000 });

    console.log('[Direccion] Click en Aceptar');
    let modalCerrado = false;
    for (let intento = 1; intento <= 3 && !modalCerrado; intento++) {
        await btnAceptar.click({ force: true }).catch(() => { });
        modalCerrado = await modalDireccion.waitFor({ state: 'hidden', timeout: 3000 }).then(() => true).catch(() => false);
        if (!modalCerrado && intento < 3) await page.waitForTimeout(300);
    }
    if (!modalCerrado) {
        throw new Error('[Direccion] El modal de direccion no cerro tras 3 intentos de Aceptar.');
    }
    console.log('[Direccion] Modal Dirección cerrado');
    return true;
}

async function seleccionarNoEnTodosLosSelectbuttonDeScope(page: Page, scope: Locator) {
    const grupos = scope.locator('.p-selectbutton, [data-pc-name="selectbutton"], [role="radiogroup"]');
    const total = await grupos.count().catch(() => 0);
    let cambios = 0;

    for (let i = 0; i < total; i++) {
        const grupo = grupos.nth(i);
        const opcionNo = grupo
            .locator('[role="radio"][aria-label*="No" i], button:has-text("No"), [data-pc-section="button"]:has-text("No")')
            .first();
        const visibleNo = await opcionNo.isVisible().catch(() => false);
        if (!visibleNo) continue;

        await opcionNo.scrollIntoViewIfNeeded().catch(() => { });
        await opcionNo.click({ force: true }).catch(() => { });
        await page.waitForTimeout(100);
        cambios++;
    }

    return cambios > 0;
}

async function completarReferenciaPersonalPostProducto(page: Page) {
    const btnAnadirReferencia = page.getByRole('button', { name: /A(?:\u00f1|n)adir referencia/i }).first();
    const visible = await btnAnadirReferencia.isVisible().catch(() => false);
    if (!visible) return false;

    await btnAnadirReferencia.scrollIntoViewIfNeeded().catch(() => { });
    await btnAnadirReferencia.click({ force: true });

    const modalReferencia = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Referencia personal/i })
        .first();
    await modalReferencia.waitFor({ state: 'visible', timeout: 12000 });

    await llenarInputEnScope(modalReferencia, /Nombres y apellidos/i, randomTexto('Referencia'));
    await llenarInputEnScope(modalReferencia, /Tel[eé]fono/i, randomTelefono());
    await llenarInputEnScope(modalReferencia, /Relaci[oó]n/i, 'Amigo');

    const btnAceptar = modalReferencia.getByRole('button', { name: /^Aceptar$/i }).first();
    await btnAceptar.waitFor({ state: 'visible', timeout: 12000 });
    await btnAceptar.click({ force: true });
    await modalReferencia.waitFor({ state: 'hidden', timeout: 15000 });
    return true;
}

async function asegurarPantallaPostProductoAntesDeTaller(page: Page) {
    const esVisibleBloquePostProducto = async () => {
        const nivel = await page.getByText(/Nivel de estudio/i).first().isVisible().catch(() => false);
        const reside = await page.getByText(/Reside actualmente en la Rep(?:u|\u00fa)blica Dominicana/i).first().isVisible().catch(() => false);
        const addDir = await page.getByRole('button', { name: /A(?:\u00f1|n)adir direcci(?:o|\u00f3)n/i }).first().isVisible().catch(() => false);
        return nivel || reside || addDir;
    };

    console.log('[PostProducto] asegurarPantallaPostProducto - verificando visibilidad inicial...');
    if (await esVisibleBloquePostProducto()) {
        console.log('[PostProducto] Bloque post-producto ya visible.');
        return true;
    }

    // Si el elemento esta en el DOM (aunque colapsado/fuera del viewport), hacer scroll hasta el
    const nivelEnDom = page.getByText(/Nivel de estudio/i).first();
    const nivelAttached = await nivelEnDom.waitFor({ state: 'attached', timeout: 2000 }).then(() => true).catch(() => false);
    console.log(`[PostProducto] nivelEnDom attached=${nivelAttached}`);

    if (nivelAttached) {
        // Forzar scroll al elemento independientemente del viewport
        await nivelEnDom.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => { });
        await page.waitForTimeout(FAST_UI ? 350 : 700);
        if (await esVisibleBloquePostProducto()) {
            console.log('[PostProducto] Bloque visible tras scrollIntoViewIfNeeded.');
            return true;
        }
        // Si la seccion esta colapsada (acordeon), intentar expandirla
        const seccionColapsada = page.locator(
            'p-accordiontab:has-text("Nivel de estudio"), .p-accordion-tab:has-text("Nivel de estudio"), details:has-text("Nivel de estudio")'
        ).first();
        const haySeccionColapsada = await seccionColapsada.isVisible().catch(() => false);
        if (haySeccionColapsada) {
            console.log('[PostProducto] Seccion colapsada detectada. Intentando expandir...');
            const header = seccionColapsada.locator('.p-accordion-header, summary').first();
            await header.click({ force: true }).catch(() => { });
            await page.waitForTimeout(FAST_UI ? 350 : 700);
        }
    }

    // Scroll al top absoluto
    console.log('[PostProducto] Scrolleando al top absoluto...');
    await page.evaluate(() => window.scrollTo(0, 0)).catch(() => { });
    await page.waitForTimeout(FAST_UI ? 400 : 900);
    if (await esVisibleBloquePostProducto()) {
        console.log('[PostProducto] Bloque post-producto visible tras scroll al top.');
        return true;
    }

    console.log(`[PostProducto] Scroll unico hacia seccion datos cliente... attached=${nivelAttached}`);
    const step2 = page.locator(
        'li.p-steps-item:has-text("2"), [data-pc-section="menuitem"]:has-text("2"), .p-steps-item:has(.p-steps-number:has-text("2"))'
    ).first();
    if (await step2.isVisible().catch(() => false)) {
        await step2.click({ force: true }).catch(() => { });
    }
    await page.evaluate(() => window.scrollTo(0, 0)).catch(() => { });
    await page.mouse.wheel(0, -3000).catch(() => { });
    await page.waitForTimeout(FAST_UI ? 450 : 1200);

    if (await esVisibleBloquePostProducto()) {
        console.log('[PostProducto] Bloque post-producto visible tras scroll unico.');
        return true;
    }

    console.log('[PostProducto][WARN] No se pudo detectar pantalla post-producto (Nivel de estudio / Direccion).');
    return false;
}

async function completarInformacionClientePostProductoAntesDeTaller(
    page: Page,
    options?: { required?: boolean }
) {
    const required = options?.required ?? false;
    const gatillo = await asegurarPantallaPostProductoAntesDeTaller(page);
    if (!gatillo) {
        if (required) {
            throw new Error("[CRITICO] No se detecto la pantalla post-producto (Nivel de estudio / Direccion) antes de ir a Taller.");
        }
        return false;
    }

    console.log('[PostProducto] Completando informacion del cliente antes de Taller...');
    const inicioPostProducto = Date.now();

    const inicioNivel = Date.now();
    await clickReintentarListaSiVisible(page, 'Nivel de estudio', LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
    await asegurarDropdownConIndex0SiVacio(page, /Nivel de estudio/i);
    console.log(`[PostProducto][Tiempo] Nivel de estudio listo en ${Date.now() - inicioNivel}ms`);

    const inicioPreguntas = Date.now();
    await seleccionarRespuestaPregunta(
        page,
        /Reside actualmente en la Rep(?:u|\u00fa)blica Dominicana/i,
        'Si',
        { required: false }
    );
    await seleccionarRespuestaPregunta(
        page,
        /Posee otras identificaciones o nacionalidades/i,
        'No',
        { required: false }
    );
    console.log(`[PostProducto][Tiempo] Preguntas base listas en ${Date.now() - inicioPreguntas}ms`);

    const inicioDireccion = Date.now();
    await completarDireccionPostProducto(page);
    console.log(`[PostProducto][Tiempo] Direccion lista en ${Date.now() - inicioDireccion}ms`);

    const correspondenciaVisible = await page.getByText(/Correspondencia/i).first().isVisible().catch(() => false);
    if (correspondenciaVisible) {
        const inicioCorrespondencia = Date.now();
        await completarDropdownsVaciosIndex0EnSeccion(page, /Correspondencia/i);
        await asegurarDireccionEnCorrespondencia(page);
        console.log(`[PostProducto][Tiempo] Correspondencia lista en ${Date.now() - inicioCorrespondencia}ms`);
    }
    const otrasInfoVisible = await page.getByText(/Otras informaciones/i).first().isVisible().catch(() => false);
    if (otrasInfoVisible) {
        const inicioOtrasInfo = Date.now();
        await completarDropdownsVaciosIndex0EnSeccion(page, /Otras informaciones/i);
        await completarCamposTextoOtrasInformaciones(page);
        console.log(`[PostProducto][Tiempo] Otras informaciones listas en ${Date.now() - inicioOtrasInfo}ms`);
    }

    const pepTitulo = page.getByText(/Informaci(?:o|\u00f3)n PEP|PEP/i).first();
    const pepVisible = await pepTitulo.isVisible().catch(() => false);
    if (pepVisible) {
        const inicioPep = Date.now();
        const seccionPep = pepTitulo.locator('xpath=ancestor::*[self::fieldset or self::div][1]').first();
        await seleccionarNoEnTodosLosSelectbuttonDeScope(page, seccionPep);
        console.log(`[PostProducto][Tiempo] PEP listo en ${Date.now() - inicioPep}ms`);
    }

    const fatcaTitulo = page.getByText(/Informaci(?:o|\u00f3)n FATCA|FATCA/i).first();
    const fatcaVisible = await fatcaTitulo.isVisible().catch(() => false);
    if (fatcaVisible) {
        const inicioFatca = Date.now();
        const seccionFatca = fatcaTitulo.locator('xpath=ancestor::*[self::fieldset or self::div][1]').first();
        await seleccionarNoEnTodosLosSelectbuttonDeScope(page, seccionFatca);
        console.log(`[PostProducto][Tiempo] FATCA listo en ${Date.now() - inicioFatca}ms`);
    }

    const inicioReferencia = Date.now();
    await completarReferenciaPersonalPostProducto(page);
    console.log(`[PostProducto][Tiempo] Referencia personal lista en ${Date.now() - inicioReferencia}ms`);

    console.log(`[PostProducto] Informacion del cliente completada antes de Taller. total=${Date.now() - inicioPostProducto}ms`);
    return true;
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

        const estaEnSolicitudMultiproducto = /\/requests\/create\/multiproduct/i.test(page.url());
        if (!estaEnSolicitudMultiproducto) {
            const forzadoDirecto = await asegurarPortalEnMultiproducto(page, `Flujo][intento=${intento}`);
            if (forzadoDirecto) {
                console.log(`[Flujo][intento=${intento}] URL tras forzado directo: ${page.url()} | enMultiproducto=true`);
            }

            if (!/\/requests\/create\/multiproduct/i.test(page.url())) {
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

        await clickReintentarListaSiVisible(page, 'Sexo', LIST_RETRY_CLICK_TIMEOUT);
        await clickReintentarListaSiVisible(page, /Pa(?:i|\u00ed)s de nacimiento/i, LIST_RETRY_CLICK_TIMEOUT);
        await clickReintentarListaSiVisible(page, 'Estado civil', LIST_RETRY_CLICK_TIMEOUT);
        await asegurarOficialPromotor(page);

        await asegurarIdentificacionHabilitada(page, { timeoutCampoMs: 60000 });
        const inputNumeroRefuerzo = await resolverInputNumeroIdentificacion(page, 20000);
        let habilitadoNumero = await inputNumeroRefuerzo.isEnabled({ timeout: 5000 }).catch(() => false);
        if (!habilitadoNumero) {
            await esperarYClickReintentarPaisIdentificacion(page, 12000);
            const clicPais = await clickReintentarListaSiVisible(page, 'País de identificación', LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
            if (!clicPais) {
                await clickReintentarListaSiVisible(page, 'Tipo de identificación', LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
            }
            habilitadoNumero = await inputNumeroRefuerzo.isEnabled({ timeout: 5000 }).catch(() => false);
            if (!habilitadoNumero) {
                throw new Error("Input 'Numero de identificacion' sigue deshabilitado tras refuerzo.");
            }
        }

        const inputNumero = await resolverInputNumeroIdentificacion(page, 20000);
        await inputNumero.click({ force: true });
        await inputNumero.fill(registro.identificacion);

        await page.keyboard.press('Enter');
        await esperarFinConsultaSolicitante(page, 120000);
        await asegurarOficialPromotor(page);

        const inputFechaExp = await resolverInputFechaExpiracion(page, 12000);
        const fechaExpVisible = await inputFechaExp.isVisible().catch(() => false);
        if (fechaExpVisible) {
            let valorFechaExp = (await inputFechaExp.inputValue().catch(() => '')).trim();
            if (!valorFechaExp) {
                await setFechaExpiracionRobusta(page, '09-01-2038').catch(() => { });
            }
        }

        await clickDepurarRobusto(page, 'primero');

        const mpnCasoActivo = await extraerCasoActivoMpn(page);
        if (mpnCasoActivo) {
            console.log(`[CasoActivo] Detectado ${mpnCasoActivo}. Cancelando en Bizagi...`);
            await cancelarCasoEnBizagiDesdePortal(page, mpnCasoActivo, { password: BIZAGI_PASSWORD });
            if (intento >= maxIntentosCasoActivo) {
                throw new Error(`[CRITICO] Se cancelÃ³ ${mpnCasoActivo} pero el caso activo persiste para '${registro.identificacion}'.`);
            }
            console.log(`[CasoActivo] ${mpnCasoActivo} cancelado. Reintentando con la misma cÃ©dula (${registro.identificacion})...`);
            await page.waitForTimeout(500);
            continue;
        }

        const msgFechaInvalida = page.getByText(/La fecha m(?:i|\u00ed)nima permitida es|La fecha m(?:a|\u00e1)xima permitida es/i).first();
        const hayMsgFechaInvalida = await msgFechaInvalida.isVisible().catch(() => false);
        if (hayMsgFechaInvalida) {
            const fechaCorregida = await setFechaExpiracionRobusta(page, '09-01-2038').catch(() => false);

            if (!fechaCorregida) {
                console.log('[ClienteNuevo][WARN] La fecha de expiracion no se pudo corregir tras el primer Depurar. Se omite un segundo Depurar para evitar reprocesar el mismo error.');
            } else {
                await clickDepurarRobusto(page, 'fecha invalida');

            const mpnCasoActivo2 = await extraerCasoActivoMpn(page);
            if (mpnCasoActivo2) {
                console.log(`[CasoActivo] Detectado ${mpnCasoActivo2} en segundo depurar. Cancelando en Bizagi...`);
                await cancelarCasoEnBizagiDesdePortal(page, mpnCasoActivo2, { password: BIZAGI_PASSWORD });
                if (intento >= maxIntentosCasoActivo) {
                    throw new Error(`[CRITICO] Se cancelÃ³ ${mpnCasoActivo2} pero el caso activo persiste para '${registro.identificacion}'.`);
                }
                console.log(`[CasoActivo] ${mpnCasoActivo2} cancelado. Reintentando con la misma cÃ©dula (${registro.identificacion})...`);
                await page.waitForTimeout(500);
                continue;
            }
            }
        }

        const estaEnPasoDatosCliente = await asegurarSeccionContactosDisponible(page);
        if (!estaEnPasoDatosCliente) {
            const mpnCasoActivoTardio = await extraerCasoActivoMpn(page);
            if (mpnCasoActivoTardio) {
                console.log(`[CasoActivo] Detectado tardio ${mpnCasoActivoTardio} antes de Contactos. Cancelando en Bizagi...`);
                await cancelarCasoEnBizagiDesdePortal(page, mpnCasoActivoTardio, { password: BIZAGI_PASSWORD });
                if (intento >= maxIntentosCasoActivo) {
                    throw new Error(`[CRITICO] Se canceló ${mpnCasoActivoTardio} pero el caso activo persiste para '${registro.identificacion}'.`);
                }
                console.log(`[CasoActivo] ${mpnCasoActivoTardio} cancelado tras revalidacion. Reintentando con la misma cédula (${registro.identificacion})...`);
                await page.waitForTimeout(500);
                continue;
            }
            const url = page.url();
            const errorMsg = await page.locator('.p-message-error:visible, .p-toast-message-error:visible, .p-inline-message-error:visible').first().textContent().catch(() => '');
            const txtError = errorMsg ? ` | Error detectado: "${errorMsg.trim()}"` : '';
            throw new Error(`[CRITICO] No aparecio la seccion Contactos luego de Depurar. url=${url}${txtError}`);
        }

        const datosClienteLlenos = await completarDatosClienteNuevo(page);
        if (!datosClienteLlenos) {
            throw new Error("[CRITICO] No se pudo abrir/completar la seccion de datos del cliente nuevo despues de Depurar.");
        }

        // Llenar campos adicionales visibles (Nivel de estudio, Tipo de vivienda, etc.) antes de avanzar
        await etapaValidacionesPrevias(page);

        await continuarResolviendoGestionDocumentalSiPide(page, { maxIntentos: 3 });
        await page.waitForTimeout(FAST_UI ? 250 : 800);
        return;
    }

    throw new Error(`[CRITICO] No se pudo completar etapa inicial para '${registro.identificacion}' tras reintentos.`);
}

async function etapaValidacionesPrevias(page: Page) {
    let hizoValidaciones = false;
    const hizoTiempoExclusion = await llenarFechaSiVisibleYVacia(page, "Tiempo de exclusion", "16-08-2030", { debug: true });
    if (hizoTiempoExclusion) hizoValidaciones = true;

    const hizoCorreoPredeterminado = await validarCorreoPredeterminadoYCorregir(page);
    if (hizoCorreoPredeterminado) hizoValidaciones = true;

    const hizoApnfd = await validarApnfdYSeleccionarNoSiVacio(page);
    if (hizoApnfd) hizoValidaciones = true;

    const labelNivelEstudio = page.locator('text="Nivel de estudio"').first();
    const visibleNivelEstudio = await labelNivelEstudio.isVisible().catch(() => false);
    console.log(`[ValidacionesPrevias] Nivel de estudio visible=${visibleNivelEstudio}`);
    if (visibleNivelEstudio) {
        await clickReintentarListaSiVisible(page, "Nivel de estudio", LIST_RETRY_CLICK_TIMEOUT);
        await seleccionarDefaultSiVacio(page, "Nivel de estudio", 0).catch(() => { });
        hizoValidaciones = true;
    }

    const labelTipoVivienda = page.locator('text="Tipo de vivienda"').first();
    const visibleTipoVivienda = await labelTipoVivienda.isVisible().catch(() => false);
    if (visibleTipoVivienda) {
        await clickReintentarListaSiVisible(page, "Tipo de vivienda", LIST_RETRY_CLICK_TIMEOUT);
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
            // Solo adjuntar documento si el boton Gestion Documental tiene una notificacion/badge con pendientes
            const pendientesRefGestion = await leerPendientesGestionDocumentalEnBoton(page).catch(() => 0);
            if (pendientesRefGestion > 0) {
                const btnGestionRef = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
                await btnGestionRef.waitFor({ state: 'visible', timeout: 10000 });
                await btnGestionRef.click();
                await cargarDocumentoEnGestionDocumental(page, path.resolve(process.cwd(), 'data', 'Documentos', 'evidencia.pdf'));
            }
            hizoValidaciones = true;
        }
    }

    if (hizoValidaciones) {
        await continuarResolviendoGestionDocumentalSiPide(page, { maxIntentos: 3 });
    }
}

const escapeRegexText = (value: string) =>
    escapeRegexTextShared(value);

const esValorDropdownVacio = (value: string) => {
    return esValorDropdownVacioShared(value);
};

const extraerCodigoProducto = (tipoCuenta: string) => {
    return extraerCodigoProductoShared(tipoCuenta);
};
const extraerNombreProducto = (tipoCuenta: string) =>
    extraerNombreProductoShared(tipoCuenta);

const productoRegistroKey = (registro: RegistroExcel) =>
    `${String(registro.identificacion ?? '').trim()}|${String(registro.tipoCuenta ?? '').trim()}`.toUpperCase();

const marcarProductoConfirmado = (registro: RegistroExcel) => {
    productosConfirmadosPorRegistro.add(productoRegistroKey(registro));
    console.log('[Producto] Etapa producto marcada como completada');
};

const productoConfirmadoEnRegistro = (registro: RegistroExcel) =>
    productosConfirmadosPorRegistro.has(productoRegistroKey(registro));

async function marcarProductoConfirmadoSiYaVisible(page: Page, registro: RegistroExcel) {
    const modalProductoVisible = await modalConfiguracionProductoVisible(page).catch(() => false);
    console.log(`[Producto][Guard] modalConfiguracionProductoVisible=${modalProductoVisible}`);
    if (modalProductoVisible) {
        console.log('[Producto][Guard] Producto visible dentro del modal; no cuenta como agregado final');
        return false;
    }
    const productoVisible = await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
    if (!productoVisible) return false;
    console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
    marcarProductoConfirmado(registro);
    return true;
}

async function detectarProductoAgregadoEnUILegacy(
    page: Page,
    seccionProductos: Locator,
    tipoCuenta: string
) {
    const tipoCuentaSafe = String(tipoCuenta ?? '').trim();
    if (!tipoCuentaSafe) return false;

    const tipoCuentaRegex = new RegExp(escapeRegexText(tipoCuentaSafe).replace(/\s+/g, '\\s+'), 'i');
    const nombreProducto = extraerNombreProducto(tipoCuentaSafe);
    const nombreRegex = nombreProducto
        ? new RegExp(escapeRegexText(nombreProducto).replace(/\s+/g, '\\s+'), 'i')
        : null;
    const codigoProducto = extraerCodigoProducto(tipoCuentaSafe);
    const codigoRegex = codigoProducto
        ? new RegExp(`\\b${escapeRegexText(codigoProducto)}\\b`, 'i')
        : null;

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

    const textoSeccion = ((await seccionProductos.innerText().catch(() => '')) || '').trim();
    if (textoSeccion) {
        if (tipoCuentaRegex.test(textoSeccion)) return true;
        if (nombreRegex && nombreRegex.test(textoSeccion)) return true;
        if (codigoRegex && codigoRegex.test(textoSeccion)) return true;
    }

    return false;
}

async function cerrarModalCancelarProcesoSiVisible(page: Page) {
    const modalCancelarProceso = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /(Cancelar|Salir) del proceso de solicitud/i })
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

function modalSolicitudFinalizadaLocatorLegacy(page: Page) {
    return page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i })
        .last();
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

async function cerrarModalSolicitudFinalizadaLegacy(
    page: Page,
    options?: { timeoutMs?: number; waitForAppearMs?: number }
) {
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

    for (let intento = 1; intento <= 5; intento++) {
        const cerro = await cerrarModalFinalizadaRapido(page).catch(() => false);
        if (cerro) return true;
        await page.waitForTimeout(200);
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
            dialogo.locator('button:has-text("Aceptar"), button:has-text("Si"), button:has-text("Sí"), button:has-text("Finalizar")').first(),
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
    const candidatoPaso5 = page.locator(
        'li.p-steps-item:has-text("5"), [data-pc-section="menuitem"]:has-text("5"), .p-steps-item:has(.p-steps-number:has-text("5"))'
    ).first();
    const visible = await candidatoPaso5.isVisible().catch(() => false);
    if (!visible) return false;

    await candidatoPaso5.click({ force: true }).catch(() => { });
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
        const finalizaTrasPaso5 = await btnFinalizar.isVisible().catch(() => false);
        if (finalizaTrasPaso5) return true;

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

        await avanzarAPasoFinalSiHaceFalta(page).catch(() => false);
        await validarGestionDocumentalSiRequerido(page).catch(() => false);

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
            await validarGestionDocumentalSiRequerido(page).catch(() => false);

            const cerroModalTrasFinalizar = await cerrarModalSolicitudFinalizada(page, {
                timeoutMs: 12000,
                waitForAppearMs: 9000,
            }).catch(() => false);
            if (cerroModalTrasFinalizar) return true;
        }

        const finalizadaSinModal = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
        if (finalizadaSinModal) return true;
        await page.waitForTimeout(FAST_UI ? 220 : 700);
    }

    return false;
}

function modalSolicitudFinalizadaLocator(page: Page) {
    return modalSolicitudFinalizadaLocatorShared(page);
}

async function cerrarModalFinalizadaRapido(page: Page) {
    return cerrarModalFinalizadaRapidoShared(page);
}

async function cerrarModalSolicitudFinalizada(
    page: Page,
    options?: { timeoutMs?: number; waitForAppearMs?: number }
) {
    return cerrarModalSolicitudFinalizadaShared(page, options);
}

async function estadoSolicitudFinalizadaVisible(page: Page) {
    return estadoSolicitudFinalizadaVisibleShared(page);
}

async function estadoPendienteFinalizarVisible(page: Page) {
    return estadoPendienteFinalizarVisibleShared(page);
}

async function confirmarFinalizacionSiHayPopup(page: Page) {
    return confirmarFinalizacionSiHayPopupShared(page);
}

async function clickBotonFinalizarRobusto(page: Page) {
    return clickBotonFinalizarRobustoShared(page);
}

async function irAPasoFinalDesdeWizard(page: Page) {
    return irAPasoFinalDesdeWizardShared(page, {
        stepNumber: 5,
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
        const finalizaTrasPaso5 = await btnFinalizar.isVisible().catch(() => false);
        if (finalizaTrasPaso5) return true;

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
        stepNumber: 5,
        getBotonContinuar,
        esperarFinActualizandoSolicitud,
        beforeEachIntent: async (currentPage) => {
            await validarGestionDocumentalSiRequerido(currentPage).catch(() => false);
        },
        afterFinalizarClick: async (currentPage) => {
            await validarGestionDocumentalSiRequerido(currentPage).catch(() => false);
        },
    });
}

async function localizarSeccionProductos(page: Page) {
    return localizarSeccionProductosShared(page, { useBroadFallback: true });
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
    tipoCuenta: string,
    options?: { useGlobalFallback?: boolean }
) {
    return detectarProductoAgregadoEnUIShared(page, seccionProductos, tipoCuenta, options ?? { useGlobalFallback: true });
}

async function estaEnPantallaProductos(page: Page) {
    // Signal 1: Categoría label visible
    const categoriaVisible = await page.getByText(/Categor[ií]a de producto/i).first().isVisible().catch(() => false);

    // Signal 2: Producto label visible
    const productoVisible = await page.getByText(/^Producto$/i).first().isVisible().catch(() => false);

    // Signal 3: "Agregar relacionado" button (typical in Productos section)
    const agregarRelacionadoVisible = await page.getByRole('button', { name: /Agregar relacionado/i }).isVisible().catch(() => false);

    // Signal 4: Check for dropdowns in section or page-wide
    const seccionProductos = await localizarSeccionProductos(page).catch(() => null);
    let dropdownCount = 0;
    if (seccionProductos) {
        dropdownCount = await seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible, select:visible').count().catch(() => 0);
    }
    if (dropdownCount === 0) {
        dropdownCount = await page.locator('body').locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible, select:visible').count().catch(() => 0);
    }

    // Signal 5: "Productos" text visible on page
    const productosTextVisible = await page.getByText(/^Productos$/i).isVisible().catch(() => false);

    // Productos screen is present if we have at least 3 of these signals
    const signals = [categoriaVisible, productoVisible, agregarRelacionadoVisible, dropdownCount >= 1, productosTextVisible];
    const signalsPresent = signals.filter(Boolean).length;

    return signalsPresent >= 3;
}

async function productoAgregadoVisible(page: Page, tipoCuenta: string) {
    const msgSinProductos = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    if (await msgSinProductos.isVisible().catch(() => false)) return false;

    const seccionProductos = await localizarSeccionProductos(page);
    const seccionVisible = await seccionProductos.isVisible().catch(() => false);
    if (seccionVisible) {
        return detectarProductoAgregadoEnUI(page, seccionProductos, tipoCuenta).catch(() => false);
    }
    return detectarProductoAgregadoEnUI(page, page.locator('body'), tipoCuenta, { useGlobalFallback: true }).catch(() => false);
}

async function productoAgregadoComoTarjetaVisible(page: Page, tipoCuenta: string) {
    if (await modalConfiguracionProductoVisible(page).catch(() => false)) return false;

    const msgSinProductos = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    if (await msgSinProductos.isVisible().catch(() => false)) return false;

    const tipoCuentaSafe = String(tipoCuenta ?? '').trim();
    if (!tipoCuentaSafe) return false;

    const nombreProducto = extraerNombreProducto(tipoCuentaSafe);
    const codigoProducto = extraerCodigoProducto(tipoCuentaSafe);
    const candidatos = [tipoCuentaSafe, nombreProducto, codigoProducto]
        .map((valor) => String(valor ?? '').trim())
        .filter(Boolean);

    const seccionProductos = await localizarSeccionProductos(page);
    const seccionVisible = await seccionProductos.isVisible().catch(() => false);
    const scope = seccionVisible ? seccionProductos : page.locator('body');

    return scope.evaluate((root, payload) => {
        const normalize = (value: string) =>
            String(value ?? '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .replace(/\s+/g, ' ')
                .trim()
                .toLowerCase();

        const rootElement = root as HTMLElement;
        const targets = payload.targets.map(normalize).filter(Boolean);
        if (!targets.length) return false;

        const excludedSelector = [
            '.p-dialog',
            '[role="dialog"]',
            '.p-component-overlay',
            '.p-dropdown',
            '[data-pc-name="dropdown"]',
            '[role="combobox"]',
        ].join(',');

        const elements = [rootElement, ...Array.from(rootElement.querySelectorAll('*'))] as HTMLElement[];
        for (const element of elements) {
            if (!(element instanceof HTMLElement)) continue;
            if (element.closest(excludedSelector)) continue;

            const text = normalize(element.innerText || element.textContent || '');
            if (!text) continue;
            if (targets.some((target) => text.includes(target))) {
                return true;
            }
        }

        return false;
    }, { targets: candidatos }).catch(() => false);
}

async function asegurarProductoConfirmadoAntesDeContinuar(page: Page, registro: RegistroExcel): Promise<void | 'producto-confirmado'> {
    if (!await estaEnPantallaProductos(page).catch(() => false)) return;

    console.log('[Producto] Pantalla Productos detectada');
    let seccionProductos = await localizarSeccionProductos(page);
    if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
        console.log('[Producto] Continuando con producto ya confirmado');
        return 'producto-confirmado';
    }
    const productoYaVisible = await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
    if (productoYaVisible) {
        console.log('[Producto] Producto ya estaba agregado, no se volverá a agregar');
        console.log('[Producto] Evitando duplicar producto/cuenta');
        marcarProductoConfirmado(registro);
        console.log('[Producto] Continuando con producto ya confirmado');
        return;
    }

    const estadoCategoria = await leerValorDropdownEnScope(seccionProductos, /Categor[ií]a de producto/i, { timeoutMs: 2500 })
        .catch(() => '');
    const estadoProducto = await leerValorDropdownEnScope(seccionProductos, /^Producto$/i, { timeoutMs: 2500 })
        .catch(() => '');
    const categoriaVacia = esValorDropdownVacio(estadoCategoria);
    const productoVacio = esValorDropdownVacio(estadoProducto);

    console.log(`[Producto] Categoria vacia=${categoriaVacia}`);
    console.log(`[Producto] Producto vacio=${productoVacio}`);

    if (categoriaVacia) {
        console.log('[Producto] Seleccionando categoria Cuentas de Efectivo');
        await seleccionarCategoriaEnSeccionProductos(page, seccionProductos);
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 8000 : 15000).catch(() => false);
        await page.waitForTimeout(FAST_UI ? 600 : 1200);
        seccionProductos = await localizarSeccionProductos(page);
        if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
            console.log('[Producto] Continuando con producto ya confirmado');
            return 'producto-confirmado';
        }
    }

    const productoVisibleAntesDeSeleccion = await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
    if (productoVisibleAntesDeSeleccion) {
        console.log('[Producto] Producto ya estaba agregado, no se volverá a agregar');
        console.log('[Producto] Evitando duplicar producto/cuenta');
        marcarProductoConfirmado(registro);
        console.log('[Producto] Continuando con producto ya confirmado');
        return;
    }

    if (productoVacio || !productoVisibleAntesDeSeleccion) {
        console.log(`[Producto] Seleccionando producto ${registro.tipoCuenta}`);
        await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta);
        const modalAbierto = await modalProductoConfigVisible(page).catch(() => false);
        console.log(`[Producto] Modal configuracion abierto=${modalAbierto}`);
        if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
            console.log('[Producto] Continuando con producto ya confirmado');
            return 'producto-confirmado';
        }
        await etapaSeccionProductosPostSeleccion(page, registro, seccionProductos);
    }

    const productoVisible = await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
    console.log(`[Producto] Producto agregado como tarjeta final visible=${productoVisible}`);
    if (!productoVisible) {
        throw new Error('[Producto][CRITICO] No se confirmo producto agregado; no se puede continuar.');
    }
    marcarProductoConfirmado(registro);
    console.log('[Producto] Producto confirmado, ahora si se puede Continuar hacia Verificaciones');
}

async function modalConfiguracionProductoVisible(page: Page): Promise<boolean> {
    const modalExplicito = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Cuentas de efectivo/i })
        .filter({ hasText: /Balance promedio/i })
        .filter({ hasText: /Tasa|Calcular tasa/i })
        .filter({ hasText: /Aceptar|Agregar/i })
        .first();
    if (await modalExplicito.isVisible().catch(() => false)) return true;
    return modalProductoConfigVisibleShared(page, /Cuentas de efectivo|Balance promedio|Moneda|Tasa|Calcular tasa|Aceptar|Agregar/i);
}

async function modalProductoConfigVisible(page: Page) {
    return modalConfiguracionProductoVisible(page);
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

async function modalProductoConfigVisibleLegacy(page: Page) {
    const modalProducto = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
        .first();
    return modalProducto.isVisible().catch(() => false);
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

async function detectarProductoSeleccionadoEnUI(
    page: Page,
    seccionProductos: Locator,
    tipoCuenta: string
) {
    return detectarProductoSeleccionadoEnUIShared(page, seccionProductos, tipoCuenta, {
        escapeRegexText,
        extraerCodigoProducto,
        modalRegex: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i,
    });
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

    const productoDropdown = dropdowns.nth(1);
    await productoDropdown.scrollIntoViewIfNeeded().catch(() => { });
    await productoDropdown.click({ force: true }).catch(() => { });
    const cerroTrasClickProducto = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    if (cerroTrasClickProducto) {
        await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
        return false;
    }

    const combobox = productoDropdown.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
    let panel: Locator | null = null;
    if (panelId) {
        const byId = page.locator(`#${panelId}`);
        const byIdVisible = await byId.waitFor({ state: 'visible', timeout: Math.min(LIST_PANEL_TIMEOUT, LIST_PANEL_QUICK_TIMEOUT) })
            .then(() => true)
            .catch(() => false);
        if (byIdVisible) panel = byId;
    }
    if (!panel) {
        const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
        const fallbackVisible = await fallback.waitFor({ state: 'visible', timeout: Math.min(LIST_PANEL_TIMEOUT, LIST_PANEL_QUICK_TIMEOUT) })
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
    const hasItems = await items.first().waitFor({ state: 'visible', timeout: LIST_PANEL_TIMEOUT })
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

async function seleccionarProductoConFiltroEnDropdownSecundario(
    page: Page,
    seccionProductos: Locator,
    tipoCuentaRegex: RegExp,
    codigoRegex: RegExp | null,
    textoFiltro: string
) {
    const dropdowns = seccionProductos.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
    const count = await dropdowns.count().catch(() => 0);
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

            const cerroTrasClick = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
            if (cerroTrasClick) {
                console.log('[ProductoFiltro] Click abrio modal de cancelar. Reintentando apertura del dropdown...');
                await page.waitForTimeout(120);
                continue;
            }

            const combobox = productoDropdown.locator('[role="combobox"]').first();
            const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
            let panel: Locator | null = null;
            if (panelId) {
                const byId = page.locator(`#${panelId}`);
                const byIdVisible = await byId.waitFor({ state: 'visible', timeout: 8000 }).then(() => true).catch(() => false);
                if (byIdVisible) panel = byId;
            }
            if (!panel) {
                const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
                const fallbackVisible = await fallback.waitFor({ state: 'visible', timeout: 8000 }).then(() => true).catch(() => false);
                if (fallbackVisible) panel = fallback;
            }
            if (panel) return panel;
        }

        return null;
    };

    const panel = await abrirProductoDropdown();
    if (!panel) return false;

    const filtro = panel
        .locator('input[data-pc-section="filterinput"], input.p-dropdown-filter, input[role="searchbox"], input[type="text"]')
        .first();
    if (await filtro.isVisible().catch(() => false)) {
        await filtro.click({ force: true }).catch(() => { });
        await filtro.fill('').catch(() => { });
        await filtro.fill(textoFiltro).catch(() => { });
        await page.waitForTimeout(300);
    }

    const btnRetryPanel = panel
        .locator('button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button.p-button-warning')
        .first();
    if (await btnRetryPanel.isVisible().catch(() => false)) {
        await btnRetryPanel.click({ force: true }).catch(() => { });
        await page.waitForTimeout(LIST_SCOPE_RETRY_WAIT_MS);
    }

    const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
    const hasItems = await items.first().waitFor({ state: 'visible', timeout: 8000 }).then(() => true).catch(() => false);
    if (!hasItems) return false;

    const byTipo = items.filter({ hasText: tipoCuentaRegex }).first();
    if (await byTipo.isVisible().catch(() => false)) {
        await byTipo.click({ force: true }).catch(() => { });
        return true;
    }

    if (codigoRegex) {
        const byCodigo = items.filter({ hasText: codigoRegex }).first();
        if (await byCodigo.isVisible().catch(() => false)) {
            await byCodigo.click({ force: true }).catch(() => { });
            return true;
        }
    }

    return false;
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
                return cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false);
            },
        }
    );
}

async function seleccionarProductoEnSeccionProductos(
    page: Page,
    seccionProductos: Locator,
    tipoCuenta: string
) {
    const modalProductoVisible = await modalConfiguracionProductoVisible(page).catch(() => false);
    if (!modalProductoVisible && await productoAgregadoComoTarjetaVisible(page, tipoCuenta).catch(() => false)) {
        console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
        return;
    }
    console.log(`[SeccionProductos] Iniciando seleccion de producto: '${tipoCuenta}'`);
    return seleccionarProductoCuentaEfectivoNuevo(page, seccionProductos, tipoCuenta, {
        confirmarSeleccionProductoRapida,
        detectarProductoSeleccionadoEnUI,
        esValorDropdownVacio,
        escapeRegexText,
        extraerCodigoProducto,
        extraerNombreProducto,
        leerValorDropdownEnScope,
        seleccionarDropdownEnScopePorTexto,
        seleccionarProductoPorDropdownSecundario,
        seleccionarProductoConFiltroEnDropdownSecundario,
        beforeAttempt: async (currentPage) => {
            console.log('[SeccionProductos] beforeAttempt: verificando estado de la UI...');
            if (await modalProductoConfigVisible(currentPage)) {
                console.log('[SeccionProductos] Modal de configuracion ya visible.');
                return;
            }
            await cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false);
            await clickReintentarListaSiVisible(currentPage, 'Producto', LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
        },
        maxIntentos: PRODUCTO_MAX_INTENTOS,
        probeTimeoutMs: LIST_PROBE_TIMEOUT,
        retryWaitMs: PRODUCTO_RETRY_WAIT_MS,
    });
}

async function etapaSeccionProductosPostSeleccion(
    page: Page,
    registro: RegistroExcel,
    seccionProductosInicial: Locator
) {
    if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
        return;
    }
    let seccionProductos = seccionProductosInicial;
    const msgSinProductos = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    const confirmarProductoAgregado = async () => {
        if (await modalProductoConfigVisible(page)) return false;
        seccionProductos = await localizarSeccionProductos(page);
        const enUI = await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
        if (enUI) return true;
        const sinProductos = await msgSinProductos.isVisible().catch(() => false);
        if (sinProductos) return false;
        await page.waitForTimeout(FAST_UI ? 500 : 1200);
        if (await modalProductoConfigVisible(page)) return false;
        seccionProductos = await localizarSeccionProductos(page);
        return productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
    };

    const clickAgregarProductoDesdeBalance = async () => {
        const modalProducto = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
            .first();
        const modalVisible = await modalProducto.isVisible().catch(() => false);
        if (modalVisible) {
            const fallbacksModal = [
                modalProducto.locator('xpath=//*[contains(normalize-space(.),"Balance promedio")]/following::button[contains(@class,"p-button")][1]').first(),
                modalProducto.locator('.p-inputgroup > .p-button:visible, .p-inputgroup > button:visible').first(),
                modalProducto.locator('button:has(.pi-plus):visible, button:has(i.pi-plus):visible').first(),
            ];
            for (const btn of fallbacksModal) {
                const visible = await btn.isVisible().catch(() => false);
                if (!visible) continue;
                await btn.scrollIntoViewIfNeeded().catch(() => { });
                await btn.click({ force: true }).catch(() => { });
                return true;
            }
            return true;
        }

        const bloqueBalance = page
            .locator('xpath=//*[contains(normalize-space(.),"Balance promedio")]/ancestor::*[self::div or self::fieldset][1]')
            .first();
        const bloqueVisible = await bloqueBalance.isVisible().catch(() => false);
        if (bloqueVisible) {
            const candidatos = [
                bloqueBalance.locator('.p-inputgroup > .p-button:visible, .p-inputgroup > button:visible').first(),
                bloqueBalance.locator('button:has(.pi-plus):visible, button:has(i.pi-plus):visible').first(),
                bloqueBalance.locator('xpath=.//*[contains(normalize-space(.),"Balance promedio")]/following::button[contains(@class,"p-button")][1]').first(),
            ];

            for (const btn of candidatos) {
                const visible = await btn.isVisible().catch(() => false);
                if (!visible) continue;
                const txt = ((await btn.innerText().catch(() => '')) || '').trim();
                if (/^cancelar$|^aceptar$|^continuar$|^guardar$|^finalizar$/i.test(txt)) continue;
                await btn.scrollIntoViewIfNeeded().catch(() => { });
                await btn.click({ force: true }).catch(() => { });
                return true;
            }
        }
        return false;
    };

    const llenarBalancePromedioEnContexto = async () => {
        const valorObjetivo = "5000";
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
                const placeholder = ((await input.getAttribute('placeholder').catch(() => '')) || '').trim();
                if (/calcular tasa/i.test(placeholder)) continue;
                return input;
            }

            const inputsVisibles = modalProducto.locator('input:visible');
            const total = await inputsVisibles.count().catch(() => 0);
            for (let i = 0; i < total; i++) {
                const input = inputsVisibles.nth(i);
                const editable = await input.isEditable().catch(() => false);
                if (!editable) continue;
                const placeholder = ((await input.getAttribute('placeholder').catch(() => '')) || '').trim();
                if (/calcular tasa/i.test(placeholder)) continue;
                return input;
            }
            return null;
        };

        const leerValorBalanceModal = async () => {
            const input = await encontrarInputBalanceModal();
            if (!input) return '';
            return (await input.inputValue().catch(() => '')).trim();
        };

        const llenarInputDirecto = async (input: Locator) => {
            const visible = await input.isVisible().catch(() => false);
            if (!visible) return false;
            const editable = await input.isEditable().catch(() => false);
            if (!editable) return false;
            await input.scrollIntoViewIfNeeded().catch(() => { });
            await input.click({ force: true }).catch(() => { });
            await input.press('Control+A').catch(() => { });
            await input.press('Delete').catch(() => { });
            await input.fill(valorObjetivo).catch(async () => {
                await input.pressSequentially(valorObjetivo, { delay: FAST_UI ? 10 : 20 }).catch(() => { });
            });
            await input.dispatchEvent('input').catch(() => { });
            await input.dispatchEvent('change').catch(() => { });
            await input.blur().catch(() => { });
            await page.waitForTimeout(FAST_UI ? 80 : 180);
            let valorFinal = (await input.inputValue().catch(() => '')).trim();
            if (!valorFinal) {
                await input.evaluate((el, v) => {
                    const inputEl = el instanceof HTMLInputElement ? el : null;
                    if (!inputEl) return;
                    inputEl.value = String(v);
                    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                    inputEl.dispatchEvent(new Event('blur', { bubbles: true }));
                }, valorObjetivo).catch(() => { });
                await page.waitForTimeout(FAST_UI ? 80 : 180);
                valorFinal = (await input.inputValue().catch(() => '')).trim();
            }
            return !!valorFinal;
        };

        if (modalVisible) {
            const inputBalance = await encontrarInputBalanceModal();
            if (!inputBalance) {
                console.log("[BalanceModal][ce-new] No se encontro input editable para 'Balance promedio'.");
            }
            if (inputBalance && await llenarInputDirecto(inputBalance)) return true;

            const candidatos = [
                modalProducto.locator('input.p-inputnumber-input:visible').first(),
                modalProducto.locator('input.p-inputtext:visible').first(),
            ];
            for (const input of candidatos) {
                if (await llenarInputDirecto(input)) {
                    const valorModal = await leerValorBalanceModal();
                    if (valorModal) return true;
                }
            }
        }

        await llenarInputNumber(page, "Balance promedio", valorObjetivo).catch(() => { });

        if (modalVisible) {
            const valorModal = await leerValorBalanceModal();
            if (!valorModal) {
                console.log("[BalanceModal][ce-new] El valor de 'Balance promedio' sigue vacio tras intentos de llenado.");
            }
            return !!valorModal;
        }

        const valorGlobal = (await page
            .locator('label')
            .filter({ hasText: /Balance promedio/i })
            .first()
            .locator('xpath=following::input[1]')
            .first()
            .inputValue()
            .catch(() => '')).trim();
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
            const txt = ((await btn.innerText().catch(() => '')) || '').trim();
            if (/^cancelar$|^aceptar$|^continuar$|^guardar$|^finalizar$/i.test(txt)) continue;
            await btn.scrollIntoViewIfNeeded().catch(() => { });
            await btn.click({ force: true }).catch(() => { });
            return true;
        }
        return false;
    };

    const cerrarModalProductoConAceptar = async () => {
        const modalProducto = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
            .first();
        const visible = await modalProducto.isVisible().catch(() => false);
        if (!visible) return true;

        const botones = [
            modalProducto.getByRole('button', { name: /^Aceptar$/i }).first(),
            modalProducto.locator('button:has-text("Aceptar")').first(),
            modalProducto.locator('.p-dialog-footer button:visible').last(),
        ];
        for (const btn of botones) {
            const btnVisible = await btn.isVisible().catch(() => false);
            if (!btnVisible) continue;
            await btn.scrollIntoViewIfNeeded().catch(() => { });
            const clicked = await btn.click({ force: true, timeout: 2500 }).then(() => true).catch(() => false);
            if (!clicked) {
                const handle = await btn.elementHandle().catch(() => null);
                if (handle) {
                    await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => { });
                }
            }
            const hidden = await modalProducto.waitFor({ state: 'hidden', timeout: FAST_UI ? 2500 : 6000 })
                .then(() => true)
                .catch(() => false);
            if (hidden) return true;
        }

        return !(await modalProducto.isVisible().catch(() => false));
    };

    // Esperar a que abra el modal de configuracion del producto tras la seleccion.
    // El modal puede tardar en aparecer en staging (Angular CDK lazy render).
    const esperarModalProductoConfigLocal = async (timeoutMs = 8000) => {
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
            if (await modalProductoConfigVisible(page)) return true;
            // Fallback: cualquier dialogo que parezca de configuracion de producto
            const anyDialog = page.locator('.p-dialog:visible, [role="dialog"]:visible').last();
            const anyVisible = await anyDialog.isVisible().catch(() => false);
            if (anyVisible) {
                const dialogText = ((await anyDialog.textContent().catch(() => '')) || '').toLowerCase();
                // Excluir dialogo de cancelar proceso
                if (!/cancelar proceso|confirmar cancelaci/i.test(dialogText) &&
                    /balance|monto|tasa|moneda|plazo|cuenta|interes/i.test(dialogText)) return true;
            }
            await page.waitForTimeout(400);
        }
        return false;
    };

    const esperarProductoVisibleTrasAgregar = async (timeoutMs = FAST_UI ? 5000 : 8000) => {
        const inicio = Date.now();
        while (Date.now() - inicio < timeoutMs) {
            const visible = await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
            if (visible) return true;
            const confirmado = await confirmarProductoAgregado().catch(() => false);
            if (confirmado) return true;
            await page.waitForTimeout(FAST_UI ? 250 : 400);
        }
        return false;
    };

    let productoAgregado = false;
    for (let intentoProducto = 1; intentoProducto <= 3; intentoProducto++) {
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        const yaAgregadoAntes = await confirmarProductoAgregado();
        if (yaAgregadoAntes) {
            console.log(`[Producto] productoAgregado=true al inicio del intento ${intentoProducto}`);
            productoAgregado = true;
            marcarProductoConfirmado(registro);
            break;
        }

        // Si el modal aun no esta visible, esperar hasta 8s para que cargue
        const hayModalAhora = await modalProductoConfigVisible(page);
        if (!hayModalAhora) {
            console.log(`[Producto] Intento ${intentoProducto}: modal no visible. Esperando hasta 8s...`);
            const modalAbrio = await esperarModalProductoConfigLocal(8000);
            console.log(`[Producto] Intento ${intentoProducto}: modal abrio=${modalAbrio}`);
            if (!modalAbrio) {
                // Sin modal → comprobar si el producto quedo agregado directamente
                const confirmoSinModal = await confirmarProductoAgregado();
                if (confirmoSinModal) {
                    console.log(`[Producto] Producto agregado directamente (sin modal de configuracion).`);
                    productoAgregado = true;
                    marcarProductoConfirmado(registro);
                    break;
                }
                if (intentoProducto < 3) {
                    if (!await modalConfiguracionProductoVisible(page).catch(() => false)
                        && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false)) {
                        console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
                        productoAgregado = true;
                        marcarProductoConfirmado(registro);
                        break;
                    }
                    console.log(`[Producto] Sin modal y sin producto. Re-seleccionando...`);
                    if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
                        productoAgregado = true;
                        break;
                    }
                    await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
                    await page.waitForTimeout(FAST_UI ? 500 : 1200);
                    continue;
                }
                throw new Error("[CRITICO] No aparecio modal de configuracion de producto tras seleccion y espera.");
            }
        }

        const balanceLleno = await llenarBalancePromedioEnContexto();
        console.log(`[Producto] Intento ${intentoProducto}: balanceLleno=${balanceLleno}`);
        console.log(`[Producto] Balance/tasa completados=${balanceLleno}`);
        if (!balanceLleno) {
            const confirmoSinBalance = await confirmarProductoAgregado();
            if (confirmoSinBalance) { productoAgregado = true; marcarProductoConfirmado(registro); break; }
            if (intentoProducto < 3) {
                if (!await modalConfiguracionProductoVisible(page).catch(() => false)
                    && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false)) {
                    console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
                    productoAgregado = true;
                    marcarProductoConfirmado(registro);
                    break;
                }
                if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
                    productoAgregado = true;
                    break;
                }
                await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
                await page.waitForTimeout(FAST_UI ? 500 : 1200);
                continue;
            }
            throw new Error("[CRITICO] No se pudo llenar 'Balance promedio' en el modal de producto.");
        }
        const clicCalcularTasa = await clickCalcularTasaEnContexto();
        if (clicCalcularTasa) {
            console.log(`[Producto] Intento ${intentoProducto}: tasa calculada. Esperando...`);
            await page.waitForTimeout(2000);
        }
        const clicAgregar = await clickAgregarProductoDesdeBalance();
        console.log(`[Producto] Intento ${intentoProducto}: clicAgregar=${clicAgregar}`);
        console.log('[Producto] Click en Agregar');
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
        if (!clicAgregar) {
            if (intentoProducto < 3) {
                if (!await modalConfiguracionProductoVisible(page).catch(() => false)
                    && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false)) {
                    console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
                    productoAgregado = true;
                    marcarProductoConfirmado(registro);
                    break;
                }
                if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
                    productoAgregado = true;
                    break;
                }
                await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
                await page.waitForTimeout(FAST_UI ? 500 : 1200);
                continue;
            }
            throw new Error(`[CRITICO] No se pudo encontrar boton para agregar producto en 'Balance promedio'.`);
        }
        await page.waitForTimeout(FAST_UI ? 600 : 1800);
        const aceptoConfirmacion = await cerrarModalProductoConAceptar().catch(() => false);
        console.log(`[Producto] Confirmacion de producto aceptada=${aceptoConfirmacion}`);
        await page.waitForTimeout(FAST_UI ? 500 : 1200);
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        const aparecioTrasAgregar = await esperarProductoVisibleTrasAgregar();
        if (aparecioTrasAgregar) {
            console.log('[Producto] Producto apareció después del primer Agregar; no se reintentará Agregar');
            console.log('[Producto] Evitando duplicar producto/cuenta');
            console.log('[Producto] Producto agregado como tarjeta final visible=true');
            productoAgregado = true;
            marcarProductoConfirmado(registro);
            break;
        }

        const confirmoAhora = await confirmarProductoAgregado();
        console.log(`[Producto] Intento ${intentoProducto}: confirmoAhora=${confirmoAhora}`);
        if (confirmoAhora) {
            productoAgregado = true;
            marcarProductoConfirmado(registro);
            break;
        }

        if (intentoProducto < 3) {
            const visibleAntesDeReintentar = !await modalConfiguracionProductoVisible(page).catch(() => false)
                && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
            if (visibleAntesDeReintentar) {
                console.log('[Producto] Producto apareció después del primer Agregar; no se reintentará Agregar');
                console.log('[Producto] Evitando duplicar producto/cuenta');
                productoAgregado = true;
                marcarProductoConfirmado(registro);
                break;
            }
            if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
                productoAgregado = true;
                break;
            }
            await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
            await page.waitForTimeout(FAST_UI ? 500 : 1200);
        }
    }

    if (!productoAgregado) {
        if (await modalProductoConfigVisible(page)) {
            await cerrarModalProductoConAceptar().catch(() => false);
            await page.waitForTimeout(FAST_UI ? 350 : 900);
            const confirmoTrasCerrar = await confirmarProductoAgregado();
            if (confirmoTrasCerrar) {
                marcarProductoConfirmado(registro);
                return;
            }
        }

        const sinProductosFinal = await msgSinProductos.isVisible().catch(() => false);
        if (sinProductosFinal) {
            throw new Error(`[CRITICO] No se agrego el producto '${registro.tipoCuenta}' en la simulacion.`);
        }
        throw new Error(`[CRITICO] No se detecta el producto agregado en pantalla: '${registro.tipoCuenta}'.`);
    }
}

async function etapaSeccionProductos(page: Page, registro: RegistroExcel) {
    if (!await modalConfiguracionProductoVisible(page).catch(() => false)
        && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false)) {
        console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
        marcarProductoConfirmado(registro);
        return;
    }
    let flujoNuevoActivado = true;
    if (flujoNuevoActivado) {
        const { seccionProductos } = await asegurarSeccionProductosVisible(page, {
            maxIntentos: 18,
            waitInicialMs: 90000,
            waitLoopMs: 12000,
            scrollPauseMs: 350,
            step2WaitMs: 20000,
            locateSection: localizarSeccionProductos,
            waitForUpdating: esperarFinActualizandoSolicitud,
            closeCancelModal: async (currentPage) => cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false),
            beforeStart: async (currentPage) => cerrarModalCancelarProcesoSiVisible(currentPage).catch(() => false),
            isPreProductos: estaEnPantallaPreProductos,
            onPreProductos: async (currentPage) => {
                // Validacion extra: si por algun motivo ya detectamos la seccion de productos (ej. scroll parcial)
                // NO le damos a continuar para evitar saltarnos la seleccion.
                const seccionExistente = await localizarSeccionProductos(currentPage);
                const visible = await seccionExistente.isVisible().catch(() => false);
                const countDrops = await seccionExistente.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible').count().catch(() => 0);
                if (visible && countDrops >= 2) {
                    console.log('[NAV-PRODUCTO] Seccion productos detectada durante onPreProductos. Omitiendo click en Continuar.');
                    return;
                }

                const btnContinuar = getBotonContinuar(currentPage);
                const continuarVisible = await btnContinuar.isVisible().catch(() => false);
                const continuarEnabled = continuarVisible
                    ? await btnContinuar.isEnabled().catch(() => false)
                    : false;
                if (continuarVisible && continuarEnabled) {
                    try {
                        await continuarResolviendoGestionDocumentalSiPide(currentPage, { maxIntentos: 2 });
                    } catch (e) {
                        const msg = e instanceof Error ? e.message : String(e);
                        if (/\[CRITICO\].*Gestion documental/i.test(msg)) throw e;
                        console.log(`[NAV-PRODUCTO] Click en Continuar para llegar a productos (${msg})...`);
                        await btnContinuar.click({ force: true }).catch(() => { });
                        await esperarFinActualizandoSolicitud(currentPage, 25000).catch(() => false);
                        await currentPage.waitForTimeout(600);
                    }
                }
            },
            onFailure: async (currentPage) => {
                const urlActual = currentPage.url();
                const dropdownsVisibles = await currentPage
                    .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
                    .count()
                    .catch(() => 0);
                const enPantallaPrevia = await estaEnPantallaPreProductos(currentPage).catch(() => false);
                const actualizandoVisible = await currentPage.getByText(/Actualizando solicitud/i).first().isVisible().catch(() => false);
                const innerTextPagina = await currentPage.evaluate(() => (document.body as HTMLElement).innerText || '').catch(() => '');
                const resumenTexto = innerTextPagina.replace(/\s+/g, ' ').substring(0, 300);
                console.log(`[NAV-PRODUCTO] categoriaVisible=false url=${urlActual} dropdownsVisibles=${dropdownsVisibles} prePantalla=${enPantallaPrevia} actualizando=${actualizandoVisible}`);
                console.log(`[NAV-PRODUCTO] texto visible (primeros 300 chars): "${resumenTexto}"`);
                const btnGDVisible = await currentPage
                    .getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i })
                    .first()
                    .isVisible()
                    .catch(() => false);
                if (dropdownsVisibles >= 2 && !enPantallaPrevia && !btnGDVisible) {
                    console.log(`[NAV-PRODUCTO] ${dropdownsVisibles} dropdowns visibles detectados. Usando seccionProductos fallback.`);
                    return localizarSeccionProductos(currentPage);
                }
                if (btnGDVisible) {
                    console.log(`[NAV-PRODUCTO] Boton GD visible -> todavia en formulario pre-productos. Intentando continuarResolviendo...`);
                    try {
                        await continuarResolviendoGestionDocumentalSiPide(currentPage, { maxIntentos: 2 });
                    } catch (e) {
                        const msg = e instanceof Error ? e.message : String(e);
                        console.log(`[NAV-PRODUCTO][WARN] continuarResolviendo fallo: ${msg}`);
                    }
                    throw new Error("[CRITICO] No se pudo navegar a pantalla de productos; pagina sigue en formulario con Gestion Documental pendiente.");
                }
                return null;
            },
        });

        await seleccionarCategoriaEnSeccionProductos(page, seccionProductos);
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        // Esperar a que el dropdown de Producto se pueble tras seleccionar la categoria.
        // Angular lanza un fetch async al cambiar la categoria; el dropdown de Producto
        // puede mostrar "" o "Seleccione" hasta que llega la respuesta (no siempre hay
        // spinner "Actualizando"). Espera fija + polling por items en el panel.
        await esperarFinActualizandoSolicitud(page, FAST_UI ? 8000 : 15000).catch(() => { });
        await page.waitForTimeout(FAST_UI ? 2000 : 3500); // margen para fetch async de productos

        const seccionActualizada = await localizarSeccionProductos(page);
        await seccionActualizada.waitFor({ state: 'visible', timeout: 10000 }).catch(() => { });

        // Se remueve la espera manual y apertura del dropdown aqui, 
        // ya que seleccionarProductoEnSeccionProductos ya maneja la espera interna 
        // de forma mas robusta a traves de seleccionarDropdownEnScopePorTexto.

        await seleccionarProductoEnSeccionProductos(page, seccionActualizada, registro.tipoCuenta);
        return etapaSeccionProductosPostSeleccion(page, registro, seccionActualizada);
    }

    console.log(`[SeccionProductos] inicio url=${page.url()}`);
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    const alertSinProductosPaso = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    const labelPropositoPaso = page.getByText(/Prop(?:o|\u00f3)sito/i).first();
    const tituloProductos = page.getByText(/Productos/i).first();
    let seccionProductos = await localizarSeccionProductos(page);

    console.log(`[SeccionProductos] esperarFinActualizando(90s)...`);
    await esperarFinActualizandoSolicitud(page, 90000).catch(() => false);
    console.log(`[SeccionProductos] esperarFinActualizando completado`);

    let categoriaVisible = false;
    for (let i = 1; i <= 18; i++) {
        console.log(`[SeccionProductos] iter=${i}`);
        await esperarFinActualizandoSolicitud(page, 12000).catch(() => false);
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        await tituloProductos.scrollIntoViewIfNeeded().catch(() => { });
        await page.mouse.wheel(0, 600).catch(() => { });
        await page.waitForTimeout(350);

        seccionProductos = await localizarSeccionProductos(page);
        const seccionVisible = await seccionProductos.isVisible().catch(() => false);
        const dropdownsEnSeccion = await seccionProductos
            .locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible')
            .count()
            .catch(() => 0);
        categoriaVisible = seccionVisible && dropdownsEnSeccion >= 2;
        if (categoriaVisible) break;

        const cayoEnPasoSinProducto = await alertSinProductosPaso.isVisible().catch(() => false);
        const yaEnTaller = await labelPropositoPaso.isVisible().catch(() => false);
        const enPantallaPrevia = await estaEnPantallaPreProductos(page);

        if (enPantallaPrevia) {
            const btnContinuar = getBotonContinuar(page);
            const continuarVisible = await btnContinuar.isVisible().catch(() => false);
            const continuarEnabled = continuarVisible
                ? await btnContinuar.isEnabled().catch(() => false)
                : false;
            if (continuarVisible && continuarEnabled) {
                try {
                    await continuarResolviendoGestionDocumentalSiPide(page, { maxIntentos: 2 });
                } catch (e) {
                    const msg = e instanceof Error ? e.message : String(e);
                    if (/\[CRITICO\].*Gestion documental/i.test(msg)) throw e;
                    await btnContinuar.click({ force: true }).catch(() => { });
                    await esperarFinActualizandoSolicitud(page, 25000).catch(() => false);
                    await page.waitForTimeout(600);
                }
            }
        }

        if (cayoEnPasoSinProducto || yaEnTaller || enPantallaPrevia) {
            const step2 = page.locator(
                'li.p-steps-item:has-text("2"), [data-pc-section="menuitem"]:has-text("2"), .p-steps-item:has(.p-steps-number:has-text("2"))'
            ).first();
            const step2Activo = await page
                .locator('li.p-steps-item.p-highlight:has-text("2"), .p-steps-item[aria-current="step"]:has-text("2")')
                .first()
                .isVisible()
                .catch(() => false);
            if (!step2Activo && await step2.isVisible().catch(() => false)) {
                await step2.click({ force: true }).catch(() => { });
                await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
                await esperarFinActualizandoSolicitud(page, 20000).catch(() => false);
                await page.waitForTimeout(700);
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
        // Loguear texto visible de la pagina para diagnostico de estructura en staging
        const innerTextPagina = await page.evaluate(() => (document.body as HTMLElement).innerText || '').catch(() => '');
        const resumenTexto = innerTextPagina.replace(/\s+/g, ' ').substring(0, 300);
        console.log(`[NAV-PRODUCTO] categoriaVisible=false url=${urlActual} paso3SinProducto=${enPaso3SinProducto} dropdownsVisibles=${dropdownsVisibles} prePantalla=${enPantallaPrevia} actualizando=${actualizandoVisible}`);
        console.log(`[NAV-PRODUCTO] texto visible (primeros 300 chars): "${resumenTexto}"`);
        // Si hay dropdowns visibles en la pagina, usar seccion fallback en vez de tirar error.
        // Pero SOLO si el boton "Gestion Documental" NO esta visible; si lo esta, seguimos en
        // el formulario previo (los 6 dropdowns son campos del formulario, no de productos).
        const btnGDVisible = await page
            .getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i })
            .first()
            .isVisible()
            .catch(() => false);
        if (dropdownsVisibles >= 2 && !enPantallaPrevia && !btnGDVisible) {
            console.log(`[NAV-PRODUCTO] ${dropdownsVisibles} dropdowns visibles detectados. Usando seccionProductos fallback.`);
            categoriaVisible = true;
        } else if (btnGDVisible) {
            // Todavia en formulario — intentar avanzar con GD upload como ultima oportunidad
            console.log(`[NAV-PRODUCTO] Boton GD visible → todavia en formulario pre-productos. Intentando continuarResolviendo...`);
            try {
                await continuarResolviendoGestionDocumentalSiPide(page, { maxIntentos: 2 });
            } catch (e) {
                const msg = e instanceof Error ? e.message : String(e);
                console.log(`[NAV-PRODUCTO][WARN] continuarResolviendo fallo: ${msg}`);
            }
            throw new Error("[CRITICO] No se pudo navegar a pantalla de productos; pagina sigue en formulario con Gestion Documental pendiente.");
        } else {
            throw new Error("[CRITICO] No se pudo visualizar 'Categoria de producto' para agregar el producto.");
        }
    }

    await seleccionarCategoriaEnSeccionProductos(page, seccionProductos);
    await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
    seccionProductos = await localizarSeccionProductos(page);
    await seccionProductos.waitFor({ state: 'visible', timeout: 10000 }).catch(() => { });
    await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta);

    const msgSinProductos = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    const confirmarProductoAgregado = async () => {
        if (await modalProductoConfigVisible(page)) return false;
        seccionProductos = await localizarSeccionProductos(page);
        const enUI = await detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta).catch(() => false);
        if (enUI) return true;
        const sinProductos = await msgSinProductos.isVisible().catch(() => false);
        if (sinProductos) return false;
        await page.waitForTimeout(FAST_UI ? 500 : 1200);
        if (await modalProductoConfigVisible(page)) return false;
        seccionProductos = await localizarSeccionProductos(page);
        return detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta).catch(() => false);
    };

    const clickAgregarProductoDesdeBalance = async () => {
        const modalProducto = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
            .first();
        const modalVisible = await modalProducto.isVisible().catch(() => false);
        if (modalVisible) {
            const fallbacksModal = [
                modalProducto.locator('xpath=//*[contains(normalize-space(.),"Balance promedio")]/following::button[contains(@class,"p-button")][1]').first(),
                modalProducto.locator('.p-inputgroup > .p-button:visible, .p-inputgroup > button:visible').first(),
                modalProducto.locator('button:has(.pi-plus):visible, button:has(i.pi-plus):visible').first(),
            ];
            for (const btn of fallbacksModal) {
                const visible = await btn.isVisible().catch(() => false);
                if (!visible) continue;
                await btn.scrollIntoViewIfNeeded().catch(() => { });
                await btn.click({ force: true }).catch(() => { });
                return true;
            }
            return true;
        }

        const bloqueBalance = page
            .locator('xpath=//*[contains(normalize-space(.),"Balance promedio")]/ancestor::*[self::div or self::fieldset][1]')
            .first();
        const bloqueVisible = await bloqueBalance.isVisible().catch(() => false);
        if (bloqueVisible) {
            const candidatos = [
                bloqueBalance.locator('.p-inputgroup > .p-button:visible, .p-inputgroup > button:visible').first(),
                bloqueBalance.locator('button:has(.pi-plus):visible, button:has(i.pi-plus):visible').first(),
                bloqueBalance.locator('xpath=.//*[contains(normalize-space(.),"Balance promedio")]/following::button[contains(@class,"p-button")][1]').first(),
            ];

            for (const btn of candidatos) {
                const visible = await btn.isVisible().catch(() => false);
                if (!visible) continue;
                const txt = ((await btn.innerText().catch(() => '')) || '').trim();
                if (/^cancelar$|^aceptar$|^continuar$|^guardar$|^finalizar$/i.test(txt)) continue;
                await btn.scrollIntoViewIfNeeded().catch(() => { });
                await btn.click({ force: true }).catch(() => { });
                return true;
            }
        }
        return false;
    };

    const llenarBalancePromedioEnContexto = async () => {
        const valorObjetivo = "5000";
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
                const placeholder = ((await input.getAttribute('placeholder').catch(() => '')) || '').trim();
                if (/calcular tasa/i.test(placeholder)) continue;
                return input;
            }

            const inputsVisibles = modalProducto.locator('input:visible');
            const total = await inputsVisibles.count().catch(() => 0);
            for (let i = 0; i < total; i++) {
                const input = inputsVisibles.nth(i);
                const editable = await input.isEditable().catch(() => false);
                if (!editable) continue;
                const placeholder = ((await input.getAttribute('placeholder').catch(() => '')) || '').trim();
                if (/calcular tasa/i.test(placeholder)) continue;
                return input;
            }
            return null;
        };

        const leerValorBalanceModal = async () => {
            const input = await encontrarInputBalanceModal();
            if (!input) return '';
            return (await input.inputValue().catch(() => '')).trim();
        };

        const llenarInputDirecto = async (input: Locator) => {
            const visible = await input.isVisible().catch(() => false);
            if (!visible) return false;
            const editable = await input.isEditable().catch(() => false);
            if (!editable) return false;
            await input.scrollIntoViewIfNeeded().catch(() => { });
            await input.click({ force: true }).catch(() => { });
            await input.press('Control+A').catch(() => { });
            await input.press('Delete').catch(() => { });
            await input.fill(valorObjetivo).catch(async () => {
                await input.pressSequentially(valorObjetivo, { delay: FAST_UI ? 10 : 20 }).catch(() => { });
            });
            await input.dispatchEvent('input').catch(() => { });
            await input.dispatchEvent('change').catch(() => { });
            await input.blur().catch(() => { });
            await page.waitForTimeout(FAST_UI ? 80 : 180);
            let valorFinal = (await input.inputValue().catch(() => '')).trim();
            if (!valorFinal) {
                await input.evaluate((el, v) => {
                    const inputEl = el instanceof HTMLInputElement ? el : null;
                    if (!inputEl) return;
                    inputEl.value = String(v);
                    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                    inputEl.dispatchEvent(new Event('blur', { bubbles: true }));
                }, valorObjetivo).catch(() => { });
                await page.waitForTimeout(FAST_UI ? 80 : 180);
                valorFinal = (await input.inputValue().catch(() => '')).trim();
            }
            return !!valorFinal;
        };

        if (modalVisible) {
            const inputBalance = await encontrarInputBalanceModal();
            if (!inputBalance) {
                console.log("[BalanceModal][ce-new] No se encontro input editable para 'Balance promedio'.");
            }
            if (inputBalance && await llenarInputDirecto(inputBalance)) return true;

            const candidatos = [
                modalProducto.locator('input.p-inputnumber-input:visible').first(),
                modalProducto.locator('input.p-inputtext:visible').first(),
            ];
            for (const input of candidatos) {
                if (await llenarInputDirecto(input)) {
                    const valorModal = await leerValorBalanceModal();
                    if (valorModal) return true;
                }
            }
        }

        await llenarInputNumber(page, "Balance promedio", valorObjetivo).catch(() => { });

        if (modalVisible) {
            const valorModal = await leerValorBalanceModal();
            if (!valorModal) {
                console.log("[BalanceModal][ce-new] El valor de 'Balance promedio' sigue vacio tras intentos de llenado.");
            }
            return !!valorModal;
        }

        const valorGlobal = (await page
            .locator('label')
            .filter({ hasText: /Balance promedio/i })
            .first()
            .locator('xpath=following::input[1]')
            .first()
            .inputValue()
            .catch(() => '')).trim();
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
            const txt = ((await btn.innerText().catch(() => '')) || '').trim();
            if (/^cancelar$|^aceptar$|^continuar$|^guardar$|^finalizar$/i.test(txt)) continue;
            await btn.scrollIntoViewIfNeeded().catch(() => { });
            await btn.click({ force: true }).catch(() => { });
            return true;
        }
        return false;
    };

    const cerrarModalProductoConAceptar = async () => {
        const modalProducto = page
            .locator('.p-dialog:visible, [role="dialog"]:visible')
            .filter({ hasText: /Cuentas de efectivo|Balance promedio|Moneda|Tasa/i })
            .first();
        const visible = await modalProducto.isVisible().catch(() => false);
        if (!visible) return true;

        const botones = [
            modalProducto.getByRole('button', { name: /^Aceptar$/i }).first(),
            modalProducto.locator('button:has-text("Aceptar")').first(),
            modalProducto.locator('.p-dialog-footer button:visible').last(),
        ];
        for (const btn of botones) {
            const btnVisible = await btn.isVisible().catch(() => false);
            if (!btnVisible) continue;
            await btn.scrollIntoViewIfNeeded().catch(() => { });
            const clicked = await btn.click({ force: true, timeout: 2500 }).then(() => true).catch(() => false);
            if (!clicked) {
                const handle = await btn.elementHandle().catch(() => null);
                if (handle) {
                    await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => { });
                }
            }
            const hidden = await modalProducto.waitFor({ state: 'hidden', timeout: FAST_UI ? 2500 : 6000 })
                .then(() => true)
                .catch(() => false);
            if (hidden) return true;
        }

        return !(await modalProducto.isVisible().catch(() => false));
    };

    let productoAgregado = false;
    for (let intentoProducto = 1; intentoProducto <= 3; intentoProducto++) {
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        const yaAgregadoAntes = await confirmarProductoAgregado();
        if (yaAgregadoAntes) {
            productoAgregado = true;
            break;
        }

        const balanceLleno = await llenarBalancePromedioEnContexto();
        if (!balanceLleno) {
            if (intentoProducto < 3) {
                if (!await modalConfiguracionProductoVisible(page).catch(() => false)
                    && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false)) {
                    console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
                    productoAgregado = true;
                    break;
                }
                await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
                await page.waitForTimeout(FAST_UI ? 220 : 600);
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
                if (!await modalConfiguracionProductoVisible(page).catch(() => false)
                    && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false)) {
                    console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
                    productoAgregado = true;
                    break;
                }
                await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
                await page.waitForTimeout(FAST_UI ? 220 : 600);
                continue;
            }
            throw new Error(`[CRITICO] No se pudo encontrar boton para agregar producto en 'Balance promedio'.`);
        }
        await page.waitForTimeout(FAST_UI ? 400 : 1500);
        await cerrarModalProductoConAceptar().catch(() => false);
        await page.waitForTimeout(FAST_UI ? 500 : 1200);
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        const confirmoAhora = await confirmarProductoAgregado();
        if (confirmoAhora) {
            productoAgregado = true;
            break;
        }

        if (intentoProducto < 3) {
            if (!await modalConfiguracionProductoVisible(page).catch(() => false)
                && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false)) {
                console.log('[Producto] Producto ya visible antes de seleccionar; se omite selección/agregar');
                productoAgregado = true;
                break;
            }
            await seleccionarProductoEnSeccionProductos(page, seccionProductos, registro.tipoCuenta).catch(() => { });
            await page.waitForTimeout(FAST_UI ? 220 : 600);
        }
    }

    if (!productoAgregado) {
        if (await modalProductoConfigVisible(page)) {
            await cerrarModalProductoConAceptar().catch(() => false);
            await page.waitForTimeout(FAST_UI ? 350 : 900);
            const confirmoTrasCerrar = await confirmarProductoAgregado();
            if (confirmoTrasCerrar) {
                return;
            }
        }

        const sinProductosFinal = await msgSinProductos.isVisible().catch(() => false);
        if (sinProductosFinal) {
            throw new Error(`[CRITICO] No se agrego el producto '${registro.tipoCuenta}' en la simulacion.`);
        }
        throw new Error(`[CRITICO] No se detecta el producto agregado en pantalla: '${registro.tipoCuenta}'.`);
    }
}

async function estaEnPantallaTallerProductos(page: Page) {
    const checks = await Promise.all([
        page.locator('label').filter({ hasText: /Prop(?:o|\u00f3)sito/i }).first().isVisible().catch(() => false),
        page.locator('label').filter({ hasText: /Monto inicial/i }).first().isVisible().catch(() => false),
        page.locator('label').filter({ hasText: /Instrumento/i }).first().isVisible().catch(() => false),
        page.getByRole('button', { name: /Crear cuenta/i }).first().isVisible().catch(() => false),
    ]);
    const visibles = checks.filter(Boolean).length;
    return visibles >= 2 || (checks[0] && checks[3]);
}

async function esperarPantallaTallerProductos(
    page: Page,
    options?: { timeoutMs?: number; pollMs?: number }
) {
    const timeoutMs = options?.timeoutMs ?? (FAST_UI ? 9000 : 20000);
    const pollMs = options?.pollMs ?? (FAST_UI ? 140 : 300);
    const inicio = Date.now();

    while (Date.now() - inicio < timeoutMs) {
        if (await estaEnPantallaTallerProductos(page)) return true;
        await page.waitForTimeout(pollMs);
    }
    return false;
}

async function etapaRelacionadosYAsociacion(page: Page, registro: RegistroExcel) {
    await agregarRelacionadoSiAplica(page, registro);
    let postProductoLlenadoDesdeCumplimiento = false;
    const llenarPostProductoPostOfac = async () => {
        if (postProductoLlenadoDesdeCumplimiento) return;
        console.log('[Relacionados] Pantalla post-producto recibida desde cumplimiento, no se hará Continuar adicional');
        console.log('[PostProducto] Iniciando llenado inmediatamente después de OFAC');
        const completo = await completarInformacionClientePostProductoAntesDeTaller(page, { required: true });
        if (!completo) {
            throw new Error('[CRITICO] No se pudo completar la pantalla post-producto después de OFAC.');
        }
        postProductoLlenadoDesdeCumplimiento = true;
        console.log('[PostProducto] Llenado post-OFAC completado');
    };
    const avanzarConProductoConfirmado = async (contexto: string) => {
        console.log('[Producto] Producto confirmado; cortando selección y avanzando con Continuar');
        const resultado = await intentarAvanceRealHaciaTaller(page, contexto, {
            maxClicks: 2,
            tipoCuenta: registro.tipoCuenta,
        });
        if (resultado === 'gestion-documental') {
            await resolverGestionDocumentalPostContinuarPrimeraPantalla(page, { maxIntentos: 3 }).catch((e) => {
                const msg = e instanceof Error ? e.message : String(e);
                if (/\[CRITICO\]/i.test(msg)) throw e;
            });
            return 'continuar';
        }
        if (resultado === 'post-producto') {
            await llenarPostProductoPostOfac();
            return 'post-producto';
        }
        if (resultado === 'taller') {
            return 'taller';
        }
        return 'continuar';
    };

    const verificacionInicial = await procesarVerificacionesEspeciales(page).catch(() => null);
    if (verificacionInicial?.tipo === 'cumplimiento' && verificacionInicial.estado === 'post-producto') {
        await llenarPostProductoPostOfac();
    }

    const msgSinProductosDespuesContinuar = page.getByText(/No se agregaron productos en simulaci(?:o|\u00f3)n/i).first();
    const labelPropositoTaller = page.locator('label').filter({ hasText: /Prop(?:o|\u00f3)sito/i }).first();
    let seccionProductos = await localizarSeccionProductos(page);

    if (!postProductoLlenadoDesdeCumplimiento) {
        const estadoProductoInicial = await asegurarProductoConfirmadoAntesDeContinuar(page, registro);
        if (estadoProductoInicial === 'producto-confirmado') {
            const avanceTrasProductoConfirmado = await avanzarConProductoConfirmado('producto confirmado inicial');
            if (avanceTrasProductoConfirmado === 'taller') {
                return labelPropositoTaller;
            }
        }
    }

    const completoPostProductoInicial = await completarInformacionClientePostProductoAntesDeTaller(page, { required: false });
    if (completoPostProductoInicial) {
        await page.waitForTimeout(FAST_UI ? 250 : 700);
    }

    const yaEnTaller = await esperarPantallaTallerProductos(page, { timeoutMs: FAST_UI ? 1500 : 3000 });
    if (yaEnTaller) {
        return labelPropositoTaller;
    }

    for (let intentoAvance = 1; intentoAvance <= 3; intentoAvance++) {
        const verificacionAntesAvance = await procesarVerificacionesEspeciales(page).catch(() => null);
        if (verificacionAntesAvance?.tipo === 'cumplimiento' && verificacionAntesAvance.estado === 'post-producto') {
            await llenarPostProductoPostOfac();
            continue;
        }
        console.log(`[Relacionados] intento ${intentoAvance}/3 url=${page.url()}`);
        const resultadoInicial = await intentarAvanceRealHaciaTaller(
            page,
            `etapa relacionados ${intentoAvance}/3`,
            { maxClicks: 2, tipoCuenta: registro.tipoCuenta }
        );

        if (resultadoInicial === 'gestion-documental') {
            await resolverGestionDocumentalPostContinuarPrimeraPantalla(page, { maxIntentos: 3 }).catch((e) => {
                const msg = e instanceof Error ? e.message : String(e);
                if (/\[CRITICO\]/i.test(msg)) throw e;
            });
        } else if (resultadoInicial === 'post-producto') {
            await llenarPostProductoPostOfac();
            continue;
        } else if (resultadoInicial === 'taller') {
            return labelPropositoTaller;
        } else if (resultadoInicial === 'productos') {
            if (await marcarProductoConfirmadoSiYaVisible(page, registro)) {
                const avanceTrasProductoConfirmado = await avanzarConProductoConfirmado(`producto confirmado en relacionados ${intentoAvance}/3`);
                if (avanceTrasProductoConfirmado === 'taller') {
                    return labelPropositoTaller;
                }
                continue;
            }
            const productoVisible = !await modalConfiguracionProductoVisible(page).catch(() => false)
                && await productoAgregadoComoTarjetaVisible(page, registro.tipoCuenta).catch(() => false);
            if (productoVisible) {
                marcarProductoConfirmado(registro);
                console.log('[Producto] Producto ya estaba agregado, no se volverá a agregar');
                console.log('[Producto] Evitando duplicar producto/cuenta');
                console.log('[Producto] Continuando con producto ya confirmado');
                continue;
            }
            const estadoProducto = await asegurarProductoConfirmadoAntesDeContinuar(page, registro);
            if (estadoProducto === 'producto-confirmado') {
                const avanceTrasProductoConfirmado = await avanzarConProductoConfirmado(`producto confirmado tras asegurar ${intentoAvance}/3`);
                if (avanceTrasProductoConfirmado === 'taller') {
                    return labelPropositoTaller;
                }
            }
            continue;
        }

        console.log(`[Relacionados] esperarPantallaTaller... url=${page.url()}`);
        const llegoATaller = await esperarPantallaTallerProductos(page);
        console.log(`[Relacionados] llegoATaller=${llegoATaller}`);
        if (llegoATaller) {
            return labelPropositoTaller;
        }

        console.log(`[Relacionados] completarInformacionClientePostProducto...`);
        const verificacionAntesCompletarPostProducto = await procesarVerificacionesEspeciales(page).catch(() => null);
        if (verificacionAntesCompletarPostProducto?.tipo === 'cumplimiento' && verificacionAntesCompletarPostProducto.estado === 'post-producto') {
            await llenarPostProductoPostOfac();
            continue;
        }
        const completoPostProducto = await completarInformacionClientePostProductoAntesDeTaller(page, { required: false });
        if (completoPostProducto) {
            await page.waitForTimeout(FAST_UI ? 300 : 800);
            // Si el portal nos regreso a la pantalla de datos del cliente, navegar de vuelta a productos
            const enProductosAhora = await page.getByText(/Categor[ií]a de producto/i).first().isVisible().catch(() => false);
            if (!enProductosAhora) {
                console.log('[Relacionados] Campos completados. Navegando de vuelta a la pantalla de productos...');
                await continuarResolviendoGestionDocumentalSiPide(page, { maxIntentos: 2 }).catch(() => { });
            }
            continue;
        }

        seccionProductos = await localizarSeccionProductos(page);
        const seccionProductosVisible = await seccionProductos.isVisible().catch(() => false);
        const sinProductosTrasContinuar = await msgSinProductosDespuesContinuar.isVisible().catch(() => false);
        const productoEnUITrasContinuar = seccionProductosVisible
            ? await detectarProductoAgregadoEnUI(page, seccionProductos, registro.tipoCuenta).catch(() => false)
            : true;
        if (sinProductosTrasContinuar || (seccionProductosVisible && !productoEnUITrasContinuar)) {
            throw new Error(`[CRITICO] Se pulso Continuar sin productos agregados para '${registro.tipoCuenta}'.`);
        }

        if (intentoAvance >= 3) {
            throw new Error("No se pudo abrir Taller de productos despues de Continuar.");
        }
        await page.waitForTimeout(FAST_UI ? 250 : 700);
    }

    return labelPropositoTaller;
}

async function etapaTallerProductos(
    page: Page,
    registro: RegistroExcel,
    labelPropositoTaller: Locator,
    capturasRef: string[]
) {
    const visibleProposito = await labelPropositoTaller.waitFor({ state: "visible", timeout: 60000 })
        .then(() => true)
        .catch(() => false);
    if (!visibleProposito) {
        const enTaller = await estaEnPantallaTallerProductos(page);
        if (!enTaller) {
            throw new Error("[CRITICO] No se detecto pantalla de Taller de productos.");
        }
    }

    await clickReintentarListaSiVisible(page, /Prop(?:o|\u00f3)sito/i, LIST_RETRY_CLICK_TIMEOUT);
    await seleccionarDropdownPorCampo(page, /Prop(?:o|\u00f3)sito/i, 1);
    const estadoProposito = await leerEstadoDropdown(page, /Prop(?:o|\u00f3)sito/i, { timeoutMs: 6000 }).catch(() => ({ vacio: true, texto: '' }));
    if (estadoProposito.vacio) {
        await clickReintentarListaSiVisible(page, /Prop(?:o|\u00f3)sito/i, LIST_RETRY_CLICK_TIMEOUT).catch(() => false);
        await seleccionarDropdownPorCampo(page, /Prop(?:o|\u00f3)sito/i, 1).catch(() => { });
    }

    await llenarInputNumber(page, "Monto inicial", "5000");
    const inputMontoInicial = page
        .locator("label")
        .filter({ hasText: /Monto inicial/i })
        .first()
        .locator('xpath=following::input[1]')
        .first();
    let valorMontoInicial = (await inputMontoInicial.inputValue().catch(() => "")).trim();
    if (!valorMontoInicial) {
        await llenarInputNumber(page, "Monto inicial", "5000");
        valorMontoInicial = (await inputMontoInicial.inputValue().catch(() => "")).trim();
    }

    await seleccionarInstrumentoRobusto(page);
    const estadoInstrumento = await leerEstadoDropdown(page, /Instrumento/i, { timeoutMs: 6000 }).catch(() => ({ vacio: true, texto: '' }));
    if (estadoInstrumento.vacio) {
        await seleccionarInstrumentoRobusto(page).catch(() => { });
    }

    await llenarInputNumber(page, /Rango m(?:a|\u00e1)ximo/i, "1000000");
    await llenarInputNumber(page, /Dep(?:o|\u00f3)sito en efectivo/i, "5000");

    const inputRangoMaximo = page
        .locator("label")
        .filter({ hasText: /Rango m(?:a|\u00e1)ximo/i })
        .first()
        .locator('xpath=following::input[1]')
        .first();
    const inputDepositoEfectivo = page
        .locator("label")
        .filter({ hasText: /Dep(?:o|\u00f3)sito en efectivo/i })
        .first()
        .locator('xpath=following::input[1]')
        .first();

    let valorRango = (await inputRangoMaximo.inputValue().catch(() => "")).trim();
    let valorDeposito = (await inputDepositoEfectivo.inputValue().catch(() => "")).trim();
    if (!valorRango) {
        await llenarInputNumber(page, /Rango m(?:a|\u00e1)ximo/i, "1000000");
        valorRango = (await inputRangoMaximo.inputValue().catch(() => "")).trim();
    }
    if (!valorDeposito) {
        await llenarInputNumber(page, /Dep(?:o|\u00f3)sito en efectivo/i, "5000");
        valorDeposito = (await inputDepositoEfectivo.inputValue().catch(() => "")).trim();
    }
    if (!valorRango || !valorDeposito) {
        throw new Error(`Campos requeridos vacios en Taller. Rango='${valorRango}' Deposito='${valorDeposito}'`);
    }
    const estadoPropositoFinal = await leerEstadoDropdown(page, /Prop(?:o|\u00f3)sito/i, { timeoutMs: 4000 }).catch(() => ({ vacio: true, texto: '' }));
    if (estadoPropositoFinal.vacio) {
        throw new Error("[CRITICO] Campo 'Proposito' quedo vacio en Taller.");
    }
    if (!valorMontoInicial) {
        throw new Error("[CRITICO] Campo 'Monto inicial' quedo vacio en Taller.");
    }
    const estadoInstrumentoFinal = await leerEstadoDropdown(page, /Instrumento/i, { timeoutMs: 4000 }).catch(() => ({ vacio: true, texto: '' }));
    if (estadoInstrumentoFinal.vacio) {
        throw new Error("[CRITICO] Campo 'Instrumento' quedo vacio en Taller.");
    }

    const btnCrearCuenta = page.getByRole('button', { name: /Crear cuenta/i }).first();
    await btnCrearCuenta.waitFor({ state: 'visible', timeout: 15000 });
    const crearEnabled = await btnCrearCuenta.isEnabled().catch(() => false);
    if (!crearEnabled) {
        throw new Error("El boton 'Crear cuenta' esta deshabilitado.");
    }
    await btnCrearCuenta.click();

    const creandoTxt = page.getByText(/Creando cuenta en efectivo/i).first();
    const toastCuentaOk = page.getByText(/Informaci(?:o|\u00f3)n guardada|Cuenta .* generad/i).first();
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

    // Cerrar el modal "Solicitud finalizada" ANTES de tomar screenshot y de cualquier otro click.
    // El modal puede aparecer ya en este punto (ganó la race con el toast).
    // Intentarlo aquí evita que el click en Continuar (más abajo) dispare sobre una página bloqueada.
    const modalYaVisible = await modalFinalizadaPreview.isVisible().catch(() => false);
    if (modalYaVisible) {
        console.log(`[Finalizacion][${registro.identificacion}] Modal "Solicitud finalizada" visible post-Crear. Cerrando...`);
        await cerrarModalSolicitudFinalizada(page, { timeoutMs: 10000, waitForAppearMs: 500 }).catch(() => { });
        await page.waitForTimeout(400);
    }

    const evidenciaPng2 = await capturarCuentaComoPNG(page, registro.identificacion, `${registro.tipoCuenta}__2`);
    capturasRef.push(evidenciaPng2);

    // Solo intentar Continuar si la solicitud no está ya finalizada (modal cerrado = finalizada)
    const yaFinalizada = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
    if (!yaFinalizada) {
        await getBotonContinuar(page).click().catch(() => { });
        await esperarFinActualizandoSolicitud(page, 15000).catch(() => false);
    }

    const finalizada = await finalizarSolicitudSiPendiente(page, registro, {
        timeoutMs: FINALIZACION_TIMEOUT_MS,
        maxIntentos: FINALIZACION_MAX_INTENTOS,
    }).catch(() => false);
    if (!finalizada) {
        throw new Error(`[NO_FINALIZADA] No se pudo confirmar finalizacion para '${registro.identificacion}'.`);
    }
}

async function etapaEvidencias(page: Page, registro: RegistroExcel, capturasRef: string[]) {
    const evidenciaPng = await capturarCuentaComoPNG(page, registro.identificacion, registro.tipoCuenta);
    capturasRef.push(evidenciaPng);
}



test('Cuenta Efectivo Cliente Nuevo - desde Excel', async () => {
    const session = await launchPortalSession({
        defaultContinueOnError: false,
        initialUrl: PORTAL_MULTIPRODUCT_URL,
        preferNewPage: true,
        extraArgs: [
            '--no-default-browser-check',
            '--disable-features=SigninInterception,msEdgeIdentityFirstRunExperience,EdgeFirstRunExperience,msEdgeSigninPromo,msEdgeSyncPromo,EdgeSyncPromo,ExplicitBrowserSignin,msEdgeAccountConsistency,msEdgeProfileSwitchingToast',
            '--disable-sync',
        ],
    });
    const { page } = session;
    let failedRegistros: Array<{ identificacion: string; error: string }> = [];
    const cedulasProcesadas = new Set<string>();

    try {
        const ids = leerRegistrosDesdeExcel('data/ids.xlsx', 'CE_Clientes_Nuevos', {
            soloClienteNuevoSI: true,
        });

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

                cedulasProcesadas.add(registro.identificacion);
                console.log(`[Excel][${registro.identificacion}] Marcado diferido en lote. Pendientes=${cedulasProcesadas.size}`);
            },
            onRegistroError: async (_registro, errorMsg) => {
                const esCritico = /\[CRITICO\]/i.test(errorMsg);
                if (session.continueOnRegistroError && !esCritico) {
                    return 'continue';
                }
                return 'throw';
            },
        });
        failedRegistros = runState.failedRegistros;

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
    } finally {
        if (cedulasProcesadas.size) {
            try {
                const marcado = marcarCedulasProcesadasEnExcel('data/ids.xlsx', [...cedulasProcesadas], {
                    hojaClientesNuevos: 'Clientes_Nuevos',
                    hojaCedulasNuevas: 'Ced_Clientes_Nuevos',
                    colorHex: 'FFF4B183',
                });
                console.log(
                    `[Excel][Batch] Cedulas=${marcado.targets} | Cliente_Nuevo->NO: ${marcado.updatedClienteNuevo} | ` +
                    `Cedula coloreada: ${marcado.paintedCedula}`
                );
            } catch (error) {
                const msg = error instanceof Error ? error.message : String(error);
                console.log(`[Excel][Batch][WARN] No se pudo aplicar el marcado diferido: ${msg}`);
            }
        }
        await closeBrowserSession(session, failedRegistros);
    }
});
