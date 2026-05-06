import { test, type Page } from '@playwright/test';
import {
    abrirBizagiDirectoDebug,
    asegurarLoginBizagiDebug,
    buscarMpnEnBizagiDebug,
    abrirMpnEnBizagiDebug,
    abrirGestionCoincidenciasDebug,
    aprobarGestionCoincidenciasMixtaDebug,
    resolverModoBizagiDebug,
    guardarYSiguienteDebug,
} from '../../helpers/ceNewBizagiDebug';

const BIZAGI_URL = process.env.BIZAGI_URL || 'https://test-bscrd-santacruz.bizagi.com/';
const DEBUG_MPN = process.env.DEBUG_MPN || '';
const DEBUG_BIZAGI_MODE = (process.env.DEBUG_BIZAGI_MODE || 'AUTO').toUpperCase();

test('DEBUG: Bizagi - Gestión de Coincidencias (directo, sin portal)', async ({ browser, context }) => {
    console.log('[DEBUG-BIZAGI] INICIO');

    if (!DEBUG_MPN) {
        console.log('[DEBUG-BIZAGI] ERROR: DEBUG_MPN no definida');
        console.log('[DEBUG-BIZAGI] Ejemplo: $env:DEBUG_MPN="MPN-1617"');
        throw new Error('DEBUG_MPN no está definida. Establécela con: $env:DEBUG_MPN="MPN-XXXX"');
    }

    console.log(`[DEBUG-BIZAGI] DEBUG_MPN=${DEBUG_MPN}`);
    console.log(`[DEBUG-BIZAGI] DEBUG_BIZAGI_MODE=${DEBUG_BIZAGI_MODE}`);
    console.log(`[DEBUG-BIZAGI] BIZAGI_URL=${BIZAGI_URL}`);

    let bizagiPage: Page | null = null;

    try {
        // Abrir Bizagi directo
        console.log('[DEBUG-BIZAGI] Abriendo Bizagi directo');
        bizagiPage = await abrirBizagiDirectoDebug(context, BIZAGI_URL);

        // Hacer login si es necesario
        const loginRequerido = await asegurarLoginBizagiDebug(bizagiPage);
        console.log(`[DEBUG-BIZAGI] Login requerido=${loginRequerido}`);

        // Buscar MPN
        console.log(`[DEBUG-BIZAGI] Buscando MPN=${DEBUG_MPN}`);
        await buscarMpnEnBizagiDebug(bizagiPage, DEBUG_MPN);
        console.log('[DEBUG-BIZAGI] MPN encontrado');

        // Abrir la fila encontrada
        console.log(`[DEBUG-BIZAGI] Abriendo MPN=${DEBUG_MPN}`);
        await abrirMpnEnBizagiDebug(bizagiPage, DEBUG_MPN);
        console.log('[DEBUG-BIZAGI] MPN abierto');

        // Verificar que está en Gestionar Coincidencias o una actividad de cumplimiento
        console.log('[DEBUG-BIZAGI] Verificando Gestionar Coincidencias');
        await abrirGestionCoincidenciasDebug(bizagiPage);
        console.log('[DEBUG-BIZAGI] Actividad Gestionar Coincidencias abierta');

        // Resolver modo automático si es AUTO
        let modoEjecutable: 'OFAC' | 'PLAFT' | 'LEXIS' | 'MIXTO' = DEBUG_BIZAGI_MODE as any;

        if (DEBUG_BIZAGI_MODE === 'AUTO') {
            console.log('[DEBUG-BIZAGI] Modo AUTO detectado, resolviendo automáticamente...');
            modoEjecutable = await resolverModoBizagiDebug(bizagiPage);
            console.log(`[DEBUG-BIZAGI] Modo AUTO resuelto a: ${modoEjecutable}`);
        } else {
            // Validar modo especificado
            const modoValido = ['OFAC', 'PLAFT', 'LEXIS', 'MIXTO'].includes(DEBUG_BIZAGI_MODE);
            if (!modoValido) {
                throw new Error(`[DEBUG-BIZAGI][CRITICO] Modo inválido: ${DEBUG_BIZAGI_MODE}. Valores válidos: AUTO, OFAC, PLAFT, LEXIS, MIXTO`);
            }
        }

        // Ejecutar modo resuelto
        console.log(`[DEBUG-BIZAGI] Ejecutando modo=${modoEjecutable}`);
        await aprobarGestionCoincidenciasMixtaDebug(bizagiPage, modoEjecutable);
        console.log(`[DEBUG-BIZAGI] Modo ${modoEjecutable} completado`);

        // Guardar y Siguiente
        console.log('[DEBUG-BIZAGI] Guardando y avanzando');
        await guardarYSiguienteDebug(bizagiPage);
        console.log('[DEBUG-BIZAGI] Guardado completado');

        // Tomar screenshot final
        const screenshotPath = `artifacts/evidencias_tmp/debug-bizagi-${DEBUG_BIZAGI_MODE}-${DEBUG_MPN}-final.png`;
        await bizagiPage.screenshot({ path: screenshotPath }).catch((e) => {
            console.log(`[DEBUG-BIZAGI][WARN] No se pudo capturar screenshot: ${e}`);
        });
        console.log(`[DEBUG-BIZAGI] Screenshot: ${screenshotPath}`);

        console.log('[DEBUG-BIZAGI] FIN OK');
    } catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        console.log(`[DEBUG-BIZAGI][CRITICO] ${errorMsg}`);
        throw error;
    } finally {
        // NO cerrar el navegador
        if (bizagiPage) {
            await bizagiPage.bringToFront().catch(() => {});
        }
        console.log('[DEBUG-BIZAGI] Test completado. Navegador mantenido abierto para inspección.');
    }
});
