/**
 * Helper para debug directo de Bizagi sin pasar por el Portal Comercial.
 * Copia la lógica real de ceExBizagi.ts que ya funciona en CE-NEW.
 * NO inventa nuevas estrategias de login/búsqueda.
 */

import { Page, Locator } from '@playwright/test';

// ────────────────────────────────────
// HELPERS COPIADAS DE ceExBizagi.ts
// ────────────────────────────────────

async function esBizagiHomePage(page: Page): Promise<boolean> {
  const loginSels = [
    'input#user', 'input#username', 'input#password', 'input[name="username"]',
    'input[name="password"]', 'select#user', 'button#btn-login', 'button#btn-adminlogin'
  ];
  for (const sel of loginSels) {
    if (await page.locator(sel).first().isVisible().catch(() => false)) {
      return false;
    }
  }

  const sels = [
    'input#menuQuery', 'input[name="menuQuery"]', '#ui-bizagi-wp-widget-searchContainer input',
    'table#ui-bizagi-wp-app-inbox-grid-cases', '[id*="menuListInbox"]', '.bz-icon-gear'
  ];
  for (const sel of sels) {
    if (await page.locator(sel).first().isVisible().catch(() => false)) {
      return true;
    }
  }
  return false;
}

async function pickVisible(page: Page, locators: Locator[]): Promise<Locator> {
  for (const loc of locators) {
    if (await loc.isVisible().catch(() => false)) return loc;
  }
  return locators[0];
}

async function findVisibleLocator(page: Page, selectors: string[], label: string = "Element"): Promise<Locator> {
  for (const selector of selectors) {
    const loc = page.locator(selector).first();
    if (await loc.isVisible().catch(() => false)) {
      console.log(`[DEBUG-BIZAGI] ${label} encontrado con selector: ${selector}`);
      return loc;
    }
  }
  return page.locator(selectors[0]).first();
}

async function cerrarMenusBizagi(bizagiPage: Page) {
  const menus = bizagiPage.locator('[id*="MenuContainer"], [class*="PopupMenu"], [role="menu"]');
  const count = await menus.count();
  for (let i = 0; i < count; i++) {
    await menus.nth(i).press('Escape').catch(() => {});
  }
}

async function navegarABandejaEntradaBizagi(bizagiPage: Page) {
  await cerrarMenusBizagi(bizagiPage);
  const candidatos = [
    bizagiPage.getByRole('link', { name: /Bandeja de entrada/i }).first(),
    bizagiPage.getByRole('button', { name: /Bandeja de entrada/i }).first(),
    bizagiPage.locator('a,button,span,div').filter({ hasText: /^Bandeja de entrada$/i }).first(),
  ];

  for (const candidato of candidatos) {
    if (await candidato.isVisible({ timeout: 2000 }).catch(() => false)) {
      console.log('[DEBUG-BIZAGI][Nav] Click en Bandeja de entrada');
      await candidato.scrollIntoViewIfNeeded().catch(() => {});
      const ok = await candidato.click({ force: true }).then(() => true).catch(() => false);
      if (ok) {
        await bizagiPage.waitForTimeout(600);
        return;
      }
    }
  }
  console.log('[DEBUG-BIZAGI][Nav] No se encontró botón Bandeja de entrada');
}

async function pantallaCumplimientoAbiertaEstricaBizagi(bizagiPage: Page): Promise<boolean> {
  const sels = [
    'div:has-text("Gestionar Coincidencias")',
    'text=Gestionar Coincidencias',
    'h1:has-text("Gestionar Coincidencias")',
    '.activity-title:has-text("Gestionar Coincidencias")',
    'span:has-text("Gestionar Coincidencias")',
    '[id*="GestionarCoincidencias"]',
    '[class*="GestionarCoincidencias"]'
  ];
  for (const sel of sels) {
    if (await bizagiPage.locator(sel).first().isVisible({ timeout: 500 }).catch(() => false)) {
      return true;
    }
  }
  return false;
}

// ────────────────────────────────────
// FUNCIONES PRINCIPALES COPIADAS
// ────────────────────────────────────

export async function abrirBizagiDirectoDebug(
    context: import('@playwright/test').BrowserContext,
    bizagiUrl: string = process.env.BIZAGI_URL || 'https://test-bscrd-santacruz.bizagi.com/'
): Promise<Page> {
    console.log('[DEBUG-BIZAGI] Abriendo Bizagi directo');
    const bizagiPage = await context.newPage();
    await bizagiPage.goto(bizagiUrl, { waitUntil: 'domcontentloaded', timeout: 30000 }).catch((e) => {
        console.log(`[DEBUG-BIZAGI][WARN] Timeout esperado en goto: ${e}`);
    });
    console.log(`[DEBUG-BIZAGI] URL cargada: ${bizagiPage.url()}`);
    return bizagiPage;
}

export async function asegurarLoginBizagiDebug(
    bizagiPage: Page,
    usuario: string = process.env.BIZAGI_USER || 'domain\\admon',
    password: string = process.env.BIZAGI_PASSWORD || 'H0la1234.'
): Promise<boolean> {
    console.log(`[DEBUG-BIZAGI][Login] Iniciando verificación de login para: ${usuario}`);

    if (await esBizagiHomePage(bizagiPage)) {
        console.log('[DEBUG-BIZAGI][Login] Ya estamos en la página de inicio de Bizagi.');
        return false;
    }

    const fin = Date.now() + 45000;

    while (Date.now() < fin) {
        if (await esBizagiHomePage(bizagiPage)) return false;

        const inputUser = await pickVisible(bizagiPage, [
            bizagiPage.locator('#user').first(),
            bizagiPage.locator('input[name="username"]').first(),
            bizagiPage.locator('input#username').first(),
            bizagiPage.locator('input[autocomplete="username"]').first(),
        ]);

        const inputPass = await pickVisible(bizagiPage, [
            bizagiPage.getByLabel(/Contrase/i).first(),
            bizagiPage.locator('input#password').first(),
            bizagiPage.locator('input[type="password"]').first(),
        ]);

        const btnIngresar = await pickVisible(bizagiPage, [
            bizagiPage.getByRole('button', { name: /^Ingresar$/i }).first(),
            bizagiPage.locator('#btn-login').first(),
            bizagiPage.locator('button#btn-login').first(),
            bizagiPage.locator('button:has-text("Ingresar")').first(),
        ]);

        const uVis = await inputUser.isVisible().catch(() => false);
        const pVis = await inputPass.isVisible().catch(() => false);
        const iVis = await btnIngresar.isVisible().catch(() => false);

        if (!uVis && !pVis && !iVis) {
            await bizagiPage.waitForTimeout(600).catch(() => {});
            continue;
        }

        console.log(`[DEBUG-BIZAGI][Login] Detectado U=${uVis}, P=${pVis}, I=${iVis}`);

        if (uVis) {
            console.log(`[DEBUG-BIZAGI][Login] Llenando usuario: ${usuario}`);
            await inputUser.fill(usuario).catch(() => {});
            await bizagiPage.waitForTimeout(300);
        }

        if (pVis) {
            console.log(`[DEBUG-BIZAGI][Login] Llenando contraseña`);
            await inputPass.fill(password).catch(() => {});
            await bizagiPage.waitForTimeout(300);
        }

        if (iVis) {
            console.log(`[DEBUG-BIZAGI][Login] Haciendo click en Ingresar`);
            await btnIngresar.click({ force: true }).catch(() => {});
            await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
            await bizagiPage.waitForTimeout(2000);
        }

        console.log(`[DEBUG-BIZAGI][Login] Login completado`);
        return true;
    }

    throw new Error('[DEBUG-BIZAGI][CRITICO] No se detectó login ni sesión activa después de 45s');
}

export async function buscarMpnEnBizagiDebug(bizagiPage: Page, mpn: string): Promise<void> {
    console.log(`[DEBUG-BIZAGI][Buscar] Buscando MPN=${mpn}`);

    await navegarABandejaEntradaBizagi(bizagiPage).catch(() => false);
    await bizagiPage.waitForTimeout(500);

    const selectoresBuscador = [
        '#menuQuery',
        'input#menuQuery',
        '#ui-bizagi-wp-widget-searchContainer input',
        'input[id*="menuQuery"]',
        'input[name="menuQuery"]',
        'input[placeholder="Buscar"]',
    ];

    const inputBuscador = await findVisibleLocator(bizagiPage, selectoresBuscador, 'BuscadorSuperior');

    if (!await inputBuscador.isVisible().catch(() => false)) {
        throw new Error(`[DEBUG-BIZAGI][CRITICO] No se encontró buscador superior después de navegar a bandeja`);
    }

    console.log(`[DEBUG-BIZAGI][Buscar] Buscando ${mpn} en el buscador superior...`);

    // Cerrar menus que puedan tapar
    await cerrarMenusBizagi(bizagiPage).catch(() => {});

    // Llenar el buscador
    await inputBuscador.focus().catch(() => {});
    await inputBuscador.fill(mpn).catch(async () => {
        await inputBuscador.click({ timeout: 1000 }).catch(() => {});
        await inputBuscador.fill(mpn).catch(() => {});
    });

    // Disparar búsqueda con múltiples métodos
    await inputBuscador.evaluate((el, v) => {
        const inputEl = el as HTMLInputElement;
        inputEl.value = String(v);
        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
        inputEl.dispatchEvent(new Event('change', { bubbles: true }));
        inputEl.dispatchEvent(new Event('search', { bubbles: true }));
        inputEl.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter' }));
        inputEl.dispatchEvent(new KeyboardEvent('keypress', { bubbles: true, key: 'Enter', code: 'Enter' }));
        inputEl.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Enter', code: 'Enter' }));
    }, mpn).catch(() => {});

    await inputBuscador.press('Enter').catch(() => {});
    await bizagiPage.keyboard.press('Enter').catch(() => {});
    await bizagiPage.waitForTimeout(1500);

    // Si Enter no funcionó, intentar con icono de búsqueda
    const triggerBusqueda = inputBuscador
        .locator('xpath=(preceding-sibling::*[contains(@class,"search") or contains(@class,"icon")] | following-sibling::*[contains(@class,"search") or contains(@class,"icon")])[1]')
        .first();

    if (await triggerBusqueda.isVisible().catch(() => false)) {
        console.log(`[DEBUG-BIZAGI][Buscar] Click en icono de búsqueda para ${mpn}`);
        await triggerBusqueda.click({ timeout: 3000 }).catch(() => {});
        await bizagiPage.waitForTimeout(1500);
    }

    console.log(`[DEBUG-BIZAGI][Buscar] Búsqueda completada para ${mpn}`);
}

export async function abrirMpnEnBizagiDebug(bizagiPage: Page, mpn: string): Promise<void> {
    console.log(`[DEBUG-BIZAGI][Abrir] Abriendo fila con MPN=${mpn}`);

    const filasInbox = bizagiPage.locator('table#ui-bizagi-wp-app-inbox-grid-cases tbody tr[data-idworkflow]');

    // Buscar la fila con el MPN, intentando diferentes actividades en orden de preferencia
    const filaGestionCoincidencias = filasInbox
        .filter({ hasText: new RegExp(mpn, 'i') })
        .filter({ hasText: /Gestionar Coincidencias/i })
        .first();

    const filaVerificarPlaft = filasInbox
        .filter({ hasText: new RegExp(mpn, 'i') })
        .filter({ hasText: /Verificar PLAFT/i })
        .first();

    const filaPlaftGenerica = filasInbox
        .filter({ hasText: new RegExp(mpn, 'i') })
        .filter({ hasText: /PLAFT/i })
        .first();

    let fila = filaGestionCoincidencias;
    let tipoActividad = 'Gestionar Coincidencias';
    let visible = await fila.isVisible().catch(() => false);

    if (!visible) {
        fila = filaVerificarPlaft;
        tipoActividad = 'Verificar PLAFT';
        visible = await fila.isVisible().catch(() => false);
    }

    if (!visible) {
        fila = filaPlaftGenerica;
        tipoActividad = 'PLAFT';
        visible = await fila.isVisible().catch(() => false);
    }

    if (!visible) {
        throw new Error(`[DEBUG-BIZAGI][CRITICO] No se encontró fila para MPN=${mpn}`);
    }

    console.log(`[DEBUG-BIZAGI][Abrir] Abriendo ${tipoActividad} para ${mpn}`);
    await fila.scrollIntoViewIfNeeded().catch(() => {});

    // Intentar hacer click en la fila de múltiples formas
    const candidatos = [
        fila.locator('td[class*="RadNumber"] span').first(),
        fila.locator('td[class*="RadNumber"]').first(),
        fila.locator('a').first(),
        fila
    ];

    let clickOk = false;
    for (const candidato of candidatos) {
        if (await candidato.isVisible().catch(() => false)) {
            const ok = await candidato.click({ force: true }).then(() => true).catch(() => false);
            if (ok) {
                clickOk = true;
                break;
            }
        }
    }

    if (!clickOk) {
        throw new Error(`[DEBUG-BIZAGI][CRITICO] No se pudo hacer click en la fila para MPN=${mpn}`);
    }

    await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
    await bizagiPage.waitForTimeout(1500);

    console.log(`[DEBUG-BIZAGI][Abrir] Fila abierta para ${mpn}`);
}

export async function abrirGestionCoincidenciasDebug(bizagiPage: Page): Promise<void> {
    console.log('[DEBUG-BIZAGI][Actividad] Esperando carga de actividad de cumplimiento');

    const sels = [
        'div:has-text("Gestionar Coincidencias")',
        'text=Gestionar Coincidencias',
        'h1:has-text("Gestionar Coincidencias")',
        'span:has-text("Gestionar Coincidencias")',
        '[id*="GestionarCoincidencias"]',
        'div:has-text("Verificar PLAFT")',
        'text=Verificar PLAFT',
        '[id*="VerificarPLAFT"]',
        'div:has-text("PLAFT")',
    ];

    let encontrada = false;
    const tiempoInicio = Date.now();
    const TIMEOUT = 15000;

    while (Date.now() - tiempoInicio < TIMEOUT && !encontrada) {
        for (const sel of sels) {
            if (await bizagiPage.locator(sel).first().isVisible({ timeout: 500 }).catch(() => false)) {
                console.log(`[DEBUG-BIZAGI][Actividad] Actividad detectada con selector: ${sel}`);
                encontrada = true;
                break;
            }
        }
        if (!encontrada) {
            await bizagiPage.waitForTimeout(300);
        }
    }

    if (!encontrada) {
        throw new Error('[DEBUG-BIZAGI][CRITICO] Actividad de cumplimiento no detectada después de 15s');
    }

    console.log('[DEBUG-BIZAGI][Actividad] Actividad de cumplimiento cargada');
}

export async function aprobarOfacGestionCoincidenciasDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][OFAC] Iniciando aprobación OFAC');

    try {
        const tieneOfac = await bizagiPage
            .getByText(/Coincidencias OFAC|Acción Coincidencias OFAC|Motivo Coincidencias OFAC/i)
            .first()
            .isVisible({ timeout: 3000 })
            .catch(() => false);

        if (!tieneOfac) {
            console.log('[DEBUG-BIZAGI][OFAC] No hay sección OFAC visible');
            return false;
        }

        const comboAccionOfac = bizagiPage
            .locator('xpath=(//*[contains(normalize-space(.), "Acción Coincidencias OFAC")]/following::input[@role="combobox"])[1]')
            .first();

        if (await comboAccionOfac.isVisible({ timeout: 2000 }).catch(() => false)) {
            await comboAccionOfac.click({ force: true }).catch(() => {});
            await bizagiPage.waitForTimeout(300);

            const opcionDescartar = bizagiPage
                .locator('li, div[role="option"]')
                .filter({ hasText: /Descartar/ })
                .first();

            if (await opcionDescartar.isVisible({ timeout: 2000 }).catch(() => false)) {
                await opcionDescartar.click({ force: true }).catch(() => {});
                await bizagiPage.waitForTimeout(500);
            }
        }

        const comboMotivo = bizagiPage
            .locator('xpath=(//*[contains(normalize-space(.), "Motivo Coincidencias OFAC")]/following::input[@role="combobox"])[1]')
            .first();

        if (await comboMotivo.isVisible({ timeout: 2000 }).catch(() => false)) {
            await comboMotivo.click({ force: true }).catch(() => {});
            await bizagiPage.waitForTimeout(300);

            const primeraOpcion = bizagiPage.locator('li, div[role="option"]').first();

            if (await primeraOpcion.isVisible({ timeout: 2000 }).catch(() => false)) {
                await primeraOpcion.click({ force: true }).catch(() => {});
                await bizagiPage.waitForTimeout(500);
            }
        }

        console.log('[DEBUG-BIZAGI][OFAC] OFAC completado');
        return true;
    } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        console.log(`[DEBUG-BIZAGI][OFAC][ERROR] ${msg}`);
        return false;
    }
}

export async function completarLexisNexisOtrasCoincidenciasDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][LEXIS] Iniciando procesamiento Lexis Nexis / Otras Coincidencias');

    try {
        const tieneLexisNexis = await bizagiPage
            .getByText(/Lexis\s*Nexis|Otras Coincidencias|Falso Positivo/i)
            .first()
            .isVisible({ timeout: 3000 })
            .catch(() => false);

        if (!tieneLexisNexis) {
            console.log('[DEBUG-BIZAGI][LEXIS] No hay sección Lexis Nexis visible');
            return false;
        }

        const radiosOpciones = bizagiPage.locator('input[type="radio"], input[type="checkbox"]');
        const countOpciones = await radiosOpciones.count().catch(() => 0);

        const falsoPositivo = bizagiPage
            .locator('label, span')
            .filter({ hasText: /Falso Positivo|No es coincidencia/i })
            .first();

        if (await falsoPositivo.isVisible({ timeout: 2000 }).catch(() => false)) {
            console.log('[DEBUG-BIZAGI][LEXIS] Encontrada opción Falso Positivo, seleccionando...');
            await falsoPositivo.click({ force: true }).catch(() => {});
            await bizagiPage.waitForTimeout(500);
        } else if (countOpciones > 0) {
            console.log('[DEBUG-BIZAGI][LEXIS] Seleccionando primera opción disponible');
            const primerRadio = radiosOpciones.first();
            if (await primerRadio.isVisible({ timeout: 1000 }).catch(() => false)) {
                await primerRadio.click({ force: true }).catch(() => {});
                await bizagiPage.waitForTimeout(300);
            }
        }

        console.log('[DEBUG-BIZAGI][LEXIS] Lexis Nexis completado');
        return true;
    } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        console.log(`[DEBUG-BIZAGI][LEXIS][ERROR] ${msg}`);
        return false;
    }
}

export async function completarPlaftGestionCoincidenciasDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][PLAFT] Iniciando procesamiento PLAFT');

    try {
        const tienePlaft = await bizagiPage
            .getByText(/Verificar PLAFT|Debida Diligencia PLAFT/i)
            .first()
            .isVisible({ timeout: 3000 })
            .catch(() => false);

        if (!tienePlaft) {
            console.log('[DEBUG-BIZAGI][PLAFT] No hay sección PLAFT visible');
            return false;
        }

        const combosPlaft = bizagiPage.locator('input[@role="combobox"]');
        const countCombos = await combosPlaft.count().catch(() => 0);

        if (countCombos > 0) {
            console.log(`[DEBUG-BIZAGI][PLAFT] Encontrados ${countCombos} campos combo`);

            for (let i = 0; i < Math.min(2, countCombos); i++) {
                const combo = combosPlaft.nth(i);
                if (await combo.isVisible({ timeout: 1000 }).catch(() => false)) {
                    await combo.click({ force: true }).catch(() => {});
                    await bizagiPage.waitForTimeout(300);

                    const primeraOpcion = bizagiPage.locator('li, div[role="option"]').first();
                    if (await primeraOpcion.isVisible({ timeout: 1000 }).catch(() => false)) {
                        await primeraOpcion.click({ force: true }).catch(() => {});
                        await bizagiPage.waitForTimeout(300);
                    }
                }
            }
        }

        console.log('[DEBUG-BIZAGI][PLAFT] PLAFT completado');
        return true;
    } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        console.log(`[DEBUG-BIZAGI][PLAFT][ERROR] ${msg}`);
        return false;
    }
}

export async function aprobarGestionCoincidenciasMixtaDebug(
    bizagiPage: Page,
    modo: 'OFAC' | 'PLAFT' | 'LEXIS' | 'MIXTO'
): Promise<void> {
    console.log(`[DEBUG-BIZAGI][Modo] Ejecutando modo=${modo}`);

    let resultOfac = true;
    let resultPlaft = true;
    let resultLexis = true;

    // En modo MIXTO, intentamos ejecutar todo pero no fallamos si algo no existe
    // En modo específico, si existe la sección pero falla, entonces es error
    if (modo === 'MIXTO' || modo === 'OFAC') {
        resultOfac = await aprobarOfacGestionCoincidenciasDebug(bizagiPage);
        if (!resultOfac && modo === 'OFAC') {
            // Si es OFAC puro y la sección no existe, intentar con PLAFT
            console.log('[DEBUG-BIZAGI][Modo] OFAC no encontrado, intentando PLAFT...');
            resultPlaft = await completarPlaftGestionCoincidenciasDebug(bizagiPage);
            if (!resultPlaft) {
                throw new Error('[DEBUG-BIZAGI][CRITICO] Ni OFAC ni PLAFT encontrados');
            }
        }
    }

    if (modo === 'MIXTO' || modo === 'LEXIS') {
        resultLexis = await completarLexisNexisOtrasCoincidenciasDebug(bizagiPage);
        if (!resultLexis && modo === 'LEXIS') {
            // Si es LEXIS puro y no existe, intentar con PLAFT
            console.log('[DEBUG-BIZAGI][Modo] LEXIS no encontrado, intentando PLAFT...');
            resultPlaft = await completarPlaftGestionCoincidenciasDebug(bizagiPage);
            if (!resultPlaft) {
                throw new Error('[DEBUG-BIZAGI][CRITICO] Ni LEXIS ni PLAFT encontrados');
            }
        }
    }

    if (modo === 'PLAFT') {
        resultPlaft = await completarPlaftGestionCoincidenciasDebug(bizagiPage);
        if (!resultPlaft) {
            throw new Error('[DEBUG-BIZAGI][CRITICO] Modo PLAFT falló');
        }
    }

    console.log(`[DEBUG-BIZAGI][Modo] Modos completados: OFAC=${resultOfac}, LEXIS=${resultLexis}, PLAFT=${resultPlaft}`);
}

export async function resolverModoBizagiDebug(bizagiPage: Page): Promise<'OFAC' | 'PLAFT' | 'LEXIS' | 'MIXTO'> {
    console.log('[DEBUG-BIZAGI][Auto] Resolviendo modo automático...');

    const pageText = await bizagiPage.locator('body').innerText({ timeout: 3000 }).catch(() => '');

    // Detecciones flexibles con múltiples variantes
    const tieneGestionCoincidencias = /Gesti[oó]n.?de.?Coincidencias|Gestion.?de.?Coincidencias|Gestionar.*Coincidencias/i.test(pageText);
    const tieneVerificarPlaft = /Verificar.?PLAFT|Debida.?Diligencia.?PLAFT|PLAFT|Revisión.*PLAFT/i.test(pageText);
    const tieneLexisNexis = /Lexis\s*Nexis|Otras\s*Coincidencias|Falso\s*Positivo/i.test(pageText);
    const tieneOfac = /OFAC|Listas.?OFAC|Coincidencias.?OFAC/i.test(pageText);

    console.log(`[DEBUG-BIZAGI][Auto] Detectados: GestionCoincidencias=${tieneGestionCoincidencias}, VerificarPlaft=${tieneVerificarPlaft}, LexisNexis=${tieneLexisNexis}, OFAC=${tieneOfac}`);

    let modoResuelto: 'OFAC' | 'PLAFT' | 'LEXIS' | 'MIXTO' = 'MIXTO';

    if (tieneGestionCoincidencias && tieneVerificarPlaft) {
        modoResuelto = 'MIXTO';
        console.log('[DEBUG-BIZAGI][Auto] Ambas actividades detectadas => MIXTO');
    } else if (tieneGestionCoincidencias) {
        if (tieneLexisNexis) {
            modoResuelto = 'LEXIS';
            console.log('[DEBUG-BIZAGI][Auto] Gestionar Coincidencias con Lexis => LEXIS');
        } else if (tieneOfac) {
            modoResuelto = 'OFAC';
            console.log('[DEBUG-BIZAGI][Auto] Gestionar Coincidencias con OFAC => OFAC');
        } else {
            modoResuelto = 'MIXTO';
            console.log('[DEBUG-BIZAGI][Auto] Gestionar Coincidencias (ambigua) => MIXTO');
        }
    } else if (tieneVerificarPlaft) {
        modoResuelto = 'PLAFT';
        console.log('[DEBUG-BIZAGI][Auto] Verificar PLAFT detectado => PLAFT');
    }

    console.log(`[DEBUG-BIZAGI][Auto] Modo resuelto: ${modoResuelto}`);
    return modoResuelto;
}

export async function guardarYSiguienteDebug(bizagiPage: Page): Promise<void> {
    console.log('[DEBUG-BIZAGI][Guardar] Buscando botones Guardar/Siguiente');

    const btnGuardar = bizagiPage.getByRole('button', { name: /Guardar/i }).first();
    const guardaVisible = await btnGuardar.isVisible({ timeout: 2000 }).catch(() => false);

    if (guardaVisible) {
        console.log('[DEBUG-BIZAGI][Guardar] Click en Guardar');
        await btnGuardar.click({ force: true }).catch(() => {});
        await bizagiPage.waitForTimeout(1000);
    }

    const btnSiguiente = bizagiPage.getByRole('button', { name: /Siguiente/i }).first();
    const siguienteVisible = await btnSiguiente.isVisible({ timeout: 2000 }).catch(() => false);

    if (siguienteVisible) {
        console.log('[DEBUG-BIZAGI][Guardar] Click en Siguiente');
        await btnSiguiente.click({ force: true }).catch(() => {});
        await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
        await bizagiPage.waitForTimeout(1500);
    }

    console.log('[DEBUG-BIZAGI][Guardar] Guardado/Siguiente completado');
}
