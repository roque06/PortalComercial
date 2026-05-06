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
  const selectorMenuBandeja = '#mnuBt\\.Inbox, [id*="mnuBt"][id*="Inbox"], li.ui-bizagi-wp-app-menu-cases, a:has-text("Bandeja"), a:has-text("Inbox")';
  const botonBandeja = await findVisibleLocator(bizagiPage, selectorMenuBandeja.split(','), 'BandejaBtn');
  if (await botonBandeja.isVisible().catch(() => false)) {
    await botonBandeja.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(800);
  }
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
// HELPERS PARA DETECCIÓN DE ESTADO
// ────────────────────────────────────

function normalizarTextoCombo(valor: string | null | undefined): string {
  return (valor || '').replace(/\s+/g, ' ').trim();
}

function esAccionLexisFalsoPositivo(estado: string): boolean {
  return /Falso\s+Positivo/i.test(estado || '');
}

function esAccionLexisPendiente(estado: string): boolean {
  const t = normalizarTextoCombo(estado);
  if (esAccionLexisFalsoPositivo(t)) return false;
  return !t || /Por favor seleccione/i.test(t) || /^-+$/.test(t) || /Seleccione/i.test(t) || /opt-/i.test(t);
}

function cssEscapeDebug(value: string): string {
  return (value || '').replace(/([ #;?%&,.+*~':"!^$[\]()=>|/@])/g, '\\$1');
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

    const esperarHomeBizagi = async (): Promise<boolean> => {
        const deadline = Date.now() + 12000;
        while (Date.now() < deadline) {
            if (await esBizagiHomePage(bizagiPage)) return true;

            const menuQuery = await bizagiPage.locator('#menuQuery, input[name="menuQuery"]').first().isVisible().catch(() => false);
            const inbox = await bizagiPage.locator('table#ui-bizagi-wp-app-inbox-grid-cases, [id*="menuListInbox"]').first().isVisible().catch(() => false);
            const bandeja = await bizagiPage.getByText(/Bandeja de entrada|Inbox/i).first().isVisible().catch(() => false);
            if (menuQuery || inbox || bandeja) return true;
            await bizagiPage.waitForTimeout(400).catch(() => {});
        }
        return false;
    };

    const diagnosticarPostLogin = async () => {
        const url = bizagiPage.url();
        const title = await bizagiPage.title().catch(() => '');
        const bodyPreview = ((await bizagiPage.locator('body').innerText().catch(() => '')) || '')
            .replace(/\s+/g, ' ')
            .trim()
            .slice(0, 400);
        const inputsVisibles = await bizagiPage.locator('input:visible').count().catch(() => 0);
        const existeMenuQuery = await bizagiPage.locator('#menuQuery').first().isVisible().catch(() => false);
        const existeInbox = await bizagiPage.locator('table#ui-bizagi-wp-app-inbox-grid-cases').first().isVisible().catch(() => false);

        console.log(`[DEBUG-BIZAGI][Login][Diag] url=${url}`);
        console.log(`[DEBUG-BIZAGI][Login][Diag] title=${title}`);
        console.log(`[DEBUG-BIZAGI][Login][Diag] bodyPreview=${bodyPreview}`);
        console.log(`[DEBUG-BIZAGI][Login][Diag] inputsVisibles=${inputsVisibles} menuQuery=${existeMenuQuery} inbox=${existeInbox}`);
    };

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

        const selectUsuario = await pickVisible(bizagiPage, [
            bizagiPage.locator('select#user').first(),
            bizagiPage.locator('select[name="user"]').first(),
            bizagiPage.locator('select[name="username"]').first(),
            bizagiPage.locator('select:visible').first(),
        ]);

        const uVis = await inputUser.isVisible().catch(() => false);
        const pVis = await inputPass.isVisible().catch(() => false);
        const suVis = await selectUsuario.isVisible().catch(() => false);
        const iVis = await btnIngresar.isVisible().catch(() => false);

        if (!uVis && !pVis && !suVis && !iVis) {
            await bizagiPage.waitForTimeout(600).catch(() => {});
            continue;
        }

        console.log(`[DEBUG-BIZAGI][Login] Detectado U=${uVis}, P=${pVis}, SU=${suVis}, I=${iVis}`);

        if (suVis && iVis && !uVis && !pVis) {
            console.log('[DEBUG-BIZAGI][Login] Login solo-selector detectado');
            console.log('[DEBUG-BIZAGI][Login] Seleccionando usuario admon del combo...');

            const opcionesUsuario = await selectUsuario.evaluate((sel: HTMLSelectElement) =>
                Array.from(sel.options).map((o, idx) => ({
                    idx,
                    value: (o.value || '').trim(),
                    text: (o.textContent || '').trim(),
                    selected: o.selected,
                }))
            ).catch(() => [] as Array<{ idx: number; value: string; text: string; selected: boolean }>);
            console.log(`[DEBUG-BIZAGI][Login] opciones usuario=${JSON.stringify(opcionesUsuario)}`);

            const userEnv = process.env.BIZAGI_USER || usuario || '';
            const normalizar = (s: string) => s.toLowerCase().trim().replace(/\\\\/g, '\\').replace(/\s+/g, ' ');
            const userNorm = normalizar(userEnv);
            const userSimple = userNorm.includes('\\') ? userNorm.split('\\').pop() || userNorm : userNorm;

            const esInvalida = (value: string, text: string) => {
                const v = normalizar(value || '');
                const t = normalizar(text || '');
                return !v || !t || /seleccione/.test(t) || /^-+$/.test(t);
            };

            const pick = (pred: (o: { idx: number; value: string; text: string }) => boolean) =>
                opcionesUsuario.find((o) => !esInvalida(o.value, o.text) && pred(o));

            let candidata =
                (userNorm ? pick((o) => normalizar(o.value) === userNorm || normalizar(o.text) === userNorm) : undefined) ||
                (userSimple ? pick((o) => normalizar(o.value).includes(userSimple) || normalizar(o.text).includes(userSimple)) : undefined) ||
                pick((o) => /admon/.test(normalizar(o.value)) || /admon/.test(normalizar(o.text))) ||
                opcionesUsuario.find((o) => !esInvalida(o.value, o.text));

            if (!candidata) {
                throw new Error('[DEBUG-BIZAGI][Login][CRITICO] No se encontró opción válida de usuario admon en el selector');
            }

            await selectUsuario.selectOption({ index: candidata.idx }).catch(async () => {
                await selectUsuario.evaluate((sel, idx) => {
                    const s = sel as HTMLSelectElement;
                    if (s.options[idx]) {
                        s.selectedIndex = idx;
                        s.dispatchEvent(new Event('input', { bubbles: true }));
                        s.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }, candidata.idx).catch(() => {});
            });

            const seleccionadoFinal = await selectUsuario.evaluate((sel: HTMLSelectElement) => {
                const opt = sel.options[sel.selectedIndex];
                return {
                    value: (sel.value || '').trim(),
                    text: (opt?.textContent || '').trim(),
                    selectedIndex: sel.selectedIndex,
                };
            }).catch(() => ({ value: '', text: '', selectedIndex: -1 }));

            console.log(`[DEBUG-BIZAGI][Login] usuario seleccionado final=${JSON.stringify(seleccionadoFinal)}`);
            if (!seleccionadoFinal.value || !seleccionadoFinal.text || /seleccione/i.test(seleccionadoFinal.text)) {
                throw new Error('[DEBUG-BIZAGI][Login][CRITICO] Usuario seleccionado quedó vacío');
            }

            console.log('[DEBUG-BIZAGI][Login] Click Ingresar (selector)');

            await btnIngresar.click({ force: true }).catch(() => {});
            await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
            await bizagiPage.waitForTimeout(2200);

            console.log('[DEBUG-BIZAGI][Login] Esperando home después de login selector');
            const homeOk = await esperarHomeBizagi();
            console.log(`[DEBUG-BIZAGI][Login] homeDetectado=${homeOk}`);
            if (!homeOk) {
                await diagnosticarPostLogin();
                throw new Error('[DEBUG-BIZAGI][CRITICO] Login solo-selector no llegó a home/buscador');
            }

            console.log('[DEBUG-BIZAGI][Login] Login solo-selector exitoso');
            return true;
        }

        if (uVis && pVis && iVis) {
            console.log('[DEBUG-BIZAGI][Login] Login usuario/password detectado');

            console.log(`[DEBUG-BIZAGI][Login] Llenando usuario: ${usuario}`);
            await inputUser.fill(usuario).catch(() => {});
            await bizagiPage.waitForTimeout(300);

            console.log('[DEBUG-BIZAGI][Login] Llenando contraseña');
            await inputPass.fill(password).catch(() => {});
            await bizagiPage.waitForTimeout(300);

            console.log('[DEBUG-BIZAGI][Login] Haciendo click en Ingresar (usuario/password)');
            await btnIngresar.click({ force: true }).catch(() => {});
            await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
            await bizagiPage.waitForTimeout(2000);

            const homeOk = await esBizagiHomePage(bizagiPage);
            if (!homeOk) {
                await diagnosticarPostLogin();
                throw new Error('[DEBUG-BIZAGI][CRITICO] Login usuario/password no llegó a home/buscador');
            }

            console.log('[DEBUG-BIZAGI][Login] Login usuario/password completado');
            return true;
        }

        if (iVis && !uVis && !pVis && !suVis) {
            console.log('[DEBUG-BIZAGI][Login][WARN] Solo botón Ingresar visible (sin U/P/SU). No se hace click prematuro.');
            await bizagiPage.waitForTimeout(800).catch(() => {});
            continue;
        }

        await bizagiPage.waitForTimeout(600).catch(() => {});
    }

    throw new Error('[DEBUG-BIZAGI][CRITICO] No se detectó login ni sesión activa después de 45s');
}

export async function buscarMpnEnBizagiDebug(bizagiPage: Page, mpn: string): Promise<void> {
    console.log(`[DEBUG-BIZAGI][Buscar] Buscando MPN=${mpn}`);

    await navegarABandejaEntradaBizagi(bizagiPage).catch(() => false);

    const selectoresBuscador = [
        '#menuQuery',
        'input#menuQuery',
        '#ui-bizagi-wp-widget-searchContainer input',
        'input[name="menuQuery"]',
        'input[placeholder="Buscar"]',
    ];

    const inputBuscadorDetectado = await findVisibleLocator(bizagiPage, selectoresBuscador, 'BuscadorSuperior');

    if (await inputBuscadorDetectado.isVisible().catch(() => false)) {
        console.log(`[DEBUG-BIZAGI][Buscar] buscando ${mpn} en el buscador superior detectado...`);

        await cerrarMenusBizagi(bizagiPage).catch(() => {});
        await inputBuscadorDetectado.focus().catch(() => {});
        await inputBuscadorDetectado.fill(mpn).catch(() => {});
        await bizagiPage.waitForTimeout(300);

        // Disparar búsqueda
        await inputBuscadorDetectado.evaluate((el) => {
            const inputEl = el as HTMLInputElement;
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            inputEl.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Enter', code: 'Enter' }));
        }).catch(() => {});

        await inputBuscadorDetectado.press('Enter').catch(() => {});
        await bizagiPage.waitForTimeout(1000);

        // Buscar y click en fila
        const filasConMpn = bizagiPage.locator(`tr:has-text("${mpn}"), div:has-text("${mpn}")`);
        const countFilas = await filasConMpn.count().catch(() => 0);

        if (countFilas > 0) {
            console.log(`[DEBUG-BIZAGI][Buscar] fila con ${mpn} encontrada`);
            const filaConMpn = filasConMpn.first();

            // Buscar icono de búsqueda
            const triggerBusqueda = inputBuscadorDetectado
                .locator('xpath=(preceding-sibling::*[contains(@class,"search") or contains(@class,"icon")] | following-sibling::*[contains(@class,"search") or contains(@class,"icon")])[1]')
                .first();

            if (await triggerBusqueda.isVisible().catch(() => false)) {
                console.log(`[DEBUG-BIZAGI][Buscar] Click en icono de busqueda para ${mpn}`);
                await triggerBusqueda.click({ timeout: 3000 }).catch(() => {});
            }

            await bizagiPage.waitForTimeout(1500);
        } else {
            throw new Error(`[DEBUG-BIZAGI][CRITICO] MPN ${mpn} no encontrado en resultados`);
        }
    } else {
        throw new Error(`[DEBUG-BIZAGI][CRITICO] No se encontró buscador superior`);
    }
}

async function pantallaGestionCoincidenciasAbiertaDebug(page: Page): Promise<{
    abierta: boolean;
    body: string;
    gestionVisible: boolean;
    ofacVisible: boolean;
    otrasVisible: boolean;
    lexisVisible: boolean;
}> {
    const body = await page.locator('body').innerText({ timeout: 3000 }).catch(() => '');

    const gestionVisible = /Gestionar\s+Coincidencias|Gesti[oó]n\s+de\s+Coincidencias/i.test(body);
    const ofacVisible = /Coincidencias\s+OFAC|Acci[oó]n\s+Coincidencias\s+OFAC|Motivo\s+Coincidencias\s+OFAC/i.test(body);
    const otrasVisible = /Otras\s+Coincidencias|Falso\s+Positivo|No\s+es\s+coincidencia/i.test(body);
    const lexisVisible = /Lexis\s+Nexis/i.test(body);

    const abierta = gestionVisible || ofacVisible || otrasVisible || lexisVisible;

    return { abierta, body, gestionVisible, ofacVisible, otrasVisible, lexisVisible };
}

export async function abrirFilaGestionCoincidenciasDebug(bizagiPage: Page, mpn: string): Promise<Page> {
    console.log(`[DEBUG-BIZAGI][Abrir] Abriendo fila con MPN=${mpn}`);

    const mpnEsc = mpn.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const reMpn = new RegExp(mpnEsc, 'i');

    const filas = bizagiPage
        .locator('table#ui-bizagi-wp-app-inbox-grid-cases tbody tr[data-idworkflow], table#ui-bizagi-wp-app-inbox-grid-cases tbody tr, table tbody tr, tr')
        .filter({ hasText: reMpn });

    const totalFilas = await filas.count().catch(() => 0);
    console.log(`[DEBUG-BIZAGI][Abrir][Diag] filas candidatas por MPN=${totalFilas}`);
    for (let i = 0; i < Math.min(totalFilas, 5); i++) {
        const txt = (await filas.nth(i).innerText().catch(() => '')).replace(/\s+/g, ' ').trim().slice(0, 260);
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] fila[${i}]='${txt}'`);
    }

    if (totalFilas === 0) {
        throw new Error(`[DEBUG-BIZAGI][CRITICO] No se encontró fila para ${mpn}`);
    }

    const filasGestion = filas.filter({ hasText: /Gestionar\s+Coincidencias/i });
    const countGestion = await filasGestion.count().catch(() => 0);
    if (countGestion === 0) {
        throw new Error(`[DEBUG-BIZAGI][Abrir][CRITICO] No existe fila Gestionar Coincidencias para ${mpn}`);
    }
    const fila = filasGestion.first();

    const textoFilaSeleccionada = (await fila.innerText().catch(() => '')).replace(/\s+/g, ' ').trim();
    console.log(`[DEBUG-BIZAGI][Abrir] fila seleccionada origen='Gestionar Coincidencias'`);
    console.log(`[DEBUG-BIZAGI][Abrir] fila seleccionada texto='${textoFilaSeleccionada.slice(0, 280)}'`);

    console.log(`[DEBUG-BIZAGI][Abrir] Abriendo Gestionar Coincidencias para ${mpn}`);

    const pantallaAbierta = async () => await pantallaGestionCoincidenciasAbiertaDebug(bizagiPage);

    const detectarPageConGestion = async (): Promise<Page | null> => {
        const pages = bizagiPage.context().pages();
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] paginas abiertas=${pages.length}`);
        for (let i = 0; i < pages.length; i++) {
            const p = pages[i];
            const url = p.url();
            const body = (await p.locator('body').innerText().catch(() => '')).replace(/\s+/g, ' ').trim();
            const bodyTieneGestion = /Lexis\s+Nexis|Otras\s+Coincidencias|Coincidencias\s+OFAC|Acci[oó]n\s+Coincidencias\s+OFAC|Motivo\s+Coincidencias\s+OFAC|Gestionar\s+Coincidencias|Gesti[oó]n\s+de\s+Coincidencias/i.test(body);
            console.log(`[DEBUG-BIZAGI][Abrir][Diag] page[${i}] url='${url}' bodyTieneGestion=${bodyTieneGestion}`);
            if (bodyTieneGestion) {
                console.log(`[DEBUG-BIZAGI][Abrir] actividad encontrada en page index=${i}`);
                return p;
            }
        }
        return null;
    };

    const postIntent = async (name: string): Promise<Page | null> => {
        await bizagiPage.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {});
        await bizagiPage.waitForTimeout(1500);
        const estado = await pantallaAbierta();
        const body = (estado.body || '').replace(/\s+/g, ' ').trim();
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] intento=${name} abierto=${estado.abierta}`);
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] post-click bodyTieneGestion=${estado.abierta}`);
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] bodyPreview='${body.slice(0, 260)}'`);
        if (estado.abierta) return bizagiPage;
        return await detectarPageConGestion();
    };

    let pageAbierta: Page | null = null;

    // 1) RadNumber
    const radNumber = fila.locator('td[class*="RadNumber"] span, td[class*="RadNumber"] a').first();
    console.log('[DEBUG-BIZAGI][Abrir][Diag] intento=RadNumber (ejecutando)');
    if (await radNumber.isVisible().catch(() => false)) {
        await radNumber.scrollIntoViewIfNeeded().catch(() => {});
        await radNumber.click({ force: true }).catch(() => {});
    }
    pageAbierta = await postIntent('RadNumber');

    // 2) activity-name
    if (!pageAbierta) {
        const activityName = fila.locator('.ui-bizagi-wp-app-inbox-activity-name').first();
        console.log('[DEBUG-BIZAGI][Abrir][Diag] intento=activity-name (ejecutando)');
        if (await activityName.isVisible().catch(() => false)) {
            await activityName.click({ force: true }).catch(() => {});
        }
        pageAbierta = await postIntent('activity-name');
    }

    // 3) td texto Gestionar
    if (!pageAbierta) {
        const tdGestionar = fila.locator('td').filter({ hasText: /Gestionar\s+Coincidencias/i }).first();
        console.log('[DEBUG-BIZAGI][Abrir][Diag] intento=text-Gestionar (ejecutando)');
        if (await tdGestionar.isVisible().catch(() => false)) {
            await tdGestionar.click({ force: true }).catch(() => {});
        }
        pageAbierta = await postIntent('text-Gestionar');
    }

    // 4) span texto Gestionar
    if (!pageAbierta) {
        const spanGestionar = fila.locator('span').filter({ hasText: /Gestionar\s+Coincidencias/i }).first();
        if (await spanGestionar.isVisible().catch(() => false)) {
            await spanGestionar.click({ force: true }).catch(() => {});
        }
        pageAbierta = await postIntent('span-Gestionar');
    }

    // 5) getByText Gestionar
    if (!pageAbierta) {
        const actividad = fila.getByText(/Gestionar\s+Coincidencias/i).first();
        if (await actividad.isVisible().catch(() => false)) {
            await actividad.click({ force: true }).catch(() => {});
        }
        pageAbierta = await postIntent('getByText-Gestionar');
    }

    // 6) primer td visible
    if (!pageAbierta) {
        const primerTd = fila.locator('td:visible').first();
        if (await primerTd.isVisible().catch(() => false)) {
            await primerTd.click({ force: true }).catch(() => {});
        }
        pageAbierta = await postIntent('primer-td');
    }

    // 7) fila click
    if (!pageAbierta) {
        await fila.click({ force: true }).catch(() => {});
        pageAbierta = await postIntent('fila-click');
    }

    // 8) fila double click
    if (!pageAbierta) {
        console.log('[DEBUG-BIZAGI][Abrir][Diag] intento=fila-dblclick (ejecutando)');
        await fila.dblclick({ force: true }).catch(() => {});
        pageAbierta = await postIntent('fila-dblclick');
    }

    // 9) Enter sobre fila
    if (!pageAbierta) {
        await fila.focus().catch(() => {});
        await bizagiPage.keyboard.press('Enter').catch(() => {});
        pageAbierta = await postIntent('fila-enter');
    }

    // 10) Coordenadas centro de fila
    if (!pageAbierta) {
        console.log('[DEBUG-BIZAGI][Abrir][Diag] intento=coordenadas (ejecutando)');
        const box = await fila.boundingBox().catch(() => null);
        if (box) {
            await bizagiPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2).catch(() => {});
        }
        pageAbierta = await postIntent('coordenadas');
    }

    if (!pageAbierta) {
        const url = bizagiPage.url();
        const title = await bizagiPage.title().catch(() => '');
        const bodyPreview = ((await bizagiPage.locator('body').innerText().catch(() => '')) || '')
            .replace(/\s+/g, ' ')
            .trim()
            .slice(0, 450);
        const screenshotPath = `artifacts/evidencias_tmp/debug-bizagi-no-abrio-gestion-${mpn}.png`;
        await bizagiPage.screenshot({ path: screenshotPath }).catch(() => {});
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] url=${url}`);
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] title=${title}`);
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] bodyPreview=${bodyPreview}`);
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] screenshot=${screenshotPath}`);
        const pagesCount = bizagiPage.context().pages().length;
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] pagesCount=${pagesCount}`);
        console.log(`[DEBUG-BIZAGI][Abrir][Diag] filaSeleccionada='${textoFilaSeleccionada.slice(0, 280)}'`);
        throw new Error(`[DEBUG-BIZAGI][CRITICO] No se abrió Gestionar Coincidencias para ${mpn}`);
    }

    console.log(`[DEBUG-BIZAGI][Abrir] Fila abierta para ${mpn}`);
    return pageAbierta;
}

export async function abrirGestionCoincidenciasDebug(bizagiPage: Page): Promise<void> {
    console.log('[DEBUG-BIZAGI][Actividad] Verificando acceso a Gestionar Coincidencias');

    const estado = await pantallaGestionCoincidenciasAbiertaDebug(bizagiPage);
    const bodyPreview = (estado.body || '').replace(/\s+/g, ' ').trim().slice(0, 350);

    console.log(`[DEBUG-BIZAGI][Actividad][Diag] abierta=${estado.abierta} gestionVisible=${estado.gestionVisible} ofacVisible=${estado.ofacVisible} otrasVisible=${estado.otrasVisible} lexisVisible=${estado.lexisVisible}`);
    console.log(`[DEBUG-BIZAGI][Actividad][Diag] bodyPreview='${bodyPreview}'`);

    if (!estado.abierta) {
        throw new Error('[DEBUG-BIZAGI][CRITICO] Pantalla Gestionar Coincidencias no detectada');
    }

    console.log('[DEBUG-BIZAGI][Actividad] Pantalla Gestionar Coincidencias detectada');
}

async function seleccionarAccionOfacDescartarDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][OFAC] Buscando Acción Coincidencias OFAC');

    const accionDiag = await bizagiPage.evaluate(() => {
        const norm = (s?: string | null) => (s || '').replace(/\s+/g, ' ').trim();
        const visibles = (els: Element[]) => els.filter((el) => {
            const h = el as HTMLElement;
            const r = h.getBoundingClientRect();
            const st = window.getComputedStyle(h);
            return r.width > 0 && r.height > 0 && st.visibility !== 'hidden' && st.display !== 'none';
        });
        const labelsAccion = visibles(Array.from(document.querySelectorAll('label, span, div, td, th')))
            .filter((el) => /Acci[oó]n\s+Coincidencias\s+OFAC/i.test(norm((el as HTMLElement).innerText)));
        const opcionesConfirmar = visibles(Array.from(document.querySelectorAll('label, span, div')))
            .filter((el) => /^Confirmar$/i.test(norm((el as HTMLElement).innerText)));
        const opcionesDescartar = visibles(Array.from(document.querySelectorAll('label, span, div')))
            .filter((el) => /^Descartar$/i.test(norm((el as HTMLElement).innerText)));
        const radios = visibles(Array.from(document.querySelectorAll('input[type="radio"]')));

        const diagTextos = [...labelsAccion.slice(0, 6), ...opcionesConfirmar.slice(0, 6), ...opcionesDescartar.slice(0, 6)].map((el) => {
            const h = el as HTMLElement;
            const r = h.getBoundingClientRect();
            const cercano = h.closest('label, div, td, tr, section')?.querySelector('input[type="radio"]') as HTMLInputElement | null;
            return {
                texto: norm(h.innerText).slice(0, 140),
                tag: h.tagName,
                id: h.id || '',
                className: h.className || '',
                x: Math.round(r.left),
                y: Math.round(r.top),
                w: Math.round(r.width),
                h: Math.round(r.height),
                inputCercano: cercano ? `${cercano.id || ''}|${cercano.name || ''}|${cercano.value || ''}` : '',
            };
        });

        const diagRadios = radios.slice(0, 20).map((el, i) => {
            const rEl = el as HTMLInputElement;
            const r = rEl.getBoundingClientRect();
            return {
                i,
                checked: !!rEl.checked,
                value: rEl.value || '',
                id: rEl.id || '',
                name: rEl.name || '',
                x: Math.round(r.left),
                y: Math.round(r.top),
            };
        });

        return {
            labelsAccion: labelsAccion.length,
            opcionesConfirmar: opcionesConfirmar.length,
            opcionesDescartar: opcionesDescartar.length,
            radioInputs: radios.length,
            diagTextos,
            diagRadios,
        };
    }).catch(() => ({ labelsAccion: 0, opcionesConfirmar: 0, opcionesDescartar: 0, radioInputs: 0, diagTextos: [], diagRadios: [] as any[] }));

    console.log(`[DEBUG-BIZAGI][OFAC][AccionDiag] labelsAccion=${accionDiag.labelsAccion}`);
    console.log(`[DEBUG-BIZAGI][OFAC][AccionDiag] opcionesConfirmar=${accionDiag.opcionesConfirmar}`);
    console.log(`[DEBUG-BIZAGI][OFAC][AccionDiag] opcionesDescartar=${accionDiag.opcionesDescartar}`);
    console.log(`[DEBUG-BIZAGI][OFAC][AccionDiag] radioInputs=${accionDiag.radioInputs}`);
    for (const r of accionDiag.diagRadios) {
        console.log(`[DEBUG-BIZAGI][OFAC][AccionDiag] radio[${r.i}] checked=${r.checked} value='${r.value}' id='${r.id}' name='${r.name}' x=${r.x} y=${r.y}`);
    }

    const labelAccion = bizagiPage.getByText(/Acci[oó]n\s+Coincidencias\s+OFAC/i).first();
    const labelVisible = await labelAccion.isVisible({ timeout: 2500 }).catch(() => false);
    if (!labelVisible) {
        throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] No se seleccionó Acción Coincidencias OFAC=Descartar');
    }

    console.log('[DEBUG-BIZAGI][OFAC] seleccionando Acción=Descartar');

    let accionDescartar = false;
    console.log('[DEBUG-BIZAGI][OFAC] radios aclaraciones ignorados=true');

    async function validarAccionOfacDescartarDebug(page: Page): Promise<boolean> {
        const radioDirecto = page.locator('input[type="radio"][id*="sidP_AccOFAC-2"], input[type="radio"][id*="AccOFAC"][value="2"], input[type="radio"][id*="OFAC"][value="2"]').first();
        const radioDirectoExiste = (await radioDirecto.count().catch(() => 0)) > 0;
        if (radioDirectoExiste) {
            const radioChecked = await radioDirecto.isChecked().catch(() => false);
            console.log(`[DEBUG-BIZAGI][OFAC] validarDescartar metodo='id-value' radioChecked=${radioChecked}`);
            if (radioChecked) {
                console.log("[DEBUG-BIZAGI][OFAC] validarDescartar ariaChecked=false");
                console.log("[DEBUG-BIZAGI][OFAC] validarDescartar textoAccion='Descartar'");
                return true;
            }
        }

        const v = await page.evaluate(() => {
            const norm = (s?: string | null) => (s || '').replace(/\s+/g, ' ').trim();
            const radios = Array.from(document.querySelectorAll('input[type="radio"]')) as HTMLInputElement[];
            const textosDescartar = Array.from(document.querySelectorAll('label, span, div'))
                .filter((el) => /^Descartar$/i.test(norm((el as HTMLElement).innerText))) as HTMLElement[];

            let radioChecked = false;
            for (const t of textosDescartar) {
                const tb = t.getBoundingClientRect();
                for (const r of radios) {
                    const rb = r.getBoundingClientRect();
                    const dy = Math.abs((rb.top + rb.height / 2) - (tb.top + tb.height / 2));
                    const dx = (tb.left - (rb.left + rb.width / 2));
                    if (dy <= 20 && dx >= -8 && dx <= 120 && r.checked) {
                        radioChecked = true;
                    }
                }
            }

            const ariaChecked = textosDescartar.some((t) => {
                const c = t.closest('label, div, td, tr, section');
                if (!c) return false;
                if (c.getAttribute('aria-checked') === 'true') return true;
                return Array.from(c.querySelectorAll('[aria-checked="true"]')).some((el) => /descartar/i.test((el as HTMLElement).innerText || ''));
            });

            const textoAccion = (() => {
                const acc = Array.from(document.querySelectorAll('label, span, div, td')).find((el) => /Acci[oó]n\s+Coincidencias\s+OFAC/i.test(norm((el as HTMLElement).innerText)));
                const scope = (acc?.closest('tr, section, div, table') as HTMLElement | null) || document.body;
                return norm(scope.innerText).slice(0, 220);
            })();

            return { radioChecked, ariaChecked, textoAccion };
        }).catch(() => ({ radioChecked: false, ariaChecked: false, textoAccion: '' }));

        console.log(`[DEBUG-BIZAGI][OFAC] validarDescartar radioChecked=${v.radioChecked}`);
        console.log(`[DEBUG-BIZAGI][OFAC] validarDescartar ariaChecked=${v.ariaChecked}`);
        console.log(`[DEBUG-BIZAGI][OFAC] validarDescartar textoAccion='${v.textoAccion}'`);
        return !!(v.radioChecked || v.ariaChecked || /Descartar/i.test(v.textoAccion));
    }

    const contenedor = labelAccion.locator('xpath=ancestor::*[self::div or self::tr or self::td or self::section][1]').first();
    const descartarLabel = contenedor.getByText(/^Descartar$/i).first();

    const radioDescartarDirecto = bizagiPage.locator([
        'input[type="radio"][id*="sidP_AccOFAC-2"]',
        'input[type="radio"][id*="AccOFAC"][value="2"]',
        'input[type="radio"][id*="OFAC"][value="2"]',
    ].join(', ')).first();
    const radioDescartarDirectoExiste = (await radioDescartarDirecto.count().catch(() => 0)) > 0;
    console.log(`[DEBUG-BIZAGI][OFAC] radio Descartar directo por id/value localizado=${radioDescartarDirectoExiste}`);

    if (radioDescartarDirectoExiste) {
        console.log('[DEBUG-BIZAGI][OFAC] click radio OFAC Descartar por id/value');
        await radioDescartarDirecto.click({ force: true }).catch(() => {});
        await bizagiPage.waitForTimeout(300);
        accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
        if (!accionDescartar) {
            console.log('[DEBUG-BIZAGI][OFAC][FallbackDOM] marcando radio OFAC Descartar por id/value');
            await radioDescartarDirecto.evaluate((el) => {
                const i = el as HTMLInputElement;
                i.checked = true;
                i.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                i.dispatchEvent(new Event('input', { bubbles: true }));
                i.dispatchEvent(new Event('change', { bubbles: true }));
            }).catch(() => {});
            await bizagiPage.waitForTimeout(300);
            const checkedFinal = await radioDescartarDirecto.isChecked().catch(() => false);
            console.log(`[DEBUG-BIZAGI][OFAC][FallbackDOM] checkedFinal=${checkedFinal}`);
            accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
        }
    }

    const descartarTextos = bizagiPage.getByText(/^Descartar$/i);
    const totalDescartar = await descartarTextos.count().catch(() => 0);
    let radioCandidato: Locator | null = null;
    let radioEncontrado = false;
    for (let i = 0; i < totalDescartar; i++) {
        const texto = descartarTextos.nth(i);
        if (!(await texto.isVisible().catch(() => false))) continue;
        const tBox = await texto.boundingBox().catch(() => null);
        if (!tBox) continue;
        console.log(`[DEBUG-BIZAGI][OFAC] Descartar texto bbox=x:${Math.round(tBox.x)} y:${Math.round(tBox.y)} w:${Math.round(tBox.width)} h:${Math.round(tBox.height)}`);

        const radios = bizagiPage.locator('input[type="radio"]');
        const nRadios = await radios.count().catch(() => 0);
        let bestIdx = -1;
        let bestScore = Number.POSITIVE_INFINITY;
        for (let j = 0; j < nRadios; j++) {
            const r = radios.nth(j);
            if (!(await r.isVisible().catch(() => false))) continue;
            const rBox = await r.boundingBox().catch(() => null);
            if (!rBox) continue;
            const dy = Math.abs((rBox.y + rBox.height / 2) - (tBox.y + tBox.height / 2));
            const dx = (tBox.x - (rBox.x + rBox.width / 2));
            if (dy <= 20 && dx >= -8 && dx <= 80) {
                const score = dy * 10 + Math.abs(dx);
                if (score < bestScore) {
                    bestScore = score;
                    bestIdx = j;
                }
            }
        }

        if (bestIdx >= 0) {
            radioCandidato = radios.nth(bestIdx);
            radioEncontrado = true;
            break;
        }
    }
    console.log(`[DEBUG-BIZAGI][OFAC] radio cercano Descartar encontrado=${radioEncontrado}`);

    if (radioCandidato) {
        console.log('[DEBUG-BIZAGI][OFAC] click radio Descartar');
        await radioCandidato.click({ force: true }).catch(() => {});
        await bizagiPage.waitForTimeout(250);
        const checked = await radioCandidato.isChecked().catch(() => false);
        console.log(`[DEBUG-BIZAGI][OFAC] radio Descartar checked=${checked}`);
        accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
    }

    if (!accionDescartar && await descartarLabel.isVisible().catch(() => false)) {
        console.log('[DEBUG-BIZAGI][OFAC] click texto Descartar');
        await descartarLabel.click({ force: true }).catch(() => {});
        await bizagiPage.waitForTimeout(250);
        accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
    }

    if (!accionDescartar) {
        const contDesc = descartarLabel.locator('xpath=ancestor::*[self::label or self::div or self::span][1]').first();
        if (await contDesc.isVisible().catch(() => false)) {
            console.log('[DEBUG-BIZAGI][OFAC] click label/contenedor Descartar');
            await contDesc.click({ force: true }).catch(() => {});
            await bizagiPage.waitForTimeout(250);
            accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
        }
    }

    if (!accionDescartar && radioCandidato) {
        const rb = await radioCandidato.boundingBox().catch(() => null);
        if (rb) {
            console.log('[DEBUG-BIZAGI][OFAC] click coordenadas radio Descartar');
            await bizagiPage.mouse.click(rb.x + rb.width / 2, rb.y + rb.height / 2).catch(() => {});
            await bizagiPage.waitForTimeout(250);
            accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
        }
    }

    if (!accionDescartar && radioCandidato) {
        console.log('[DEBUG-BIZAGI][OFAC][FallbackDOM] marcando radio Descartar por JS');
        await radioCandidato.evaluate((el) => {
            const i = el as HTMLInputElement;
            i.checked = true;
            i.dispatchEvent(new Event('click', { bubbles: true }));
            i.dispatchEvent(new Event('change', { bubbles: true }));
            i.dispatchEvent(new Event('input', { bubbles: true }));
        }).catch(() => {});
        await bizagiPage.waitForTimeout(250);
        const checkedFinal = await radioCandidato.isChecked().catch(() => false);
        console.log(`[DEBUG-BIZAGI][OFAC][FallbackDOM] checkedFinal=${checkedFinal}`);
        accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
    }

    if (!accionDescartar) {
        const radioDescartar = contenedor.locator('input[type="radio"][value*="descartar" i], input[type="radio"][value*="discard" i]').first();
        if (await radioDescartar.isVisible().catch(() => false)) {
            await radioDescartar.click({ force: true }).catch(() => {});
            await bizagiPage.waitForTimeout(250);
            accionDescartar = await validarAccionOfacDescartarDebug(bizagiPage);
        }
    }

    const accionFinal = accionDescartar ? 'Descartar' : '';
    console.log(`[DEBUG-BIZAGI][OFAC] accionFinal='${accionFinal || ''}'`);
    console.log(`[DEBUG-BIZAGI][OFAC] accionDescartar=${accionDescartar}`);

    if (!accionDescartar) {
        throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] No se seleccionó Acción Coincidencias OFAC=Descartar');
    }

    return true;
}

async function seleccionarMotivoOfacDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][OFAC] Buscando Motivo Coincidencias OFAC');

    const labelMotivo = bizagiPage.getByText(/Motivo\s+Coincidencias\s+OFAC/i).first();
    const labelVisible = await labelMotivo.isVisible({ timeout: 2500 }).catch(() => false);
    if (!labelVisible) {
        throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] No se seleccionó Motivo Coincidencias OFAC');
    }

    let comboMotivo = labelMotivo.locator('xpath=following::input[@role="combobox" or contains(@class,"ui-selectmenu-value")][1]').first();
    let comboVisible = await comboMotivo.isVisible().catch(() => false);

    if (!comboVisible) {
        const contenedor = labelMotivo.locator('xpath=ancestor::*[self::div or self::tr or self::td or self::section][1]').first();
        comboMotivo = contenedor.locator('input[role="combobox"]:visible, input.ui-select-data.ui-selectmenu-value:visible, .ui-selectmenu input:visible').first();
        comboVisible = await comboMotivo.isVisible().catch(() => false);
    }

    console.log(`[DEBUG-BIZAGI][OFAC] comboMotivo localizado=${comboVisible}`);
    if (!comboVisible) {
        throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] No se seleccionó Motivo Coincidencias OFAC');
    }

    await comboMotivo.scrollIntoViewIfNeeded().catch(() => {});
    await comboMotivo.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(300);

    const listboxId = (await comboMotivo.getAttribute('aria-controls').catch(() => '')) || '';
    let opciones = listboxId
        ? bizagiPage.locator(`#${listboxId} li:visible, #${listboxId} [role="option"]:visible`)
        : bizagiPage.locator('li[role="option"]:visible, [role="option"]:visible, .ui-selectmenu-menu:visible li:visible');

    let opcionesCount = await opciones.count().catch(() => 0);
    if (opcionesCount === 0) {
        opciones = bizagiPage.locator('li[role="option"]:visible, [role="option"]:visible, .ui-selectmenu-menu:visible li:visible');
        opcionesCount = await opciones.count().catch(() => 0);
    }
    console.log(`[DEBUG-BIZAGI][OFAC] opcionesMotivo visibles=${opcionesCount}`);

    let opcion = opciones
        .filter({ hasText: /Coincidencia\s+descartada\s+por\s+no\s+corresponderse.*Listas\s+de\s+Control/i })
        .first();
    if (!(await opcion.isVisible().catch(() => false))) {
        opcion = opciones.filter({ hasText: /Coincidencia\s+descartada|no\s+corresponderse|Listas\s+de\s+Control/i }).first();
    }

    if (!(await opcion.isVisible().catch(() => false))) {
        throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] No se seleccionó Motivo Coincidencias OFAC');
    }

    const opcionText = ((await opcion.innerText().catch(() => '')) || '').replace(/\s+/g, ' ').trim();
    console.log(`[DEBUG-BIZAGI][OFAC] seleccionando motivo='${opcionText}'`);
    await opcion.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(500);
    await comboMotivo.press('Tab').catch(() => {});
    await bizagiPage.waitForTimeout(250);

    const motivoFinal = (
        (await comboMotivo.inputValue().catch(() => '')) ||
        (await comboMotivo.getAttribute('value').catch(() => '')) ||
        (await comboMotivo.getAttribute('title').catch(() => '')) ||
        (await comboMotivo.innerText().catch(() => ''))
    ).replace(/\s+/g, ' ').trim();

    const motivoSeleccionado = /Coincidencia\s+descartada|no\s+corresponderse|Listas\s+de\s+Control/i.test(motivoFinal);
    console.log(`[DEBUG-BIZAGI][OFAC] motivoFinal='${motivoFinal}'`);
    console.log(`[DEBUG-BIZAGI][OFAC] motivoSeleccionado=${motivoSeleccionado}`);

    if (!motivoSeleccionado) {
        throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] No se seleccionó Motivo Coincidencias OFAC');
    }

    return true;
}

export async function aprobarOfacGestionCoincidenciasDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][OFAC] Iniciando aprobación OFAC');

    const body = await bizagiPage.locator('body').innerText({ timeout: 3000 }).catch(() => '');
    const hayOfac = /Coincidencias\s+OFAC|Acci[oó]n\s+Coincidencias\s+OFAC|Motivo\s+Coincidencias\s+OFAC/i.test(body);

    if (!hayOfac) {
        console.log('[DEBUG-BIZAGI][OFAC] No hay sección OFAC visible');
        return false;
    }

    const accionDescartar = await seleccionarAccionOfacDescartarDebug(bizagiPage);
    const motivoSeleccionado = await seleccionarMotivoOfacDebug(bizagiPage);

    const ok = accionDescartar && motivoSeleccionado;
    console.log(`[DEBUG-BIZAGI][OFAC] validacionFinal accionDescartar=${accionDescartar} motivoSeleccionado=${motivoSeleccionado}`);

    if (!ok) {
        throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] OFAC no quedó completado');
    }

    console.log('[DEBUG-BIZAGI][OFAC] OFAC completado');
    return true;
}

async function localizarSeccionOtrasCoincidenciasDebug(bizagiPage: Page): Promise<Locator> {
    const header = bizagiPage.getByText(/^Otras Coincidencias$/i).first();
    const headerVisible = await header.isVisible({ timeout: 2500 }).catch(() => false);

    let seccion = bizagiPage.locator('body').first();
    if (headerVisible) {
        const candidata = header.locator(
            'xpath=ancestor::*[self::div or self::section or self::table][.//*[contains(normalize-space(.),"Lexis Nexis")] and .//*[contains(normalize-space(.),"Acción")]][1]'
        ).first();
        if (await candidata.isVisible().catch(() => false)) {
            seccion = candidata;
        } else {
            const candidata2 = header.locator(
                'xpath=ancestor::*[self::div or self::section or self::table][1]'
            ).first();
            if (await candidata2.isVisible().catch(() => false)) {
                seccion = candidata2;
            }
        }
    } else {
        console.log('[DEBUG-BIZAGI][LEXIS][WARN] Header Otras Coincidencias no visible; usando fallback body');
    }

    const seccionVisible = await seccion.isVisible().catch(() => false);
    console.log(`[DEBUG-BIZAGI][LEXIS] sección Otras Coincidencias localizada=${seccionVisible}`);
    return seccion;
}

type ComboLexisMeta = {
    idx: number;
    globalIdx: number;
    id: string;
    selector: string;
    value: string;
    title: string;
    dataValue: string;
    aria: string;
    ariaControls: string;
    textAround: string;
    x: number;
    y: number;
    width: number;
    height: number;
    visible: boolean;
    esLexisPorFila: boolean;
    esAccionPorColumna: boolean;
};

async function obtenerCombosAccionLexisMetaDebug(bizagiPage: Page): Promise<{ combos: ComboLexisMeta[]; textosLexis: number; }> {
    return await bizagiPage.evaluate(() => {
        const norm = (s?: string | null) => (s || '').replace(/\s+/g, ' ').trim();

        const headers = Array.from(document.querySelectorAll('.bz-rn-grid-header-title')) as HTMLElement[];
        const header = headers.find(h => /^Otras Coincidencias$/i.test(norm(h.innerText)));
        const seccion =
            header?.closest('.ui-bizagi-grid') ||
            header?.closest('.ui-bizagi-render') ||
            header?.closest('.bzg-form-grid') ||
            document.body;

        const allLexis = Array.from(seccion.querySelectorAll('label, span, div')) as HTMLElement[];
        const textosLexis = allLexis.filter(el => /^Lexis Nexis$/i.test(norm(el.innerText))).length;

        const lexisRefs = allLexis
            .filter(el => /^Lexis Nexis$/i.test(norm(el.innerText)))
            .map(el => {
                const r = el.getBoundingClientRect();
                return { y: r.top + r.height / 2, x: r.left };
            });

        const headersAccion = allLexis.find(el => /%\s*Coinc\.?/i.test(norm(el.innerText)));
        const xPct = headersAccion ? headersAccion.getBoundingClientRect().left : 0;

        const rawInputs = Array.from(
            seccion.querySelectorAll('input[role="combobox"], input.ui-select-data.ui-selectmenu-value, .ui-selectmenu input')
        ) as HTMLInputElement[];

        const combos = rawInputs.map((input, idx) => {
            const rect = input.getBoundingClientRect();
            const visible = rect.width > 0 && rect.height > 0;
            const centerY = rect.top + rect.height / 2;
            const centerX = rect.left + rect.width / 2;
            const value = norm(input.value || input.getAttribute('value'));
            const title = norm(input.getAttribute('title'));
            const dataValue = norm(input.getAttribute('data-value'));
            const ariaControls = norm(input.getAttribute('aria-controls'));
            const aria = norm(input.getAttribute('aria-activedescendant') || input.getAttribute('aria-controls'));

            const parent = input.closest('tr, .ui-bizagi-grid-row, [role="row"], .ui-bizagi-grid-cell, .ui-bizagi-render-control') as HTMLElement | null;
            const textAround = norm(parent?.innerText || input.closest('div')?.innerText || '');

            const esLexisPorFila =
                /Lexis Nexis/i.test(textAround) ||
                lexisRefs.some(r => Math.abs(centerY - r.y) <= 25 && centerX > r.x);

            const esMotivoOfac =
                /Motivo\s+Coincidencias\s+OFAC/i.test(textAround) ||
                /Coincidencia descartada por no corresponderse/i.test([value, title, dataValue, textAround].join(' | '));

            const esAccionPorColumna = xPct ? rect.left > xPct : true;

            return {
                idx,
                globalIdx: idx,
                id: input.id || '',
                selector: input.id ? `#${input.id}` : '',
                value,
                title,
                dataValue,
                aria,
                ariaControls,
                textAround,
                x: rect.left,
                y: rect.top,
                width: rect.width,
                height: rect.height,
                visible,
                esLexisPorFila,
                esAccionPorColumna,
                esMotivoOfac,
            };
        }).filter(c => c.visible && c.esLexisPorFila && c.esAccionPorColumna && !c.esMotivoOfac);

        return { combos, textosLexis };
    });
}

async function localizarComboLexisPorCandidato(bizagiPage: Page, candidato: ComboLexisMeta): Promise<Locator> {
    if (candidato.id) {
        return bizagiPage.locator(`#${cssEscapeDebug(candidato.id)}`).first();
    }

    if (candidato.ariaControls) {
        return bizagiPage.locator(`input[aria-controls="${cssEscapeDebug(candidato.ariaControls)}"]`).first();
    }

    const inputs = bizagiPage.locator('input[role="combobox"]:visible, input.ui-select-data.ui-selectmenu-value:visible, .ui-selectmenu input:visible');
    const count = await inputs.count().catch(() => 0);
    let bestIdx = 0;
    let bestScore = Number.POSITIVE_INFINITY;
    for (let i = 0; i < count; i++) {
        const box = await inputs.nth(i).boundingBox().catch(() => null);
        if (!box) continue;
        const dx = Math.abs(box.x - candidato.x);
        const dy = Math.abs(box.y - candidato.y);
        const dw = Math.abs(box.width - candidato.width);
        const dh = Math.abs(box.height - candidato.height);
        const score = dy * 20 + dx * 5 + dw + dh;
        if (score < bestScore) {
            bestScore = score;
            bestIdx = i;
        }
    }
    return inputs.nth(bestIdx);
}

function estadoCombinadoCombo(meta: ComboLexisMeta): string {
    return normalizarTextoCombo([meta.value, meta.title, meta.dataValue, meta.aria].filter(Boolean).join(' | '));
}

async function seleccionarFalsoPositivoEnComboLexisDebug(
    bizagiPage: Page,
    combo: Locator,
    comboIndex: number,
    estadoInicialParam?: string
): Promise<boolean> {
    const label = `[DEBUG-BIZAGI][LEXIS] combo ${comboIndex + 1}`;
    const estadoInicial = estadoInicialParam ?? await combo.inputValue().catch(async () => {
        return await combo.getAttribute('value').catch(() => '') || '';
    });

    if (esAccionLexisFalsoPositivo(estadoInicial)) {
        console.log(`${label} ya tiene Falso Positivo; se omite`);
        return true;
    }

    const pendiente = esAccionLexisPendiente(estadoInicial);
    const esFp = esAccionLexisFalsoPositivo(estadoInicial);
    console.log(`${label} estadoInicial='${estadoInicial}'`);
    console.log(`${label} esFalsoPositivo=${esFp} esPendiente=${pendiente}`);
    console.log(`${label} pendiente=${pendiente}; seleccionando Falso Positivo`);
    if (!pendiente) return true;

    await combo.scrollIntoViewIfNeeded().catch(() => {});
    await combo.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(250);

    const ariaControls = (await combo.getAttribute('aria-controls').catch(() => '')) || '';
    console.log(`${label} aria-controls='${ariaControls}'`);

    let opciones: Locator;
    let totalOpciones = 0;
    if (ariaControls) {
        opciones = bizagiPage.locator(`#${ariaControls} li:visible, #${ariaControls} [role="option"]:visible`);
        totalOpciones = await opciones.count().catch(() => 0);
        if (totalOpciones > 0) {
            console.log(`${label} usando opciones por aria-controls`);
        }
    } else {
        opciones = bizagiPage.locator('li[role="option"]:visible, [role="option"]:visible');
    }

    if (!ariaControls || totalOpciones === 0) {
        console.log(`${label}[WARN] usando fallback global de opciones`);
        opciones = bizagiPage.locator('li[role="option"]:visible, [role="option"]:visible, .ui-selectmenu-menu:visible li:visible');
        totalOpciones = await opciones.count().catch(() => 0);
    }

    console.log(`${label} opciones visibles=${totalOpciones}`);

    let opcionFalso = opciones.filter({ hasText: /^\s*Falso\s*Positivo\s*$/i }).first();
    if (!(await opcionFalso.isVisible().catch(() => false))) {
        opcionFalso = opciones.filter({ hasText: /Falso\s*Positivo/i }).first();
    }

    if (await opcionFalso.isVisible().catch(() => false)) {
        console.log(`${label} Opción 'Falso Positivo' encontrada`);
        await opcionFalso.click({ force: true }).catch(() => {});
    } else {
        // fallback por teclado
        await combo.focus().catch(() => {});
        await bizagiPage.keyboard.press('ArrowDown').catch(() => {});
        await bizagiPage.waitForTimeout(200);
        await bizagiPage.keyboard.press('Enter').catch(() => {});
    }

    await bizagiPage.waitForTimeout(500);
    await combo.press('Tab').catch(() => {});
    await bizagiPage.waitForTimeout(300);

    const estadoFinal = await combo.inputValue().catch(async () => {
        return await combo.getAttribute('value').catch(() => '') || '';
    });
    console.log(`${label} estadoFinal='${estadoFinal}'`);
    return esAccionLexisFalsoPositivo(estadoFinal);
}

export async function completarLexisNexisOtrasCoincidenciasDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][LEXIS] Iniciando procesamiento multi-fila Lexis Nexis');

    try {
        const seccion = await localizarSeccionOtrasCoincidenciasDebug(bizagiPage);
        const seccionVisible = await seccion.isVisible().catch(() => false);
        if (!seccionVisible) {
            console.log('[DEBUG-BIZAGI][LEXIS] No hay sección Lexis Nexis visible');
            return false;
        }

        const inicial = await obtenerCombosAccionLexisMetaDebug(bizagiPage);
        const combosIniciales = inicial.combos;
        console.log(`[DEBUG-BIZAGI][LEXIS] combosAccionLexis detectados=${combosIniciales.length}`);
        combosIniciales.forEach((c, i) => {
            console.log(`[DEBUG-BIZAGI][LEXIS] combo[${i}] id='${c.id}' selector='${c.id ? `#${cssEscapeDebug(c.id)}` : ''}' estado='${estadoCombinadoCombo(c)}' y=${c.y}`);
            if (/Coincidencia descartada por no corresponderse|Motivo\s+Coincidencias\s+OFAC/i.test(c.textAround || '')) {
                console.log(`[DEBUG-BIZAGI][LEXIS] combo descartado por pertenecer a Motivo OFAC id='${c.id}'`);
            }
        });

        if (inicial.textosLexis > combosIniciales.length) {
            console.log(`[DEBUG-BIZAGI][LEXIS][WARN] textosLexis=${inicial.textosLexis} pero combosAccionLexis=${combosIniciales.length}`);
        }

        if (combosIniciales.length === 0) {
            console.log('[DEBUG-BIZAGI][LEXIS] Sección Lexis presente pero sin combos procesables');
            return false;
        }

        const combosPendientesInicial = combosIniciales.filter(c => esAccionLexisPendiente(estadoCombinadoCombo(c))).length;
        let combosActualizados = 0;

        for (let i = 0; i < combosIniciales.length; i++) {
            const meta = combosIniciales[i];
            console.log(`[DEBUG-BIZAGI][LEXIS] procesando combo id='${meta.id}'`);

            const combo = await localizarComboLexisPorCandidato(bizagiPage, meta);
            const idReal = (await combo.getAttribute('id').catch(() => '')) || '';
            const ariaReal = (await combo.getAttribute('aria-controls').catch(() => '')) || '';
            console.log(`[DEBUG-BIZAGI][LEXIS] combo ${i + 1} candidatoId='${meta.id}' idReal='${idReal}' ariaReal='${ariaReal}'`);
            if (meta.id && idReal && idReal !== meta.id) {
                throw new Error('[DEBUG-BIZAGI][LEXIS][CRITICO] Locator reconstruido no coincide con candidato Lexis');
            }

            const estadoInicial = normalizarTextoCombo([
                (await combo.inputValue().catch(() => '')) || '',
                (await combo.getAttribute('value').catch(() => '')) || '',
                (await combo.getAttribute('title').catch(() => '')) || '',
                ariaReal,
            ].filter(Boolean).join(' | ')) || estadoCombinadoCombo(meta);
            const esFpInicial = esAccionLexisFalsoPositivo(estadoInicial);
            const pendienteInicial = esAccionLexisPendiente(estadoInicial);
            console.log(`[DEBUG-BIZAGI][LEXIS] combo ${i + 1} estadoInicial='${estadoInicial}'`);
            console.log(`[DEBUG-BIZAGI][LEXIS] combo ${i + 1} esFalsoPositivo=${esFpInicial} esPendiente=${pendienteInicial}`);

            const ok = await seleccionarFalsoPositivoEnComboLexisDebug(bizagiPage, combo, i, estadoInicial).catch((e) => {
                console.log(`[DEBUG-BIZAGI][LEXIS][ERROR] combo ${i + 1} error=${String(e)}`);
                return false;
            });
            if (ok && pendienteInicial) combosActualizados++;
        }

        const final = await obtenerCombosAccionLexisMetaDebug(bizagiPage);
        const combosPendientesFinal = final.combos.filter(c => esAccionLexisPendiente(estadoCombinadoCombo(c))).length;
        console.log(`[DEBUG-BIZAGI][LEXIS] combosTotal=${final.combos.length}`);
        console.log(`[DEBUG-BIZAGI][LEXIS] combosPendientesInicial=${combosPendientesInicial}`);
        console.log(`[DEBUG-BIZAGI][LEXIS] combosActualizados=${combosActualizados}`);
        console.log(`[DEBUG-BIZAGI][LEXIS] combosPendientesFinal=${combosPendientesFinal}`);
        console.log(`[DEBUG-BIZAGI][LEXIS] filasPendientesFinal=${combosPendientesFinal}`);

        if (final.textosLexis > final.combos.length) {
            console.log(`[DEBUG-BIZAGI][LEXIS][WARN] textosLexis=${final.textosLexis} pero combosAccionLexis=${final.combos.length}`);
        }

        if (combosPendientesFinal > 0) {
            throw new Error('[DEBUG-BIZAGI][LEXIS][CRITICO] Quedan combos Lexis Nexis pendientes');
        }

        console.log('[DEBUG-BIZAGI][LEXIS] Todas las filas Lexis Nexis fueron completadas');
        return true;
    } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        console.log(`[DEBUG-BIZAGI][LEXIS][ERROR] ${msg}`);
        throw e;
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

    if (modo === 'MIXTO' || modo === 'OFAC') {
        resultOfac = await aprobarOfacGestionCoincidenciasDebug(bizagiPage);
        const bodyOfac = await bizagiPage.locator('body').innerText({ timeout: 2000 }).catch(() => '');
        const hayOfac = /Coincidencias\s+OFAC|Acci[oó]n\s+Coincidencias\s+OFAC|Motivo\s+Coincidencias\s+OFAC/i.test(bodyOfac);
        if (hayOfac && !resultOfac) {
            throw new Error('[DEBUG-BIZAGI][OFAC][CRITICO] OFAC detectado pero no fue completado');
        }
        if (!resultOfac && modo === 'OFAC') {
            throw new Error('[DEBUG-BIZAGI][CRITICO] Modo OFAC falló');
        }
    }

    if (modo === 'MIXTO' || modo === 'LEXIS') {
        const bodyLexis = await bizagiPage.locator('body').innerText({ timeout: 2000 }).catch(() => '');
        const hayLexis = /Lexis\s+Nexis|Otras\s+Coincidencias/i.test(bodyLexis);
        try {
            resultLexis = await completarLexisNexisOtrasCoincidenciasDebug(bizagiPage);
        } catch (e) {
            const msg = e instanceof Error ? e.message : String(e);
            console.log(`[DEBUG-BIZAGI][Modo] Error en LEXIS: ${msg}`);
            if (hayLexis || modo === 'LEXIS') {
                throw new Error(`[DEBUG-BIZAGI][LEXIS][CRITICO] Lexis Nexis detectado pero no fue completado: ${msg}`);
            }
            resultLexis = false;
        }
        if (hayLexis && !resultLexis) {
            throw new Error('[DEBUG-BIZAGI][LEXIS][CRITICO] Lexis Nexis detectado pero resultó false');
        }
    }

    if (modo === 'PLAFT') {
        resultPlaft = await completarPlaftGestionCoincidenciasDebug(bizagiPage);
        if (!resultPlaft) {
            throw new Error('[DEBUG-BIZAGI][CRITICO] Modo PLAFT falló');
        }
    }

    const bodyFinal = await bizagiPage.locator('body').innerText({ timeout: 2000 }).catch(() => '');
    const hayOfac = /Coincidencias\s+OFAC|Acci[oó]n\s+Coincidencias\s+OFAC|Motivo\s+Coincidencias\s+OFAC/i.test(bodyFinal);
    const hayLexis = /Lexis\s+Nexis|Otras\s+Coincidencias/i.test(bodyFinal);
    console.log(`[DEBUG-BIZAGI][Modo] validacionFinal hayOfac=${hayOfac} resultOfac=${resultOfac} hayLexis=${hayLexis} resultLexis=${resultLexis}`);
    if (hayOfac && !resultOfac) {
        throw new Error('[DEBUG-BIZAGI][Modo][CRITICO] MIXTO no puede completar con OFAC=false');
    }
    if (hayLexis && !resultLexis) {
        throw new Error('[DEBUG-BIZAGI][Modo][CRITICO] MIXTO no puede completar con LEXIS=false');
    }

    console.log(`[DEBUG-BIZAGI][Modo] Modos completados: OFAC=${resultOfac}, LEXIS=${resultLexis}, PLAFT=${resultPlaft}`);
}

export async function resolverModoBizagiDebug(bizagiPage: Page): Promise<'OFAC' | 'PLAFT' | 'LEXIS' | 'MIXTO'> {
    console.log('[DEBUG-BIZAGI][Auto] Resolviendo modo automático...');

    const pageText = await bizagiPage.locator('body').innerText({ timeout: 3000 }).catch(() => '');

    const tieneGestionCoincidencias = /Gesti[oó]n de Coincidencias|Gestion de Coincidencias/i.test(pageText);
    const tieneVerificarPlaft = /Verificar PLAFT|Debida Diligencia PLAFT/i.test(pageText);
    const tieneLexisNexis = /Lexis\s*Nexis/i.test(pageText);
    const tieneOfac = /OFAC|Listas OFAC|Coincidencias OFAC/i.test(pageText);

    console.log(`[DEBUG-BIZAGI][Auto] Detectados: GestionCoincidencias=${tieneGestionCoincidencias}, VerificarPlaft=${tieneVerificarPlaft}, LexisNexis=${tieneLexisNexis}, OFAC=${tieneOfac}`);

    let modoResuelto: 'OFAC' | 'PLAFT' | 'LEXIS' | 'MIXTO' = 'LEXIS';

    if (tieneGestionCoincidencias && tieneVerificarPlaft) {
        modoResuelto = 'MIXTO';
        console.log('[DEBUG-BIZAGI][Auto] Ambas actividades detectadas => MIXTO');
    } else if (tieneGestionCoincidencias) {
        if (tieneLexisNexis) {
            modoResuelto = 'LEXIS';
            console.log('[DEBUG-BIZAGI][Auto] Solo Gestionar Coincidencias con Lexis => LEXIS');
        } else if (tieneOfac) {
            modoResuelto = 'OFAC';
            console.log('[DEBUG-BIZAGI][Auto] Solo Gestionar Coincidencias con OFAC => OFAC');
        } else {
            modoResuelto = 'MIXTO';
            console.log('[DEBUG-BIZAGI][Auto] Solo Gestionar Coincidencias (ambigua) => MIXTO');
        }
    } else if (tieneVerificarPlaft) {
        modoResuelto = 'PLAFT';
        console.log('[DEBUG-BIZAGI][Auto] Solo Verificar PLAFT => PLAFT');
    }

    console.log(`[DEBUG-BIZAGI][Auto] Modo resuelto: ${modoResuelto}`);
    return modoResuelto;
}

export async function seleccionarSolicitarAclaracionesNoDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][Aclaraciones] Buscando campo Solicitar Aclaraciones');

    const label = bizagiPage.getByText(/Solicitar\s+Aclaraciones\??/i).first();
    const labelVisible = await label.isVisible({ timeout: 2500 }).catch(() => false);
    console.log(`[DEBUG-BIZAGI][Aclaraciones] label visible=${labelVisible}`);

    if (!labelVisible) {
        return false;
    }

    console.log('[DEBUG-BIZAGI][Aclaraciones] intentando seleccionar No');

    let seleccionado = false;

    // 1) Radio/input/label cercano
    const contenedor = label.locator('xpath=ancestor::*[self::div or self::span or self::section][1]').first();
    const noLabel = contenedor.getByText(/^No$/i).first();
    if (await noLabel.isVisible().catch(() => false)) {
        await noLabel.click({ force: true }).catch(() => {});
        seleccionado = true;
    }

    // 2) Buscar input radio cercano y marcar
    if (!seleccionado) {
        const radioNo = bizagiPage.locator('input[type="radio"][value="false"], input[type="radio"][value="No"], input[type="radio"][value="0"]').first();
        if (await radioNo.isVisible().catch(() => false)) {
            await radioNo.click({ force: true }).catch(async () => {
                await radioNo.evaluate((el) => {
                    const i = el as HTMLInputElement;
                    i.checked = true;
                    i.dispatchEvent(new Event('input', { bubbles: true }));
                    i.dispatchEvent(new Event('change', { bubbles: true }));
                }).catch(() => {});
            });
            seleccionado = true;
        }
    }

    // 3) Fallback geométrico
    if (!seleccionado) {
        const labelBox = await label.boundingBox().catch(() => null);
        if (labelBox) {
            const noCandidates = bizagiPage.locator('label:visible, span:visible, div:visible').filter({ hasText: /^No$/i });
            const totalNo = await noCandidates.count().catch(() => 0);
            let bestIdx = -1;
            let bestScore = Number.POSITIVE_INFINITY;
            for (let i = 0; i < totalNo; i++) {
                const c = noCandidates.nth(i);
                const box = await c.boundingBox().catch(() => null);
                if (!box) continue;
                const dy = Math.abs((box.y + box.height / 2) - (labelBox.y + labelBox.height / 2));
                const dx = Math.max(0, labelBox.x - box.x);
                const score = dy * 10 + dx;
                if (box.x >= labelBox.x && score < bestScore) {
                    bestScore = score;
                    bestIdx = i;
                }
            }
            if (bestIdx >= 0) {
                await noCandidates.nth(bestIdx).click({ force: true }).catch(() => {});
                seleccionado = true;
            }
        }
    }

    await bizagiPage.waitForTimeout(400);

    const validacionFinalNo = await bizagiPage.evaluate(() => {
        const body = (document.body?.innerText || '').replace(/\s+/g, ' ').trim();
        const radios = Array.from(document.querySelectorAll('input[type="radio"]')) as HTMLInputElement[];
        const algunoNoChecked = radios.some(r => r.checked && (/false|no|0/i.test(String(r.value || ''))));
        const ariaTrueNo = Array.from(document.querySelectorAll('[aria-checked="true"]')).some((el) => /\bNo\b/i.test((el as HTMLElement).innerText || ''));
        return algunoNoChecked || ariaTrueNo || /Solicitar\s+Aclaraciones/i.test(body);
    }).catch(() => false);

    console.log(`[DEBUG-BIZAGI][Aclaraciones] No seleccionado=${seleccionado}`);
    console.log(`[DEBUG-BIZAGI][Aclaraciones] validacionFinalNo=${validacionFinalNo}`);

    if (!seleccionado || !validacionFinalNo) {
        throw new Error('[DEBUG-BIZAGI][Aclaraciones][CRITICO] No se pudo seleccionar Solicitar Aclaraciones=No');
    }

    return true;
}

async function aceptarModalConfirmacionBizagiDebug(bizagiPage: Page): Promise<boolean> {
    console.log('[DEBUG-BIZAGI][Confirmacion] Buscando modal de confirmación Bizagi');

    const modal = bizagiPage
        .locator('[role="dialog"]:visible, .modal:visible, .ui-dialog:visible, div:has-text("Bizagi"), div:has-text("¿Está seguro de continuar")')
        .filter({ hasText: /Bizagi|Está seguro de continuar|no es posible deshacer/i })
        .first();

    const modalVisible = await modal.isVisible({ timeout: 8000 }).catch(() => false);
    console.log(`[DEBUG-BIZAGI][Confirmacion] modalVisible=${modalVisible}`);

    if (!modalVisible) {
        return false;
    }

    const btnAceptar = await pickVisible(bizagiPage, [
        modal.getByRole('button', { name: /Aceptar/i }).first(),
        modal.locator('button:has-text("Aceptar")').first(),
        modal.locator('input[type="button"][value="Aceptar"]').first(),
        bizagiPage.getByRole('button', { name: /Aceptar/i }).first(),
        bizagiPage.locator('button:has-text("Aceptar")').first(),
    ]);

    if (!(await btnAceptar.isVisible().catch(() => false))) {
        throw new Error('[DEBUG-BIZAGI][Confirmacion][CRITICO] Modal Bizagi visible pero no se pudo aceptar');
    }

    console.log('[DEBUG-BIZAGI][Confirmacion] Click Aceptar');
    await btnAceptar.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(600);

    const modalCerrado = !(await modal.isVisible().catch(() => false));
    console.log(`[DEBUG-BIZAGI][Confirmacion] modalCerrado=${modalCerrado}`);

    if (!modalCerrado) {
        throw new Error('[DEBUG-BIZAGI][Confirmacion][CRITICO] Modal Bizagi visible pero no se pudo aceptar');
    }

    await bizagiPage.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {});
    await bizagiPage.waitForTimeout(1500);
    const url = bizagiPage.url();
    const bodyPreview = ((await bizagiPage.locator('body').innerText().catch(() => '')) || '').replace(/\s+/g, ' ').trim().slice(0, 260);
    console.log(`[DEBUG-BIZAGI][Confirmacion] postAceptar url='${url}'`);
    console.log(`[DEBUG-BIZAGI][Confirmacion] postAceptar bodyPreview='${bodyPreview}'`);

    return true;
}

export async function clickSiguienteGestionCoincidenciasDebug(bizagiPage: Page, opts?: { requiereConfirmacion?: boolean }): Promise<void> {
    console.log('[DEBUG-BIZAGI][Siguiente] Buscando botón Siguiente');

    const candidatos = [
        bizagiPage.locator('input#formButton1[value="Siguiente"]:visible').first(),
        bizagiPage.locator('input[type="button"][value="Siguiente"]:visible').first(),
        bizagiPage.locator('button:has-text("Siguiente"):visible').first(),
        bizagiPage.getByRole('button', { name: /Siguiente/i }).first(),
        bizagiPage.getByText(/^Siguiente$/i).first(),
    ];

    const btn = await pickVisible(bizagiPage, candidatos);
    if (!(await btn.isVisible().catch(() => false))) {
        throw new Error('[DEBUG-BIZAGI][Siguiente][CRITICO] No se encontró botón Siguiente');
    }

    console.log('[DEBUG-BIZAGI][Siguiente] Click Siguiente');
    await btn.click({ force: true }).catch(() => {});
    await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
    await bizagiPage.waitForTimeout(1200);
    console.log('[DEBUG-BIZAGI][Siguiente] Click completado');

    const confirmacionOk = await aceptarModalConfirmacionBizagiDebug(bizagiPage);
    console.log(`[DEBUG-BIZAGI][Siguiente] confirmacionOk=${confirmacionOk}`);

    if (!confirmacionOk) {
        console.log('[DEBUG-BIZAGI][Siguiente][WARN] No apareció modal de confirmación después de Siguiente');
        if (opts?.requiereConfirmacion) {
            throw new Error('[DEBUG-BIZAGI][Siguiente][CRITICO] No se confirmó el modal Bizagi después de Siguiente');
        }
    }
}

export async function guardarYSiguienteDebug(bizagiPage: Page): Promise<void> {
    // Conservado por compatibilidad: ahora solo avanza con Siguiente (sin Guardar)
    await clickSiguienteGestionCoincidenciasDebug(bizagiPage);
}
