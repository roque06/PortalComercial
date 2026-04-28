import type { Locator, Page } from '@playwright/test';

const BIZAGI_URL_NUEVA = 'https://test-bscrd-santacruz.bizagi.com/';
const BIZAGI_URL_ANTIGUA_REGEX = /https?:\/\/(?:srvdevbizapp01\/BPM_BSC\/?|portalcomercial\.dev\.bsc\.com\.do\/login\/?|srvqabizapp01\/BPM_BSC_QA\/?|dev-bscrd-santacruz\.bizagi\.com\/?)/i;

function resolverBizagiUrl(url?: string) {
  const candidata = url ?? process.env.BIZAGI_URL ?? process.env.PW_BIZAGI_URL ?? BIZAGI_URL_NUEVA;
  return BIZAGI_URL_ANTIGUA_REGEX.test(candidata) ? BIZAGI_URL_NUEVA : candidata;
}

async function clickPrimeroVisible(page: Page, candidatos: Locator[], timeoutMs = 12000) {
  const fin = Date.now() + timeoutMs;
  while (Date.now() < fin) {
    for (const loc of candidatos) {
      const visible = await loc.isVisible().catch(() => false);
      if (!visible) continue;
      await loc.scrollIntoViewIfNeeded().catch(() => { });
      const ok = await loc.click({ force: true }).then(() => true).catch(() => false);
      if (ok) return true;
    }
    await page.waitForTimeout(250);
  }
  return false;
}

async function isBizagiHomeReady(page: Page) {
  const adminCandidates: Locator[] = [
    page.locator('#mnuBt\\.Admin\\.Reports').first(),
    page.locator('[id="mnuBt.Admin.Reports"]').first(),
    page.locator('li[id*="Admin"][id*="Reports"]').first(),
    page.getByRole('button', { name: /^Admin\b/i }).first(),
    page.getByRole('link', { name: /^Admin\b/i }).first(),
    page.locator('a,button,li,span').filter({ hasText: /^Admin\b/i }).first(),
    // Additional selectors for portalcomercial.dev.bsc.com.do Bizagi home
    page.locator('input#menuQuery').first(),
    page.locator('#main-layout').first(),
    page.locator('.pc-avatar, .user-name').first(),
    page.locator('[class*="inbox"], [class*="worklist"], #inbox, .bpm-inbox').first(),
    page.locator('a[href*="inbox"], a[href*="worklist"]').first(),
  ];
  for (const loc of adminCandidates) {
    if (await loc.isVisible().catch(() => false)) return true;
  }
  // URL-based fallback: if no longer on the login page and on a Bizagi host, assume home
  const url = page.url();
  if (!/\/login\b/i.test(url) && /portalcomercial|bizagi/i.test(url)) {
    const textoLogin = page.getByText(/Seleccione sus datos para ingresar a Bizagi/i).first();
    if (!(await textoLogin.isVisible().catch(() => false))) return true;
  }
  return false;
}

async function esperarHomeBizagi(page: Page, timeoutMs = 90000) {
  const fin = Date.now() + timeoutMs;
  while (Date.now() < fin) {
    if (await isBizagiHomeReady(page)) return true;
    const btnIngresar = page.getByRole('button', { name: /Ingresar/i }).first();
    const loginVisible = await btnIngresar.isVisible().catch(() => false);
    if (loginVisible) return false;
    await page.waitForTimeout(350);
  }
  return false;
}

async function abrirCasosEnBizagi(bizagiPage: Page) {
  const modalCasos = bizagiPage
    .locator('.ui-dialog:visible, [role="dialog"]:visible, .modal:visible')
    .filter({ hasText: /Casos/i })
    .first();

  if (await modalCasos.isVisible().catch(() => false)) return modalCasos;

  const homeReady = await esperarHomeBizagi(bizagiPage, 90000);
  if (!homeReady) {
    throw new Error(`No se pudo cargar home Bizagi luego del login. URL=${bizagiPage.url()}`);
  }

  const abrioAdmin = await clickPrimeroVisible(
    bizagiPage,
    [
      bizagiPage.locator('[id="mnuBt.Admin.Reports"]').first(),
      bizagiPage.locator('#mnuBt\\.Admin\\.Reports').first(),
      bizagiPage.locator('li[id*="Admin"][id*="Reports"]').first(),
      bizagiPage.getByRole('button', { name: /^Admin\b/i }).first(),
      bizagiPage.getByRole('link', { name: /^Admin\b/i }).first(),
      bizagiPage.locator('a,button,li,span').filter({ hasText: /^Admin\b/i }).first(),
    ],
    45000
  );
  if (!abrioAdmin) throw new Error('No se pudo abrir menu Admin en Bizagi.');

  const abrioAdminProcesos = await clickPrimeroVisible(
    bizagiPage,
    [
      bizagiPage.getByText(/Administraci(?:o|\u00f3)n de procesos/i).first(),
      bizagiPage.getByText(/Administracion de procesos/i).first(),
      bizagiPage.locator('li:has-text("Administracion de procesos")').first(),
      bizagiPage.locator('li:has(span.title:has-text("Administración de procesos"))').first(),
    ],
    40000
  );
  if (!abrioAdminProcesos) throw new Error('No se pudo abrir Administracion de procesos en Bizagi.');

  const abrioCasos = await clickPrimeroVisible(
    bizagiPage,
    [
      bizagiPage.getByText(/^Casos$/i).first(),
      bizagiPage.locator('li:has(span.title:has-text("Casos"))').first(),
      bizagiPage.locator('li:has-text("Casos")').first(),
    ],
    35000
  );
  if (!abrioCasos) throw new Error('No se pudo abrir la opcion Casos en Bizagi.');

  await modalCasos.waitFor({ state: 'visible', timeout: 30000 });
  return modalCasos;
}
async function findVisibleLocator(page: Page, selectors: string[], label: string = "Element") {
  // 1. Visible
  for (const sel of selectors) {
    const loc = page.locator(sel).first();
    if (await loc.isVisible().catch(() => false)) {
      console.log(`[BizagiLogin] ${label} visible en main con selector: ${sel}`);
      return loc;
    }
    for (const [idx, frame] of page.frames().entries()) {
      const fLoc = frame.locator(sel).first();
      if (await fLoc.isVisible().catch(() => false)) {
        console.log(`[BizagiLogin] ${label} visible en Frame #${idx} (${frame.name() || 'sin nombre'}) con selector: ${sel}`);
        return fLoc;
      }
    }
  }

  // 2. Existencia
  for (const sel of selectors) {
    const loc = page.locator(sel).first();
    if (await loc.count().catch(() => 0) > 0) {
      console.log(`[BizagiLogin] ${label} existe en main con selector: ${sel}`);
      return loc;
    }
    for (const [idx, frame] of page.frames().entries()) {
      const fLoc = frame.locator(sel).first();
      if (await fLoc.count().catch(() => 0) > 0) {
        console.log(`[BizagiLogin] ${label} existe en Frame #${idx} (${frame.name() || 'sin nombre'}) con selector: ${sel}`);
        return fLoc;
      }
    }
  }

  // 3. Ultimate Fallback
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

async function entrarABizagiSiHayLogin(bizagiPage: Page, usuario: string, password?: string) {
  const bizagiPass = password || process.env.BIZAGI_PASS || "H0la1234.";
  const fin = Date.now() + 90000;
  while (Date.now() < fin) {
    if (await isBizagiHomeReady(bizagiPage)) return;

    const userSels = ['#user', 'input#user', 'input[name="username"]', 'input#username', 'input[autocomplete="username"]', 'xpath=(//label[contains(normalize-space(.),"Usuario")]/following::input[1])[1]'];
    const passSels = ['#password', 'input#password', 'input[type="password"]', 'input[name="password"]', 'xpath=(//label[contains(normalize-space(.),"Contrase")]/following::input[1])[1]'];
    const btnSels = ['#btn-login', 'button#btn-login', 'button#btn-adminlogin', 'button:has-text("Ingresar")', 'button[name*="login" i]'];

    const inputUser = await findVisibleLocator(bizagiPage, userSels, "Usuario");
    const inputPass = await findVisibleLocator(bizagiPage, passSels, "Password");
    const btnIngresar = await findVisibleLocator(bizagiPage, btnSels, "BotonIngresar");

    const uHay = await inputUser.count().catch(() => 0) > 0;
    const pHay = await inputPass.count().catch(() => 0) > 0;
    const iVis = await btnIngresar.isVisible().catch(() => false);

    // Quick check: select-only login (no username/password inputs) + Ingresar button visible
    if (iVis && !uHay) {
      const selectSolo = bizagiPage.locator('select:visible').first();
      const selectSoloVisible = await selectSolo.isVisible().catch(() => false);
      if (selectSoloVisible) {
        console.log(`[BizagiLogin] Login solo-selector detectado. Seleccionando usuario admon y clicando Ingresar...`);
        const admonVal = await selectSolo.evaluate((el) => {
          const sel = el as HTMLSelectElement;
          const opts = Array.from(sel.options || []);
          const m = opts.find((o) => /\badmon\b/i.test(o.text) || /\badmon\b/i.test(o.value))
            || opts.find((o) => String(o.value || '').trim() || String(o.text || '').trim());
          return m?.value ?? null;
        }).catch(() => null);
        if (admonVal) {
          await selectSolo.selectOption(admonVal).catch(async () => {
            await selectSolo.evaluate((el, v) => {
              const s = el as HTMLSelectElement; s.value = String(v);
              s.dispatchEvent(new Event('input', { bubbles: true }));
              s.dispatchEvent(new Event('change', { bubbles: true }));
            }, admonVal).catch(() => {});
          });
        }
        await bizagiPage.waitForTimeout(400);
        await btnIngresar.click({ force: true }).catch(() => {});
        await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
        await bizagiPage.waitForTimeout(3000);
        const urlTrasLogin = bizagiPage.url();
        if (!/\/login\b/i.test(urlTrasLogin)) {
          console.log(`[BizagiLogin] Login solo-selector exitoso. URL='${urlTrasLogin}'`);
          return;
        }
        if (await isBizagiHomeReady(bizagiPage)) return;
        continue;
      }
    }

    if (!uHay || !pHay || !iVis) {
      const allFrames = bizagiPage.frames();
      console.log(`[BizagiLogin] DIAGNÓSTICO: U=${uHay}, P=${pHay}, I=${iVis}. Frames=${allFrames.length}.`);
      for (const [idx, f] of allFrames.entries()) {
          const inputs = await f.locator('input').all().catch(() => []);
          if (inputs.length > 0) {
              const dt = await Promise.all(inputs.map(async (inp) => {
                  const id = await inp.getAttribute('id').catch(() => 'no-id');
                  const name = await inp.getAttribute('name').catch(() => 'no-name');
                  const type = await inp.getAttribute('type').catch(() => 'no-type');
                  const v = await inp.isVisible().catch(() => false);
                  return `${type}(id=${id}, name=${name}, vis=${v})`;
              }));
              console.log(`[BizagiLogin] F#${idx} (URL=${f.url().substring(0,30)}): ${dt.join(' | ')}`);
          }
      }
      console.log(`[BizagiLogin] Nada detectado en la página.`);
      await bizagiPage.waitForTimeout(1000);
      continue;
    }

    console.log(`[BizagiLogin] Detectado U=${uHay}, P=${pHay}, I=${iVis}`);

    if (uHay) {
      const currentValUser = (await inputUser.inputValue().catch(() => "")).trim();
      if (!currentValUser || (!currentValUser.includes("admon") && !currentValUser.includes("domain\\"))) {
        console.log(`[BizagiLogin] Rellenando usuario: ${usuario}`);
        await inputUser.fill(usuario, { timeout: 5000 }).catch(async () => {
          await inputUser.evaluate((el, v) => (el as HTMLInputElement).value = v, usuario).catch(() => {});
        });
        await bizagiPage.waitForTimeout(300);
      }
    }
    if (pHay) {
      const currentValPass = (await inputPass.inputValue().catch(() => "")).trim();
      if (!currentValPass) {
        console.log(`[BizagiLogin] Rellenando contraseña...`);
        await inputPass.fill(bizagiPass, { timeout: 5000 }).catch(async () => {
          await inputPass.evaluate((el, v) => (el as HTMLInputElement).value = v, bizagiPass).catch(() => {});
        });
        await bizagiPage.waitForTimeout(300);
      }
    }

    // Caso: Select de usuario (legacy or special)
    const selectUsuario = bizagiPage.locator('select:visible').first();
    if (await selectUsuario.isVisible().catch(() => false)) {
      // Find option with "admon" as a whole word (not "adominguez"), then selectOption by value
      const admonOptionValue = await selectUsuario.evaluate((el) => {
        const sel = el as HTMLSelectElement;
        const options = Array.from(sel.options || []);
        const match = options.find((o) => /\badmon\b/i.test(o.text) || /\badmon\b/i.test(o.value))
          || options.find((o) => String(o.value || '').trim() || String(o.text || '').trim());
        return match?.value ?? null;
      }).catch(() => null);

      if (admonOptionValue) {
        await selectUsuario.selectOption(admonOptionValue).catch(async () => {
          await selectUsuario.evaluate((el, val) => {
            const sel = el as HTMLSelectElement;
            sel.value = String(val);
            sel.dispatchEvent(new Event('input', { bubbles: true }));
            sel.dispatchEvent(new Event('change', { bubbles: true }));
          }, admonOptionValue).catch(() => {});
        });
      }
    }

    const userVal = (await inputUser.inputValue().catch(() => "")).trim();
    const passVal = (await inputPass.inputValue().catch(() => "")).trim();
    const selectVisible = (await selectUsuario.isVisible().catch(() => false)) && 
                          (await selectUsuario.locator('option').count().catch(() => 0)) > 2;
    
    const inputsVisibles = (await inputUser.isVisible().catch(() => false)) || (await inputPass.isVisible().catch(() => false));

    console.log(`[BizagiLogin] Estado: userVal=${!!userVal}, passVal=${!!passVal}, iVis=${iVis}, selectVisible=${selectVisible}`);

    if ((inputsVisibles && userVal && passVal && iVis) || (!inputsVisibles && selectVisible && iVis)) {
      console.log(`[BizagiLogin] Condiciones cumplidas. Intentando ingresar...`);
      await bizagiPage.waitForTimeout(800);
      await btnIngresar.click({ force: true }).catch(() => { });
      await bizagiPage.waitForLoadState('domcontentloaded').catch(() => { });
      await bizagiPage.waitForTimeout(4000);
      // URL-based success: if no longer on /login page, assume successful login
      const urlTrasLogin = bizagiPage.url();
      if (!/\/login\b/i.test(urlTrasLogin)) {
        console.log(`[BizagiLogin] URL tras login='${urlTrasLogin}'. Asumiendo exito.`);
        return;
      }
    } else {
      console.log(`[BizagiLogin] No listo o campos vacíos. U=${!!userVal}, P=${!!passVal}, I=${iVis}`);
      await bizagiPage.waitForTimeout(1000);
    }
  }

  if (!(await isBizagiHomeReady(bizagiPage))) {
    throw new Error(`No se pudo completar login Bizagi. URL=${bizagiPage.url()}`);
  }
}

async function asegurarTabAdministrarCasos(modalCasos: Locator) {
  const tabAdministrar = modalCasos
    .locator('a,button,li,span')
    .filter({ hasText: /^Administrar Casos$/i })
    .first();
  const visible = await tabAdministrar.isVisible().catch(() => false);
  if (!visible) return;
  await tabAdministrar.scrollIntoViewIfNeeded().catch(() => { });
  await tabAdministrar.click({ force: true }).catch(() => { });
}

async function confirmarPopupAceptar(bizagiPage: Page) {
  const dialogos = bizagiPage.locator('.ui-dialog:visible, [role="dialog"]:visible, .modal:visible');
  const fin = Date.now() + 15000;

  while (Date.now() < fin) {
    const total = await dialogos.count().catch(() => 0);
    for (let i = total - 1; i >= 0; i--) {
      const dlg = dialogos.nth(i);
      const txt = ((await dlg.innerText().catch(() => '')) || '').toLowerCase();
      if (!txt.includes('seguro') && !txt.includes('cancelar') && !txt.includes('informaci')) continue;

      const btnAceptar = dlg.getByRole('button', { name: /^Aceptar$/i }).first();
      if (await btnAceptar.isVisible().catch(() => false)) {
        await btnAceptar.click({ force: true }).catch(() => { });
        return true;
      }

      const btnAceptarFallback = dlg.locator('button:has-text("Aceptar"), input[value="Aceptar"]').first();
      if (await btnAceptarFallback.isVisible().catch(() => false)) {
        await btnAceptarFallback.click({ force: true }).catch(() => { });
        return true;
      }
    }
    await bizagiPage.waitForTimeout(250);
  }

  return false;
}

export async function cancelarCasoEnBizagiDesdePortal(
  portalPage: Page,
  mpn: string,
  opts: { url?: string; usuario?: string; password?: string } = {}
) {
  const bizagiUrl = resolverBizagiUrl(opts.url);
  const bizagiUsuario = opts.usuario ?? process.env.BIZAGI_USER ?? 'domain\\admon';
  const bizagiPassword = opts.password ?? process.env.BIZAGI_PASS;

  const bizagiPage = await portalPage.context().newPage();
  try {
    await bizagiPage.goto(bizagiUrl, { waitUntil: 'domcontentloaded', timeout: 90000 });
    await entrarABizagiSiHayLogin(bizagiPage, bizagiUsuario, bizagiPassword);

    const modalCasos = await abrirCasosEnBizagi(bizagiPage);
    await asegurarTabAdministrarCasos(modalCasos);

    let inputNumeroCaso = modalCasos
      .locator('xpath=//*[contains(normalize-space(.),"Numero de Caso") or contains(normalize-space(.),"Numero del caso")]/following::input[1]')
      .first();
    if (!(await inputNumeroCaso.isVisible().catch(() => false))) {
      inputNumeroCaso = modalCasos
        .locator('input[type="text"]:visible, input:not([type]):visible')
        .first();
    }
    await inputNumeroCaso.waitFor({ state: 'visible', timeout: 20000 });
    await inputNumeroCaso.click({ force: true }).catch(() => { });
    await inputNumeroCaso.fill(mpn);

    const btnBuscar = modalCasos.getByRole('button', { name: /^Buscar$/i }).first();
    if (await btnBuscar.isVisible().catch(() => false)) {
      await btnBuscar.click({ force: true });
    } else {
      const btnBuscarFallback = modalCasos.locator('button:has-text("Buscar"), input[value="Buscar"]').first();
      await btnBuscarFallback.click({ force: true });
    }

    const rowCaso = modalCasos.locator(`tr:has-text("${mpn}")`).first();
    await rowCaso.waitFor({ state: 'visible', timeout: 30000 });

    const checkboxCaso = rowCaso.locator('input[type="checkbox"]').first();
    if (await checkboxCaso.isVisible().catch(() => false)) {
      const checked = await checkboxCaso.isChecked().catch(() => false);
      if (!checked) await checkboxCaso.check({ force: true }).catch(() => { });
    } else {
      const chkFallback = rowCaso.locator('td:first-child .ui-chkbox-box, td:first-child').first();
      await chkFallback.click({ force: true }).catch(() => { });
    }

    let btnCancelarOperacion = modalCasos.locator('button:visible').filter({ hasText: /^Cancelar$/i }).first();
    if (!(await btnCancelarOperacion.isVisible().catch(() => false))) {
      btnCancelarOperacion = modalCasos
        .locator('button:visible, input[type="button"]:visible')
        .filter({ hasText: /Cancelar casos/i })
        .first();
    }
    if (!(await btnCancelarOperacion.isVisible().catch(() => false))) {
      btnCancelarOperacion = modalCasos
        .locator('xpath=//*[contains(normalize-space(.),"Motivo")]/ancestor::*[self::div or self::fieldset][1]//button[contains(normalize-space(.),"Cancelar")]')
        .first();
    }
    await btnCancelarOperacion.waitFor({ state: 'visible', timeout: 12000 });
    await btnCancelarOperacion.click({ force: true });

    const confirmo = await confirmarPopupAceptar(bizagiPage);
    if (!confirmo) {
      throw new Error(`No se pudo confirmar con 'Aceptar' la cancelacion del caso '${mpn}' en Bizagi.`);
    }

    await modalCasos
      .getByText(new RegExp(`operaci(?:o|\\u00f3)n\\s+${mpn}\\s+ha sido cancelada`, 'i'))
      .first()
      .waitFor({ state: 'visible', timeout: 30000 })
      .catch(() => { });

    const btnTerminar = modalCasos.getByRole('button', { name: /^Terminar$/i }).first();
    if (await btnTerminar.isVisible().catch(() => false)) {
      await btnTerminar.click({ force: true }).catch(() => { });
    }

    await bizagiPage.waitForTimeout(800);
  } finally {
    await portalPage.bringToFront().catch(() => { });
    await bizagiPage.close().catch(() => { });
  }
}

export async function extraerCasoActivoMpn(page: Page) {
  const modalCasosActivos = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Casos activos/i })
    .first();

  const visible = await modalCasosActivos.isVisible().catch(() => false);
  if (!visible) return null;

  const txt = (await modalCasosActivos.innerText().catch(() => '')).trim();
  const match = txt.match(/MPN-\d+/i);
  const mpn = match ? match[0].toUpperCase() : null;

  const btnCerrar = modalCasosActivos.getByRole('button', { name: /Cerrar/i }).first();
  if (await btnCerrar.isVisible().catch(() => false)) {
    await btnCerrar.click({ force: true }).catch(() => { });
  }
  await modalCasosActivos.waitFor({ state: 'hidden', timeout: 8000 }).catch(() => { });

  return mpn;
}
