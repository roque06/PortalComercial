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
      if (await loc.isVisible().catch(() => false)) {
        await loc.scrollIntoViewIfNeeded().catch(() => {});
        const ok = await loc.click({ force: true }).then(() => true).catch(() => false);
        if (ok) return true;
      }
      
      const selector = (loc as any)._selector || loc.toString();
      if (typeof selector === 'string') {
          const realSelector = selector.replace(/^Locator@/, '');
          for (const frame of page.frames()) {
             const fLoc = frame.locator(realSelector).first();
              if (await fLoc.isVisible().catch(() => false)) {
                  await fLoc.scrollIntoViewIfNeeded().catch(() => {});
                  const ok = await fLoc.click({ force: true }).then(() => true).catch(() => false);
                  if (ok) return true;
              }
          }
      }
    }
    await page.waitForTimeout(250);
  }
  return false;
}

async function esperarPrimeroVisible(page: Page, candidatos: Locator[], timeoutMs = 12000) {
  const fin = Date.now() + timeoutMs;
  while (Date.now() < fin) {
    for (const loc of candidatos) {
      if (await loc.isVisible().catch(() => false)) {
        await loc.scrollIntoViewIfNeeded().catch(() => {});
        return loc;
      }
      const selector = (loc as any)._selector || loc.toString();
      if (typeof selector === 'string') {
          const realSelector = selector.replace(/^Locator@/, '');
          for (const frame of page.frames()) {
             const fLoc = frame.locator(realSelector).first();
              if (await fLoc.isVisible().catch(() => false)) {
                  await fLoc.scrollIntoViewIfNeeded().catch(() => {});
                  return fLoc;
              }
          }
      }
    }
    await page.waitForTimeout(250);
  }
  return null;
}

/**
 * Cierra cualquier menu o modal que pueda estar bloqueando la vista (ej. menu de usuario, admin, etc.)
 */
async function cerrarMenusBizagi(bizagiPage: Page) {
  const selectoresCierre = [
    '.ui-bizagi-wp-app-menu-item-selected',
    '.ui-selectmenu-open',
    '.ui-dialog-titlebar-close:visible',
    'button.ui-dialog-titlebar-close:visible',
    '.ui-bizagi-wp-app-inbox-grid-cases-container:visible ~ .ui-widget-overlay',
  ];

  // Si hay algun modal de usuario abierto (como el del top-right), click fuera o Escape
  await bizagiPage.keyboard.press('Escape').catch(() => {});
  const userMenuOpen = await bizagiPage.locator('.ui-bizagi-wp-app-menu-item-selected, .ui-bizagi-wp-app-user-menu:visible, ul.ui-menu:visible').first().isVisible().catch(() => false);
  if (userMenuOpen) {
      console.log(`[Bizagi] Menu detectado abierto (posiblemente usuario), cerrando...`);
      // Clic en el logo o en el centro para cerrar
      await bizagiPage.locator('.ui-bizagi-wp-app-logo, body').first().click({ position: { x: 50, y: 50 }, force: true }).catch(() => {});
      await bizagiPage.keyboard.press('Escape').catch(() => {});
      await bizagiPage.waitForTimeout(500);
  }
}

async function abrirCasosEnBizagi(bizagiPage: Page) {
  console.log(`[Bizagi] Buscando modal de casos...`);
  const modalCasos = bizagiPage
    .locator('.ui-dialog:visible, [role="dialog"]:visible, .modal:visible')
    .filter({ hasText: /Casos/i })
    .first();

  if (await modalCasos.isVisible().catch(() => false)) {
    console.log(`[Bizagi] Modal de casos ya visible.`);
    return modalCasos;
  }

  await cerrarMenusBizagi(bizagiPage);
  console.log(`[Bizagi] Intentando abrir menú Admin...`);
  const abrioAdmin = await clickPrimeroVisible(
    bizagiPage,
    [
      bizagiPage.locator('#menuListAdmin').first(),
      bizagiPage.locator('[id="mnuBt.Admin.Reports"]').first(),
      bizagiPage.locator('#mnuBt\\.Admin\\.Reports').first(),
      bizagiPage.getByRole('button', { name: /^Admin$/i }).first(),
      bizagiPage.getByRole('link', { name: /^Admin$/i }).first(),
      bizagiPage.locator('a,button,li,span').filter({ hasText: /^Admin$/i }).first(),
      bizagiPage.locator('.bz-icon-gear, .bz-icon-cog, i.fa-cog, i.pi-cog, [class*="icon-admin"]').first(),
      bizagiPage.locator('span:has-text("Admin")').first(),
      bizagiPage.locator('div:has-text("Admin")').first(),
    ],
    25000
  );
  if (!abrioAdmin) {
    console.log(`[Bizagi][ERROR] No se detectó menú Admin. URL actual: ${bizagiPage.url()}`);
    throw new Error('No se pudo abrir menu Admin en Bizagi.');
  }

  console.log(`[Bizagi] Intentando abrir Administración de procesos...`);
  const abrioAdminProcesos = await clickPrimeroVisible(
    bizagiPage,
    [
      bizagiPage.getByText(/Administraci(?:o|\u00f3)n de procesos/i).first(),
      bizagiPage.locator('li:has-text("Administracion de procesos")').first(),
      bizagiPage.locator('span:has-text("Administración de procesos")').first(),
      bizagiPage.locator('div:has-text("Administración de procesos")').first(),
    ],
    20000
  );
  if (!abrioAdminProcesos) {
    console.log(`[Bizagi][ERROR] No se detectó 'Administración de procesos'.`);
    throw new Error('No se pudo abrir Administracion de procesos en Bizagi.');
  }

  console.log(`[Bizagi] Intentando abrir opción Casos...`);
  const abrioCasos = await clickPrimeroVisible(
    bizagiPage,
    [
      bizagiPage.locator('#categories').getByText(/^Casos$/i).first(),
      bizagiPage.getByText(/^Casos$/i).first(),
      bizagiPage.locator('li:has(span.title:has-text("Casos"))').first(),
      bizagiPage.locator('li:has-text("Casos")').first(),
      bizagiPage.locator('span:has-text("Casos")').first(),
    ],
    20000
  );
  if (!abrioCasos) {
    console.log(`[Bizagi][ERROR] No se detecto opcion 'Casos'. URL: ${bizagiPage.url()}`);
    throw new Error('No se pudo abrir la opcion Casos en Bizagi.');
  }
  console.log(`[Bizagi] Opcion 'Casos' clickeada.`);

  console.log(`[Bizagi] Esperando visibilidad del modal de casos...`);
  await modalCasos.waitFor({ state: 'visible', timeout: 35000 });
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
        console.log(`[BizagiLogin] ${label} visible en Frame #${idx} con selector: ${sel}`);
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
        console.log(`[BizagiLogin] ${label} existe en Frame #${idx} con selector: ${sel}`);
        return fLoc;
      }
    }
  }

  // 3. Frame search
  for (const frame of page.frames()) {
    for (const sel of selectors) {
      const fLoc = frame.locator(sel).first();
      if (await fLoc.isVisible().catch(() => false)) {
        console.log(`[BizagiLogin] ${label} visible en frame con selector: ${sel}`);
        return fLoc;
      }
    }
  }

  // 4. Ultimate Fallback
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
  const bizagiUser = usuario || "domain\\admon";

  const userSels = ['#user', 'input#user', 'input[name="username"]', 'input#username', 'input[autocomplete="username"]', 'xpath=(//label[contains(normalize-space(.),"Usuario")]/following::input[1])[1]'];
  const passSels = ['#password', 'input#password', 'input[type="password"]', 'input[name="password"]', 'xpath=(//label[contains(normalize-space(.),"Contrase")]/following::input[1])[1]'];
  const btnSels = ['#btn-login', 'button#btn-login', 'button#btn-adminlogin', 'button:has-text("Ingresar")', 'button[name*="login" i]'];

  const inputUser = await findVisibleLocator(bizagiPage, userSels, "Usuario");
  const inputPass = await findVisibleLocator(bizagiPage, passSels, "Password");
  const btnIngresar = await findVisibleLocator(bizagiPage, btnSels, "BotonIngresar");

  const uHay = await inputUser.count().catch(() => 0) > 0;
  const pHay = await inputPass.count().catch(() => 0) > 0;
  const iVis = await btnIngresar.isVisible().catch(() => false);

  if (!uHay && !pHay && !iVis) {
    console.log(`[BizagiLogin] Nada detectado en la página o frames.`);
    return;
  }

  console.log(`[BizagiLogin] Detectado U=${uHay}, P=${pHay}, I=${iVis}`);

  if (uHay) {
    const currentValUser = (await inputUser.inputValue().catch(() => "")).trim();
    if (!currentValUser || (!currentValUser.includes("admon") && !currentValUser.includes("domain\\"))) {
      console.log(`[BizagiLogin] Rellenando usuario: ${bizagiUser}`);
      await inputUser.fill(bizagiUser, { timeout: 5000 }).catch(async () => {
        await inputUser.evaluate((el, v) => (el as HTMLInputElement).value = v, bizagiUser).catch(() => {});
      });
    }
  }
  if (pHay) {
    const currentValPass = (await inputPass.inputValue().catch(() => "")).trim();
    if (!currentValPass) {
      console.log(`[BizagiLogin] Rellenando contraseña...`);
      await inputPass.fill(bizagiPass, { timeout: 5000 }).catch(async () => {
        await inputPass.evaluate((el, v) => (el as HTMLInputElement).value = v, bizagiPass).catch(() => {});
      });
    }
  }

  const selectUsuario = bizagiPage.locator('select:visible').first();
  if (await selectUsuario.isVisible().catch(() => false)) {
    await selectUsuario
      .evaluate((el, user) => {
        const sel = el as HTMLSelectElement;
        const raw = String(user || '').toLowerCase();
        const terms = [raw, raw.replace('domain\\', ''), 'admon'];
        const options = Array.from(sel.options || []);
        const match = options.find((o) => {
          const txt = String(o.text || '').toLowerCase();
          const val = String(o.value || '').toLowerCase();
          return terms.some((t) => t && (txt.includes(t) || val.includes(t)));
        });
        if (!match) return;
        sel.value = match.value;
        sel.dispatchEvent(new Event('change', { bubbles: true }));
      }, bizagiUser)
      .catch(() => {});
  }

  const userVal = (await inputUser.inputValue().catch(() => "")).trim();
  const passVal = (await inputPass.inputValue().catch(() => "")).trim();
  const selectVisible = (await selectUsuario.isVisible().catch(() => false)) &&
                        (await selectUsuario.locator('option').count().catch(() => 0)) > 2;

  const inputsVisibles = (await inputUser.isVisible().catch(() => false)) || (await inputPass.isVisible().catch(() => false));
  const listoConUsuarioYPass = uHay && pHay && !!userVal && !!passVal && iVis;
  const listoSoloConUsuario = uHay && !pHay && !!userVal && iVis;
  const listoConSelector = !inputsVisibles && selectVisible && iVis;

  console.log(`[BizagiLogin] Estado: userVal=${!!userVal}, passVal=${!!passVal}, iVis=${iVis}, selectVisible=${selectVisible}, soloUsuario=${listoSoloConUsuario}`);

  if (listoConUsuarioYPass || listoSoloConUsuario || listoConSelector) {
    console.log(`[BizagiLogin] Condiciones cumplidas. Intentando ingresar...`);
    await bizagiPage.waitForTimeout(600);
    await btnIngresar.click({ force: true }).catch(() => { });
    await bizagiPage.waitForLoadState('domcontentloaded').catch(() => { });
    await bizagiPage.waitForTimeout(3000);
  } else {
    console.log(`[BizagiLogin] No listo o campos vacíos. U=${!!userVal}, P=${!!passVal}, I=${iVis}`);
  }
}

async function esBizagiHomePage(page: Page): Promise<boolean> {
  // If common login elements are visible, it's definitely NOT the home page
  const loginSels = [
    'input#user', 'input#username', 'input#password', 'input[name="username"]', 'input[name="password"]',
    'select#user', 'select#username', 'select[name*="user" i]',
    'button#btn-login', 'button#btn-adminlogin', 'button:has-text("Ingresar")',
    '.login-container', '.login-form', '#login-panel'
  ];
  for (const sel of loginSels) {
    if (await page.locator(sel).first().isVisible().catch(() => false)) {
      return false;
    }
  }

  const sels = [
    'input#menuQuery',
    'input[name="menuQuery"]',
    '#ui-bizagi-wp-widget-searchContainer input',
    '[id="mnuBt.Admin.Reports"]',
    '#mnuBt\\.Admin\\.Reports',
    '#main-layout',
    '.pc-avatar',
    '.user-name',
    // Inbox indicators visible in the actual Bizagi UI
    '[class*="inbox"]',
    '[class*="worklist"]',
    'a[href*="inbox"]',
    'a[href*="worklist"]',
    '#inbox',
    '.bpm-inbox',
    '[id*="inbox" i]',
    '.bz-icon-gear',
    '.bz-icon-cog',
    // Bizagi nav menu items visible when logged in
    'li.ui-bizagi-wp-app-menu-cases',
    '[id*="menuListInbox"]',
    '[id*="menuListAdmin"]',
    'table#ui-bizagi-wp-app-inbox-grid-cases',
  ];

  for (const sel of sels) {
    if (await page.locator(sel).first().isVisible().catch(() => false)) return true;
  }

  // Check for visible Bizagi nav text - very reliable indicator of being logged in
  const textosBizagiHome = [
    /Bandeja de entrada/i,
    /Gesti[oó]n de Transferencias/i,
    /Todos los casos/i,
    /Nuevo Caso/i,
    /Procesos en Vivo/i,
  ];
  for (const regex of textosBizagiHome) {
    if (await page.getByText(regex).first().isVisible().catch(() => false)) return true;
  }

  // Also check if the URL is no longer a login page
  const url = page.url();
  if (!/\/login\b/i.test(url) && /portalcomercial|bizagi/i.test(url)) {
    const textoLogin = page.getByText(/Seleccione sus datos para ingresar a Bizagi|Ingrese su usuario|Iniciar sesión|Authentication/i).first();
    const loginVisible = await textoLogin.isVisible().catch(() => false);
    if (!loginVisible) {
      return false; // Better to assume NOT home if unsure
    }
  }

  return false;
}

async function entrarABizagiSiHayLoginRobusto(bizagiPage: Page, usuario: string, password?: string) {
  const bizagiPass = password || process.env.BIZAGI_PASS || "H0la1234.";
  const bizagiUser = usuario || "domain\\admon";

  console.log(`[BizagiLogin] Iniciando verificación de login para: ${bizagiUser}`);
  if (await esBizagiHomePage(bizagiPage)) {
    console.log(`[BizagiLogin] Ya estamos en la página de inicio de Bizagi.`);
    return;
  }

  const pickVisible = async (locators: Locator[]) => {
    for (const loc of locators) {
      if (await loc.isVisible().catch(() => false)) return loc;
    }
    return locators[0];
  };

  const fin = Date.now() + 45000;
  let detectoAlgo = false;

  while (Date.now() < fin) {
    if (await esBizagiHomePage(bizagiPage)) return;

    const inputUser = await pickVisible([
      bizagiPage.locator('#user').first(),
      bizagiPage.locator('input#user').first(),
      bizagiPage.locator('input[name="username"]').first(),
      bizagiPage.locator('input#username').first(),
      bizagiPage.locator('input[autocomplete="username"]').first(),
      bizagiPage.locator('xpath=(//label[contains(normalize-space(.),"Usuario")]/following::input[1])[1]').first(),
      bizagiPage.locator('xpath=(//*[contains(normalize-space(.),"Usuario")]/following::input[1])[1]').first(),
    ]);
    const inputPass = await pickVisible([
      bizagiPage.getByLabel(/Contrase/i).first(),
      bizagiPage.locator('#password').first(),
      bizagiPage.locator('input#password').first(),
      bizagiPage.locator('input[type="password"]').first(),
      bizagiPage.locator('input[name="password"]').first(),
      bizagiPage.locator('xpath=(//label[contains(normalize-space(.),"Contrase")]/following::input[1])[1]').first(),
      bizagiPage.locator('xpath=(//*[contains(normalize-space(.),"Contrase")]/following::input[1])[1]').first(),
    ]);
    const selectUsuario = await pickVisible([
      bizagiPage.getByLabel(/Usuario/i).first(),
      bizagiPage.locator('xpath=(//label[contains(normalize-space(.),"Usuario")]/following::select[1])[1]').first(),
      bizagiPage.locator('xpath=(//*[contains(normalize-space(.),"Usuario")]/following::select[1])[1]').first(),
      bizagiPage.locator('select:visible').first(),
    ]);
    const selectDominio = await pickVisible([
      bizagiPage.getByLabel(/Dominio/i).first(),
      bizagiPage.locator('xpath=(//label[contains(normalize-space(.),"Dominio")]/following::select[1])[1]').first(),
      bizagiPage.locator('xpath=(//*[contains(normalize-space(.),"Dominio")]/following::select[1])[1]').first(),
    ]);
    const btnIngresar = await pickVisible([
      bizagiPage.getByRole('button', { name: /^Ingresar$/i }).first(),
      bizagiPage.locator('#btn-login').first(),
      bizagiPage.locator('button#btn-login').first(),
      bizagiPage.locator('button#btn-adminlogin').first(),
      bizagiPage.locator('button:has-text("Ingresar")').first(),
      bizagiPage.locator('button[name*="login" i]').first(),
      bizagiPage.locator('input[type="submit"][value*="Ingresar" i]').first(),
      bizagiPage.locator('.ui-bizagi-button-container button').first(),
      bizagiPage.locator('.login-button').first(),
      bizagiPage.locator('xpath=//span[contains(normalize-space(.),"Ingresar")]/ancestor::button[1]').first(),
      bizagiPage.locator('xpath=//div[contains(normalize-space(.),"Ingresar") and @role="button"]').first(),
    ]);

    const uVis = await inputUser.isVisible().catch(() => false);
    const pVis = await inputPass.isVisible().catch(() => false);
    const sVis = await selectUsuario.isVisible().catch(() => false);
    const dVis = await selectDominio.isVisible().catch(() => false);
    const iVis = await btnIngresar.isVisible().catch(() => false);

    if (!uVis && !pVis && !sVis && !dVis && !iVis) {
      await bizagiPage.waitForTimeout(600).catch(() => {});
      continue;
    }

    detectoAlgo = true;
    console.log(`[BizagiLogin] Detectado U=${uVis}, P=${pVis}, SU=${sVis}, D=${dVis}, I=${iVis}`);

    console.log(`[BizagiLogin] Obteniendo valores de usuario (uVis=${uVis}, sVis=${sVis})...`);
    
    const currentValUserRaw = uVis ? (await inputUser.inputValue({ timeout: 1500 }).catch(() => "")).trim() : "";
    const currentTextUserRaw = uVis ? ((await inputUser.textContent({ timeout: 1500 }).catch(() => "")) || "").trim() : "";
    
    const selectedOptionUserRaw = uVis 
      ? await inputUser.evaluate((el) => {
          try {
            const select = el as HTMLSelectElement;
            if (!select || !select.selectedOptions || !select.selectedOptions.length) return '';
            return (select.selectedOptions[0].textContent || select.selectedOptions[0].value || '').trim();
          } catch(e) { return ''; }
        }).catch(() => "")
      : "";

    const selectedOptionUsuarioRaw = sVis
      ? await selectUsuario.evaluate((el) => {
          try {
            const select = el as HTMLSelectElement;
            if (!select || !select.selectedOptions || !select.selectedOptions.length) return '';
            return (select.selectedOptions[0].textContent || select.selectedOptions[0].value || '').trim();
          } catch(e) { return ''; }
        }).catch(() => "")
      : "";
    
    const visibleUserValue = currentValUserRaw || currentTextUserRaw || selectedOptionUserRaw || selectedOptionUsuarioRaw;
    console.log(`[BizagiLogin] Valor detectado actualmente: '${visibleUserValue}'`);

    if (uVis) {
      const requiereRellenarUsuario =
        pVis
          ? (!visibleUserValue || (!visibleUserValue.includes("admon") && !visibleUserValue.includes("domain\\")))
          : !visibleUserValue;

      if (requiereRellenarUsuario) {
        // SEGURIDAD: verificar que el elemento encontrado es realmente un campo de login
        // y no el avatar/div de usuario del home de Bizagi.
        const esTagInput = await inputUser.evaluate((el) => {
          const tag = (el as HTMLElement).tagName?.toLowerCase();
          // Es valido si es input o select, o si está dentro de un formulario de login
          const enForm = !!(el as HTMLElement).closest('form, .login-form, .login-container, [id*="login"]');
          return tag === 'input' || tag === 'select' || enForm;
        }).catch(() => false);

        if (!esTagInput) {
          // El #user encontrado es el avatar del home page, NO un campo de login.
          // Esto significa que Bizagi ya está logueado. Salir del loop.
          console.log(`[BizagiLogin] Bizagi ya esta logueado (avatar detectado, no campo de login). Saliendo.`);
          await bizagiPage.waitForTimeout(500).catch(() => {});
          return;
        } else {
          console.log(`[BizagiLogin] Rellenando usuario: ${bizagiUser}`);
          await inputUser.click({ force: true, timeout: 2000 }).catch(() => {});
          const fillOk = await inputUser.fill(bizagiUser, { timeout: 4000 }).then(() => true).catch(() => false);

          if (!fillOk) {
            console.log(`[BizagiLogin][DEBUG] Fallo fill() normal (div en lugar de input). Usando pressSequentially...`);
            await inputUser.press('Control+a').catch(() => {});
            await inputUser.press('Delete').catch(() => {});
            await inputUser.pressSequentially(bizagiUser, { delay: 50 }).catch(async () => {
              await inputUser.evaluate((el, v) => {
                const h = el as HTMLElement;
                h.focus();
                if ('value' in h) (h as HTMLInputElement).value = String(v);
                else h.textContent = String(v);
                h.dispatchEvent(new Event('input', { bubbles: true }));
                h.dispatchEvent(new Event('change', { bubbles: true }));
              }, bizagiUser).catch(() => {});
            });
          }

          const finalUserVal = (await inputUser.inputValue().catch(() => '')).trim()
            || ((await inputUser.textContent().catch(() => '')) || '').trim()
            || ((await inputUser.evaluate((el) => (el as HTMLElement).innerText || '').catch(() => '')) || '').trim();
          console.log(`[BizagiLogin] Usuario rellenado. Valor final en campo: '${finalUserVal}'`);

          // Login de 2 pasos: solo presionar Enter si el campo ES un input de login real
          if (!pVis && !iVis) {
            console.log(`[BizagiLogin] Solo se detecta campo Usuario. Intentando presionar Enter para avanzar...`);
            await inputUser.press('Enter').catch(() => {});
            await bizagiPage.waitForTimeout(1000).catch(() => {});
          }
        }
      } else if (!pVis && visibleUserValue) {
        console.log(`[BizagiLogin] Usuario ya visible ('${visibleUserValue}'). Se omite rellenar y se intentara ingresar.`);
      }
    }


    if (pVis) {
      const currentValPass = (await inputPass.inputValue().catch(() => "")).trim();
      if (!currentValPass) {
        console.log(`[BizagiLogin] Rellenando contraseña...`);
        await inputPass.click({ force: true }).catch(() => {});
        await inputPass.fill(bizagiPass, { timeout: 5000 }).catch(async () => {
          await inputPass.evaluate((el, v) => {
            const input = el as HTMLInputElement;
            input.value = String(v);
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
          }, bizagiPass).catch(() => {});
        });
      }
    }

    if (sVis) {
      const selectedOptionUsuario = await selectUsuario
        .evaluate((el) => {
          const select = el as HTMLSelectElement;
          if (!select || !('selectedOptions' in select) || !select.selectedOptions?.length) return '';
          return (select.selectedOptions[0]?.textContent || select.selectedOptions[0]?.value || '').trim();
        })
        .catch(() => "");

      // Re-select if empty, default placeholder, OR if the selected user isn't the target "admon" user
      // ALSO: If it's the first time we see the select, we force a selection to trigger any JS events
      const requiereSeleccionUsuario =
        !selectedOptionUsuario ||
        /^seleccione/i.test(selectedOptionUsuario) ||
        /^select/i.test(selectedOptionUsuario) ||
        !/\badmon\b/i.test(selectedOptionUsuario) ||
        (detectoAlgo && !bizagiPage.url().includes('forceSelectDone')); // Hack to force at least once

      if (requiereSeleccionUsuario) {
        console.log(`[BizagiLogin] Seleccionando usuario desde combo (actual='${selectedOptionUsuario}')...`);
        const admonOptionValue = await selectUsuario.evaluate((el, target) => {
          const sel = el as HTMLSelectElement;
          const options = Array.from(sel.options || []);
          const rawTarget = String(target || '').toLowerCase();
          const cleanTarget = rawTarget.replace('domain\\', '');
          
          const match = options.find((o) => {
            const txt = String(o.text || '').toLowerCase();
            const val = String(o.value || '').toLowerCase();
            return txt.includes('admon') || val.includes('admon') || 
                   txt.includes(cleanTarget) || val.includes(cleanTarget);
          }) || options.find((o) => String(o.value || '').trim() && !/seleccione|select/i.test(o.text));
          
          return match?.value ?? null;
        }, bizagiUser).catch(() => null);

        if (admonOptionValue) {
          console.log(`[BizagiLogin] Seleccionando '${admonOptionValue}' vía evaluate...`);
          await selectUsuario.evaluate((el, val) => {
            const sel = el as HTMLSelectElement;
            sel.value = String(val);
            sel.dispatchEvent(new Event('input', { bubbles: true }));
            sel.dispatchEvent(new Event('change', { bubbles: true }));
          }, admonOptionValue).catch(() => {});
          
          // Also try Playwright's selectOption as backup but with short timeout
          await selectUsuario.selectOption(admonOptionValue, { timeout: 2000 }).catch(() => {});
        }
        await bizagiPage.waitForTimeout(1000).catch(() => {});

        if (iVis) {
          console.log(`[BizagiLogin] Intentando ingresar inmediatamente tras selección...`);
          await btnIngresar.click({ force: true, timeout: 5000 }).catch(() => {});
          await btnIngresar.evaluate((el) => (el as HTMLElement).click()).catch(() => {});
          await bizagiPage.keyboard.press('Enter').catch(() => {});
          
          await bizagiPage.waitForTimeout(2000).catch(() => {});
          if (await esBizagiHomePage(bizagiPage)) return;
        }
      } else {
        console.log(`[BizagiLogin] Usuario ya seleccionado ('${selectedOptionUsuario}').`);
        if (iVis) {
          console.log(`[BizagiLogin] Intentando ingresar (ya estaba seleccionado)...`);
          await btnIngresar.click({ force: true, timeout: 3000 }).catch(() => {});
          await btnIngresar.evaluate((el) => (el as HTMLElement).click()).catch(() => {});
          await bizagiPage.keyboard.press('Enter').catch(() => {});
        }
      }
    }

    if (dVis) {
      const totalOptions = await selectDominio.locator('option').count().catch(() => 0);
      if (totalOptions > 0) {
        await selectDominio.selectOption({ label: 'BMSC' }).catch(async () => {
          await selectDominio.selectOption({ value: 'BMSC' }).catch(() => {});
        });
      }
    }

    const userVal = ((await inputUser.inputValue().catch(() => "")).trim()) || visibleUserValue;
    const passVal = (await inputPass.inputValue().catch(() => "")).trim();
    const selectUserVal = await selectUsuario
      .evaluate((el) => {
        try {
          const select = el as HTMLSelectElement;
          if (!select || !select.selectedOptions || !select.selectedOptions.length) return '';
          return (select.selectedOptions[0].textContent || select.selectedOptions[0].value || '').trim();
        } catch(e) { return ''; }
      })
      .catch(() => "");

    console.log(`[BizagiLogin] Estado final: uVis=${uVis}, pVis=${pVis}, sVis=${sVis}, iVis=${iVis}, userVal='${userVal}', selectUserVal='${selectUserVal}'`);

    if (!uVis && !pVis && sVis && iVis) {
      console.log(`[BizagiLogin] Intentando ingresar por selector de usuario (selectUserVal='${selectUserVal}')...`);
      await bizagiPage.waitForTimeout(600).catch(() => {});
      
      let clicked = false;
      try {
        await btnIngresar.click({ force: true, timeout: 5000 });
        clicked = true;
      } catch (e) {
        console.log(`[BizagiLogin] Click falló, intentando evaluate click...`);
        await btnIngresar.evaluate((el) => (el as HTMLElement).click()).catch(() => {});
        clicked = true;
      }

      await bizagiPage.waitForLoadState('domcontentloaded').catch(() => { });
      await bizagiPage.waitForTimeout(3000).catch(() => {});

      if (await esBizagiHomePage(bizagiPage)) return;
      const urlTrasLogin = bizagiPage.url();
      if (!/\/login\b/i.test(urlTrasLogin) && /portalcomercial|bizagi/i.test(urlTrasLogin)) {
        console.log(`[BizagiLogin] URL tras login='${urlTrasLogin}'. Asumiendo éxito.`);
        return;
      }
    }

    if (uVis && !pVis && iVis && userVal) {
      console.log(`[BizagiLogin] Intentando ingresar con usuario visible y sin contraseña...`);
      await bizagiPage.waitForTimeout(600).catch(() => {});
      await btnIngresar.click({ force: true }).catch(() => { });
      await btnIngresar.evaluate((el) => (el as HTMLElement).click()).catch(() => {});
      
      await bizagiPage.waitForLoadState('domcontentloaded').catch(() => { });
      await bizagiPage.waitForTimeout(3000).catch(() => {});

      if (await esBizagiHomePage(bizagiPage)) return;
    }

    if (uVis && pVis && iVis && userVal && passVal) {
      console.log(`[BizagiLogin] Condiciones cumplidas. Intentando ingresar...`);
      await bizagiPage.waitForTimeout(400).catch(() => {});
      await btnIngresar.click({ force: true }).catch(() => { });
      await bizagiPage.waitForLoadState('domcontentloaded').catch(() => { });
      await bizagiPage.waitForTimeout(2500).catch(() => {});

      if (await esBizagiHomePage(bizagiPage)) return;
    }

    await bizagiPage.waitForTimeout(700).catch(() => {});
  }

  if (!detectoAlgo) {
    console.log(`[BizagiLogin] Nada detectado en la página o frames.`);
  } else {
    console.log(`[BizagiLogin][WARN] Se detecto el formulario de login pero no se logro completar el acceso.`);
  }
}

async function asegurarTabAdministrarCasos(modalCasos: Locator) {
  const tabAdministrar = modalCasos
    .locator('a,button,li,span')
    .filter({ hasText: /^Administrar Casos$/i })
    .first();
  const visible = await tabAdministrar.isVisible().catch(() => false);
  if (!visible) return;
  await tabAdministrar.scrollIntoViewIfNeeded().catch(() => {});
  await tabAdministrar.click({ force: true }).catch(() => {});
}

async function confirmarPopupAceptar(bizagiPage: Page, timeoutMs = 15000) {
  const dialogos = bizagiPage.locator('.ui-dialog:visible, [role="dialog"]:visible, .modal:visible');
  const fin = Date.now() + timeoutMs;

  while (Date.now() < fin) {
    const total = await dialogos.count().catch(() => 0);
    for (let i = total - 1; i >= 0; i--) {
      const dlg = dialogos.nth(i);
      const txt = ((await dlg.innerText().catch(() => '')) || '').toLowerCase();
      if (!txt.includes('seguro') && !txt.includes('cancelar') && !txt.includes('informaci')) continue;

      const btnAceptar = dlg.getByRole('button', { name: /^Aceptar$/i }).first();
      if (await btnAceptar.isVisible().catch(() => false)) {
        await btnAceptar.click({ force: true }).catch(() => {});
        return true;
      }

      const btnAceptarFallback = dlg.locator('button:has-text("Aceptar"), input[value="Aceptar"]').first();
      if (await btnAceptarFallback.isVisible().catch(() => false)) {
        await btnAceptarFallback.click({ force: true }).catch(() => {});
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
    await bizagiPage.bringToFront().catch(() => {});
    await bizagiPage.goto(bizagiUrl, { waitUntil: 'domcontentloaded', timeout: 90000 });
    await bizagiPage.bringToFront().catch(() => {});
    await entrarABizagiSiHayLoginRobusto(bizagiPage, bizagiUsuario, bizagiPassword);

    console.log(`[Bizagi][Cancelar] Intentando cancelar caso ${mpn}...`);
    const modalCasos = await abrirCasosEnBizagi(bizagiPage);
    console.log(`[Bizagi][Cancelar] Modal de casos abierto. Asegurando tab 'Administrar Casos'...`);
    await asegurarTabAdministrarCasos(modalCasos);

    console.log(`[Bizagi][Cancelar] Buscando input de numero de caso...`);
    let inputNumeroCaso = modalCasos.locator('#caseInput').first();
    if (!(await inputNumeroCaso.isVisible().catch(() => false))) {
        inputNumeroCaso = modalCasos
            .locator('xpath=//*[contains(normalize-space(.),"Numero de Caso") or contains(normalize-space(.),"Numero del caso")]/following::input[1]')
            .first();
    }
    if (!(await inputNumeroCaso.isVisible().catch(() => false))) {
      inputNumeroCaso = modalCasos
        .locator('input[type="text"]:visible, input:not([type]):visible')
        .first();
    }
    
    await inputNumeroCaso.waitFor({ state: 'visible', timeout: 20000 });

    // Bizagi acepta el numero sin prefijo (ej. "1309" en vez de "MPN-1309")
    // Intentar con numero solo primero, luego con MPN completo si no encuentra la fila
    const soloNumero = mpn.replace(/^[a-zA-Z]+-/i, '').trim();

    // Funcion auxiliar: escribe el termino en el campo y ejecuta busqueda
    const ejecutarBusqueda = async (termino: string) => {
      await inputNumeroCaso.click({ force: true }).catch(() => {});
      await inputNumeroCaso.fill(termino).catch(async () => {
        await inputNumeroCaso.evaluate((el, v) => {
          const inp = el as HTMLInputElement;
          inp.value = String(v);
          inp.dispatchEvent(new Event('input', { bubbles: true }));
          inp.dispatchEvent(new Event('change', { bubbles: true }));
        }, termino).catch(() => {});
      });
      const valorActual = (await inputNumeroCaso.inputValue().catch(() => '')).trim();
      console.log(`[Bizagi][Cancelar] Buscando '${termino}' (valor en campo: '${valorActual}')...`);
      const btnBuscar = modalCasos.getByRole('button', { name: /^Buscar$/i }).first();
      if (await btnBuscar.isVisible().catch(() => false)) {
        await btnBuscar.click({ force: true });
      } else {
        await modalCasos.locator('button:has-text("Buscar"), input[value="Buscar"]').first().click({ force: true }).catch(() => {});
      }
    };

    // 1er intento: buscar con el MPN completo (ej. "MPN-1309") y esperar hasta 30s
    await ejecutarBusqueda(mpn);
    let rowCaso = modalCasos.locator(`tr:has-text("${mpn}"), tr:has-text("${soloNumero}")`).first();
    let busquedaExitosa = await rowCaso.waitFor({ state: 'visible', timeout: 30000 }).then(() => true).catch(() => false);

    // 2do intento: buscar con solo el numero (ej. "1309") si el primero fallo
    if (!busquedaExitosa && soloNumero !== mpn) {
      console.log(`[Bizagi][Cancelar][WARN] Sin resultados para '${mpn}'. Reintentando con '${soloNumero}'...`);
      await ejecutarBusqueda(soloNumero);
      rowCaso = modalCasos.locator(`tr:has-text("${soloNumero}")`).first();
      busquedaExitosa = await rowCaso.waitFor({ state: 'visible', timeout: 20000 }).then(() => true).catch(() => false);
    }

    if (!busquedaExitosa) {
      const textGrid = await modalCasos.innerText().catch(() => '');
      console.log(`[Bizagi][Cancelar][ERROR] No se encontro fila del caso '${mpn}'. Grid: "${textGrid.substring(0, 500)}"`);
      throw new Error(`[CRITICO] No se encontro en Bizagi la fila del caso '${mpn}' para cancelar.`);
    }
    console.log(`[Bizagi][Cancelar] Fila encontrada. Marcando checkbox...`);

    const checkboxCaso = rowCaso.locator('input[name="CaseAdmin"], input[type="checkbox"]').first();
    if (await checkboxCaso.isVisible().catch(() => false)) {
      const checked = await checkboxCaso.isChecked().catch(() => false);
      if (!checked) await checkboxCaso.check({ force: true }).catch(() => {});
    } else {
      const chkFallback = rowCaso.locator('td:first-child .ui-chkbox-box, td:first-child, .ui-chkbox').first();
      await chkFallback.click({ force: true }).catch(() => {});
    }

    console.log(`[Bizagi][Cancelar] Buscando boton Cancelar...`);
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
    console.log(`[Bizagi][Cancelar] Clic en Cancelar...`);
    await btnCancelarOperacion.click({ force: true });

    console.log(`[Bizagi][Cancelar] Esperando popup de confirmacion...`);
    const confirmo = await confirmarPopupAceptar(bizagiPage);
    if (!confirmo) {
      console.log(`[Bizagi][Cancelar][ERROR] No se pudo confirmar la cancelacion.`);
      throw new Error(`No se pudo confirmar con 'Aceptar' la cancelacion del caso '${mpn}' en Bizagi.`);
    }
    console.log(`[Bizagi][Cancelar] Confirmacion aceptada.`);

    await modalCasos
      .locator('button:has-text("Terminar"), button:has-text("Finalizar"), button[aria-label="Cerrar"]')
      .first()
      .click({ force: true })
      .catch(() => {});

    // Opcional: esperar que el modal se cierre
    await modalCasos.waitFor({ state: 'hidden', timeout: 10000 }).catch(() => {});

    const btnTerminar = modalCasos.getByRole('button', { name: /^Terminar$/i }).first();
    if (await btnTerminar.isVisible().catch(() => false)) {
      await btnTerminar.click({ force: true }).catch(() => {});
    }

    console.log(`[Bizagi][Cancelar] Caso ${mpn} cancelado exitosamente.`);
    await bizagiPage.waitForTimeout(800);
  } finally {
    await portalPage.bringToFront().catch(() => {});
    await bizagiPage.close().catch(() => {});
  }
}

async function cerrarModalCasosActivosSiVisible(page: Page, modal?: Locator) {
  const modalCasosActivos = modal ?? page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Casos?\s+activos?|caso activo|solicitud activa|caso ya existe|ya existe un caso/i })
    .first();

  const visibleInicial = await modalCasosActivos.isVisible().catch(() => false);
  if (!visibleInicial) return true;

  for (let intento = 1; intento <= 4; intento++) {
    const btnCerrar = modalCasosActivos.getByRole('button', { name: /Cerrar/i }).first();
    const btnCerrarFallback = modalCasosActivos
      .locator('button:has-text("Cerrar"), .p-dialog-header-close, .p-dialog-titlebar-close, [aria-label*="cerr" i], [aria-label*="close" i]')
      .first();

    if (await btnCerrar.isVisible().catch(() => false)) {
      await btnCerrar.click({ force: true }).catch(() => {});
    } else if (await btnCerrarFallback.isVisible().catch(() => false)) {
      await btnCerrarFallback.click({ force: true }).catch(() => {});
    } else {
      await page.keyboard.press('Escape').catch(() => {});
    }

    await modalCasosActivos
      .evaluate((el) => {
        const root = el as HTMLElement;
        const candidates = Array.from(
          root.querySelectorAll('button, [role="button"], .p-dialog-header-close, .p-dialog-titlebar-close')
        ) as HTMLElement[];
        const target = candidates.find((node) => /cerrar|close/i.test(node.innerText || node.getAttribute('aria-label') || ''));
        target?.click();
      })
      .catch(() => {});

    await page.waitForTimeout(250).catch(() => {});
    const sigueVisible = await modalCasosActivos.isVisible().catch(() => false);
    console.log(`[CasoActivo] cerrar modal intento=${intento} visible=${sigueVisible}`);
    if (!sigueVisible) return true;
  }

  return !(await modalCasosActivos.isVisible().catch(() => false));
}

export async function extraerCasoActivoMpn(page: Page) {
  const modalCasosActivos = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Casos?\s+activos?|caso activo|solicitud activa|caso ya existe|ya existe un caso/i })
    .first();

  const visible = await modalCasosActivos.isVisible().catch(() => false);
  if (!visible) return null;

  let identificador: string | null = null;
  let txt = '';
  const inicio = Date.now();
  while (Date.now() - inicio < 2500 && !identificador) {
    txt = (await modalCasosActivos.innerText().catch(() => '')).trim();
    const matchMpn = txt.match(/MPN-\d+/i);
    const matchCaso = txt.match(/(?:n[uú]mero\s+de\s+caso|caso|solicitud)[^\d]{0,20}(\d{4,})/i);
    const matchNumeroSuelto = txt.match(/\b\d{4,}\b/);
    identificador = matchMpn
      ? matchMpn[0].toUpperCase()
      : matchCaso?.[1]
        ? matchCaso[1]
        : matchNumeroSuelto?.[0]
          ? matchNumeroSuelto[0]
          : null;
    if (!identificador) {
      await page.waitForTimeout(180);
    }
  }

  console.log(`[CasoActivo] Modal detectado. identificador=${identificador ?? 'no-encontrado'}`);

  const cerro = await cerrarModalCasosActivosSiVisible(page, modalCasosActivos);
  if (!cerro) {
    console.log('[CasoActivo][WARN] No se pudo cerrar el modal de casos activos antes de continuar.');
  }

  return identificador;
}

async function navegarATodosLosCasosBizagi(bizagiPage: Page) {
  const linkTodosLosCasos = bizagiPage.getByRole('link', { name: /Todos\s+los\s+casos/i }).first();
  if (await linkTodosLosCasos.isVisible().catch(() => false)) {
    await linkTodosLosCasos.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(1500);
    return true;
  }

  const candidatos = [
    bizagiPage.locator('a,span,div,li,[role="menuitem"],[role="treeitem"]').filter({ hasText: /^Todos\s+los\s+casos$/i }).first(),
    bizagiPage.locator('text=/^Todos\\s+los\\s+casos$/i').first(),
  ];

  return clickPrimeroVisible(bizagiPage, candidatos, 12000);
}

async function navegarABandejaEntradaBizagi(bizagiPage: Page) {
  await cerrarMenusBizagi(bizagiPage);
  const candidatos = [
    bizagiPage.getByRole('link', { name: /Bandeja de entrada/i }).first(),
    bizagiPage.getByRole('button', { name: /Bandeja de entrada/i }).first(),
    bizagiPage.locator('a,button,span,div').filter({ hasText: /^Bandeja de entrada$/i }).first(),
  ];

  const hizoClick = await clickPrimeroVisible(bizagiPage, candidatos, 5000).catch(() => false);
  if (hizoClick) {
    await bizagiPage.waitForTimeout(600).catch(() => {});
  }
  return hizoClick;
}

async function buscarCasoPorMpnEnBizagi(bizagiPage: Page, mpn: string, opts?: { useAdminFallback?: boolean; openIfFound?: boolean }) {
  await navegarABandejaEntradaBizagi(bizagiPage).catch(() => false);

  const selectoresBuscador = [
    '#menuQuery',
    'input#menuQuery',
    '#ui-bizagi-wp-widget-searchContainer input',
    'input[id*="menuQuery"]',
    'input[name="menuQuery"]',
    'input[placeholder="Buscar"]',
  ];

  const dispararBusqueda = async (input: Locator, modo: 'rapido' | 'humano' = 'rapido') => {
    // Asegurar que no hay menus tapando el buscador
    await cerrarMenusBizagi(bizagiPage).catch(() => {});
    
    await input.focus().catch(() => {});
    // Usar fill directamente para evitar el clic central que puede dar al avatar si estan pegados
    await input.fill(mpn).catch(async () => {
        await input.click({ timeout: 1000 }).catch(() => {});
        await input.fill(mpn).catch(() => {});
    });

    if (modo === 'humano') {
      await input.pressSequentially(mpn, { delay: 45 }).catch(async () => {
        await input.fill(mpn).catch(() => {});
      });
    } else {
      await input.fill(mpn).catch(async () => {
        await input.press('Control+A').catch(() => {});
        await input.press('Delete').catch(() => {});
        await input.pressSequentially(mpn, { delay: 12 }).catch(() => {});
      });
    }

    const valor = ((await input.inputValue().catch(() => '')) || '').trim();
    if (valor !== mpn) {
      await input.evaluate((el, v) => {
        const inputEl = el as HTMLInputElement;
        inputEl.value = String(v);
        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
        inputEl.dispatchEvent(new Event('change', { bubbles: true }));
      }, mpn).catch(() => {});
    }

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
    }).catch(() => {});

    await input.press('Enter').catch(() => {});
    await bizagiPage.keyboard.press('Enter').catch(() => {});

    // Solo si Enter no funciono y el valor no cambio, intentamos el icono
    const valorPostEnter = ((await input.inputValue().catch(() => '')) || '').trim();
    if (valorPostEnter === mpn) {
        const triggerBusqueda = input
          .locator('xpath=(preceding-sibling::*[contains(@class,"search") or contains(@class,"icon")] | following-sibling::*[contains(@class,"search") or contains(@class,"icon")])[1]')
          .first();
        if (await triggerBusqueda.isVisible().catch(() => false)) {
          console.log(`[Cumplimiento][Bizagi] Click en icono de busqueda para ${mpn}`);
          await triggerBusqueda.click({ timeout: 3000 }).catch(() => {});
        }
    }

    await bizagiPage.waitForTimeout(modo === 'humano' ? 700 : 350);
  };

  let buscadorUsado = false;
  let inputBuscador: Locator | null = null;

  // Intentar encontrar el buscador en la pagina principal o frames
  const inputBuscadorDetectado = await findVisibleLocator(bizagiPage, selectoresBuscador, "BuscadorSuperior");
  if (await inputBuscadorDetectado.isVisible().catch(() => false)) {
      console.log(`[Cumplimiento][Bizagi] buscando ${mpn} en el buscador superior detectado...`);
      await dispararBusqueda(inputBuscadorDetectado, 'rapido');
      buscadorUsado = true;
      inputBuscador = inputBuscadorDetectado;
  }

  if (!buscadorUsado) {
    console.log('[Cumplimiento][Bizagi] buscador superior no encontrado o no visible.');
    if (opts?.useAdminFallback) {
      console.log('[Cumplimiento][Bizagi] Intentando buscar via Admin -> Procesos -> Casos (Fallback)...');
      try {
        const modalCasos = await abrirCasosEnBizagi(bizagiPage);
        await asegurarTabAdministrarCasos(modalCasos);
        
        let inputNumeroCaso = modalCasos
          .locator('xpath=//*[contains(normalize-space(.),"Numero de Caso") or contains(normalize-space(.),"Numero del caso")]/following::input[1]')
          .first();
        if (!(await inputNumeroCaso.isVisible().catch(() => false))) {
          inputNumeroCaso = modalCasos.locator('input[type="text"]:visible').first();
        }
        await inputNumeroCaso.fill(mpn);
        
        const btnBuscar = modalCasos.getByRole('button', { name: /^Buscar$/i }).first();
        await btnBuscar.click({ force: true }).catch(() => {});
        
        const rowCaso = modalCasos.locator(`tr:has-text("${mpn}")`).first();
        const encontrado = await rowCaso.isVisible().catch(() => false);
        if (encontrado) {
          console.log(`[Cumplimiento][Bizagi] Caso ${mpn} encontrado en modal Admin.`);
          if (opts?.openIfFound) {
              console.log(`[Cumplimiento][Bizagi] Intentando abrir caso ${mpn} desde modal Admin...`);
              const btnVerDetalle = rowCaso.locator('a, button, [role="button"]').first();
              await btnVerDetalle.click({ force: true }).catch(() => {});
              await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
          }
          return true;
        }
      } catch (e) {
        console.log(`[Cumplimiento][Bizagi][WARN] Fallback de Admin falló: ${e}`);
      }
    }
    return false;
  }

  const filasPorMpn = bizagiPage
    .locator('table#ui-bizagi-wp-app-inbox-grid-cases tbody tr[data-idworkflow]:visible')
    .filter({ hasText: new RegExp(mpn, 'i') });

  const inicio = Date.now();
  const TIMEOUT_ESPERA = 20000;
  let ultimoSegundoLog = -1;
  let reintentosDisparoCount = 0;
  const maxReintentosDisparo = 3;

  while (Date.now() - inicio < TIMEOUT_ESPERA) {
    const countMpn = await filasPorMpn.count().catch(() => 0);
    if (countMpn > 0) {
      console.log(`[Cumplimiento][Bizagi] fila con ${mpn} encontrada (count=${countMpn})`);
      return true;
    }

    const elapsedMs = Date.now() - inicio;
    // Reintentar disparar la busqueda cada 5 segundos si no hay resultados
    if (inputBuscador && elapsedMs > (reintentosDisparoCount + 1) * 5000 && reintentosDisparoCount < maxReintentosDisparo) {
      reintentosDisparoCount++;
      console.log(`[Cumplimiento][Bizagi] Reintento de búsqueda #${reintentosDisparoCount} para ${mpn}...`);
      await cerrarMenusBizagi(bizagiPage).catch(() => {});
      await dispararBusqueda(inputBuscador, 'humano');
    }

    const elapsed = Math.floor(elapsedMs / 1000);
    if (elapsed !== ultimoSegundoLog) {
      ultimoSegundoLog = elapsed;
      console.log(`[Cumplimiento][Bizagi] esperando resultados para ${mpn} (${elapsed}s)...`);
    }
    await bizagiPage.waitForTimeout(500);
  }

  return false;
}

async function abrirPrimeraSolicitudCumplimientoBizagi(bizagiPage: Page, mpn: string) {
  const filasInbox = bizagiPage
    .locator('table#ui-bizagi-wp-app-inbox-grid-cases tbody tr[data-idworkflow]');

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

  const filaGestionDebidaPlaft = filasInbox
    .filter({ hasText: new RegExp(mpn, 'i') })
    .filter({ hasText: /Gesti[oó]n Debida Diligencia PLAFT/i })
    .first();

  let fila = filaGestionCoincidencias;
  let origenFila = 'Gestionar Coincidencias';
  let visibleFila = await fila.isVisible().catch(() => false);
  if (!visibleFila) {
    fila = filaVerificarPlaft;
    origenFila = 'Verificar PLAFT';
    visibleFila = await fila.isVisible().catch(() => false);
  }
  if (!visibleFila) {
    fila = filaGestionDebidaPlaft;
    origenFila = 'Gestión Debida Diligencia PLAFT';
    visibleFila = await fila.isVisible().catch(() => false);
  }
  if (!visibleFila) {
    fila = filaPlaftGenerica;
    origenFila = 'PLAFT';
    visibleFila = await fila.isVisible().catch(() => false);
  }

  if (!visibleFila) {
    throw new Error(`[CRITICO] No se encontro en Bizagi una fila util de Cumplimiento para '${mpn}'.`);
  }

  console.log(`[Cumplimiento][Bizagi] intentando abrir fila origen='${origenFila}' para ${mpn}`);
  await fila.scrollIntoViewIfNeeded().catch(() => {});

  const xpathTextoMpn = `contains(normalize-space(translate(., '\u00A0', ' ')), '${mpn}')`;
  const selectorCeldaNumeroCaso = 'td._I_RadNumber, td._i_RadNumber, td.I_RadNumber, td.i_RadNumber, td[class*="RadNumber"]';
  const selectorSpanNumeroCaso = `${selectorCeldaNumeroCaso} span.ui-bizagi-wp-app-inbox-activity-name`;
  const spanNumeroCaso = fila.locator(
    `xpath=.//td[contains(@class,'RadNumber')]//span[contains(@class,'ui-bizagi-wp-app-inbox-activity-name') and ${xpathTextoMpn}]`
  ).first();
  const celdaNumeroCaso = fila.locator(
    `xpath=.//td[contains(@class,'RadNumber') and .//span[${xpathTextoMpn}]]`
  ).first();
  const celdaProceso = fila.locator('td._i_idWFClass').first();
  const textoNumeroCaso = fila.getByText(new RegExp(`^${mpn}$`, 'i')).first();
  const linkNumeroCaso = fila.locator(`${selectorSpanNumeroCaso}, ${selectorCeldaNumeroCaso} a, ${selectorCeldaNumeroCaso}`).first();

  const candidatosApertura =
    origenFila === 'PLAFT' || origenFila === 'Gestión Debida Diligencia PLAFT'
      ? [
          spanNumeroCaso,
          textoNumeroCaso,
          celdaNumeroCaso,
          linkNumeroCaso,
          celdaProceso,
          fila,
        ]
      : [
          spanNumeroCaso,
          textoNumeroCaso,
          celdaNumeroCaso,
          linkNumeroCaso,
          celdaProceso,
          fila,
        ];

  let ultimoError = '';
  let indiceCandidato = 0;
  for (const candidato of candidatosApertura) {
    indiceCandidato++;
    const visible = await candidato.isVisible().catch(() => false);
    if (!visible) continue;
    await candidato.scrollIntoViewIfNeeded().catch(() => {});
    console.log(`[Cumplimiento][Bizagi] probando candidato ${indiceCandidato}/${candidatosApertura.length} para abrir ${origenFila}`);

    const acciones: Array<() => Promise<void>> = [
      async () => { await candidato.click(); },
      async () => { await candidato.click({ force: true }); },
      async () => { await candidato.dblclick(); },
      async () => { await candidato.dblclick({ force: true }); },
      async () => {
        const box = await candidato.boundingBox();
        if (!box) throw new Error('sin bounding box');
        await bizagiPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { clickCount: 1 });
      },
      async () => {
        await candidato.evaluate((el) => {
          const h = el as HTMLElement;
          h.focus?.();
          h.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          h.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          h.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        });
      },
      async () => {
        await candidato.evaluate((el) => {
          const h = el as HTMLElement;
          h.focus?.();
          h.click?.();
          h.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        });
      },
      async () => {
        await candidato.focus().catch(() => {});
        await candidato.press('Enter').catch(() => {});
      },
    ];

    for (const accion of acciones) {
      try {
        await accion();
        await bizagiPage.waitForTimeout(250);
        const inicio = Date.now();
        while (Date.now() - inicio < 12000) {
          const tieneComboAccion = await bizagiPage.locator('input[role="combobox"], input.ui-select-data').first().isVisible().catch(() => false);
          const tieneAclaraciones = await bizagiPage.getByText(/Solicitar Aclaraciones/i).first().isVisible().catch(() => false);
          const tieneSiguiente = await bizagiPage.getByRole('button', { name: /Siguiente/i }).first().isVisible().catch(() => false);
          const tieneGestionCoincidencias = await bizagiPage.getByText(/Gesti[oó]n de Coincidencias/i).first().isVisible().catch(() => false);
          const tieneOtrasCoincidencias = await bizagiPage.getByText(/Otras Coincidencias/i).first().isVisible().catch(() => false);
          if (tieneComboAccion || tieneAclaraciones || tieneSiguiente || tieneGestionCoincidencias || tieneOtrasCoincidencias) {
            return;
          }
          await bizagiPage.waitForTimeout(300);
        }
      } catch (e) {
        ultimoError = e instanceof Error ? e.message : String(e);
      }
    }
  }

  throw new Error(`[CRITICO] No se pudo abrir la solicitud de Cumplimiento para '${mpn}' desde la fila encontrada.${ultimoError ? ` Detalle: ${ultimoError}` : ''}`);
}

async function abrirComboSelectBizagi(bizagiPage: Page, combo: Locator, scope?: Locator) {
  await combo.scrollIntoViewIfNeeded().catch(() => {});
  await combo.click({ force: true }).catch(() => {});
  await combo.focus().catch(() => {});

  const contenedor = scope ?? combo.locator('xpath=ancestor::*[self::div or self::td or self::tr][1]').first();
  const btnCombo = contenedor
    .locator('.ui-selectmenu-btn, .ui-selectmenu-button, .ui-selectmenu-trigger, div.ui-selectmenu-btn')
    .first();

  if (await btnCombo.isVisible().catch(() => false)) {
    await btnCombo.click({ force: true }).catch(() => {});
  } else {
    await combo.press('ArrowDown').catch(() => {});
  }
}

async function seleccionarOpcionComboBizagi(
  bizagiPage: Page,
  combo: Locator,
  opcionRegex: RegExp,
  verifyRegex: RegExp,
  opts?: { scope?: Locator; descripcion?: string }
) {
  const descripcion = opts?.descripcion ?? String(opcionRegex);

  for (let intento = 1; intento <= 4; intento++) {
    await abrirComboSelectBizagi(bizagiPage, combo, opts?.scope).catch(() => {});

    const opcion = bizagiPage
      .locator('li, .ui-menuitem, .ui-selectmenu-item, [role="option"]')
      .filter({ hasText: opcionRegex })
      .first();

    const visible = await opcion.waitFor({ state: 'visible', timeout: 2500 }).then(() => true).catch(() => false);
    if (!visible) {
      await combo.press('ArrowDown').catch(() => {});
      await bizagiPage.waitForTimeout(200);
      const visibleReintento = await opcion.waitFor({ state: 'visible', timeout: 1500 }).then(() => true).catch(() => false);
      if (!visibleReintento) {
        console.log(`[Cumplimiento][Bizagi] No abrió la lista para ${descripcion} (intento ${intento}/4).`);
        await bizagiPage.keyboard.press('Escape').catch(() => {});
        await bizagiPage.waitForTimeout(250);
        continue;
      }
    }

    await opcion.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(300);

    const valorFinal = ((await combo.inputValue().catch(() => '')) || '').trim().toLowerCase();
    if (verifyRegex.test(valorFinal)) return true;
  }

  return false;
}

async function comboBizagiEstaHabilitado(combo: Locator) {
  return await combo.evaluate((el) => {
    const input = el as HTMLInputElement;
    const wrapper = input.closest('.ui-bizagi-control, .ui-bizagi-control-wrapper, .ui-select, .ui-select-data-container, .ui-selectmenu');
    const trigger = wrapper?.querySelector('.ui-selectmenu-btn, .ui-selectmenu-button, .ui-selectmenu-trigger') as HTMLElement | null;
    const classNames = [
      input.className || '',
      wrapper?.className || '',
      trigger?.className || '',
    ].join(' ');

    return !input.disabled
      && input.getAttribute('aria-disabled') !== 'true'
      && trigger?.getAttribute('aria-disabled') !== 'true'
      && !/disabled|state-disabled|ui-state-disabled|p-disabled/i.test(classNames);
  }).catch(() => false);
}

async function esperarComboBizagiHabilitado(
  bizagiPage: Page,
  combo: Locator,
  descripcion: string,
  timeoutMs = 8000,
) {
  const inicio = Date.now();
  while (Date.now() - inicio < timeoutMs) {
    if (await comboBizagiEstaHabilitado(combo)) return true;
    await bizagiPage.waitForTimeout(200);
  }

  const habilitadoFinal = await comboBizagiEstaHabilitado(combo);
  if (!habilitadoFinal) {
    console.log(`[Cumplimiento][Bizagi] ${descripcion} sigue deshabilitado tras ${timeoutMs}ms.`);
  }
  return habilitadoFinal;
}

async function seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage: Page, context: Page | import('@playwright/test').Frame | Locator) {
  // --- INTENTO PRIORITARIO (Basado en grabación exitosa) ---
  const comboGrabado = context.getByRole('combobox', { name: 'Motivo Coincidencias OFAC:' }).first();
  if (await comboGrabado.isVisible().catch(() => false)) {
      console.log('[Cumplimiento][Bizagi] Aplicando secuencia grabada para Motivo OFAC (Triple Click)...');
      // El usuario indico que necesita 3 clics para que el combo abra de forma fiable en su ambiente
      await comboGrabado.click({ force: true }).catch(() => {});
      await bizagiPage.waitForTimeout(150);
      await comboGrabado.click({ force: true }).catch(() => {});
      await bizagiPage.waitForTimeout(150);
      await comboGrabado.click({ force: true }).catch(() => {});
      
      const opcionGrabada = bizagiPage.getByRole('option', { 
          name: 'Coincidencia descartada por no corresponderse con la persona incluida en las Listas de Control Internas de Clientes', 
          exact: true 
      }).first();
      
      if (await opcionGrabada.isVisible({ timeout: 5000 }).catch(() => false)) {
          await opcionGrabada.click({ force: true }).catch(() => {});
          console.log('[Cumplimiento][Bizagi] Motivo OFAC seleccionado exitosamente vía grabación.');
          return;
      } else {
          console.log('[Cumplimiento][Bizagi][WARN] No se desplegaron las opciones tras el triple click.');
          // Intentar ArrowDown como fallback inmediato
          await comboGrabado.press('ArrowDown').catch(() => {});
          await bizagiPage.waitForTimeout(300);
          if (await opcionGrabada.isVisible().catch(() => false)) {
              await opcionGrabada.click({ force: true }).catch(() => {});
              return;
          }
      }
  }
  // ---------------------------------------------------------

  const bloqueMotivo = context
    .locator('xpath=//*[contains(normalize-space(.),"Motivo Coincidencias OFAC")]/ancestor::*[self::div or self::td or self::section or contains(@class,"ui-bizagi-render")][1]')
    .first();
  await bloqueMotivo.waitFor({ state: 'visible', timeout: 10000 });

  let comboMotivo = context.getByRole('combobox', { name: 'Motivo Coincidencias OFAC:' }).first();
  if (!(await comboMotivo.isVisible().catch(() => false))) {
    comboMotivo = context
      .locator('xpath=(//*[contains(normalize-space(.),"Motivo Coincidencias OFAC")]/following::input[contains(@class,"ui-select-data") or @role="combobox"])[1]')
      .first();
  }
  if (!(await comboMotivo.isVisible().catch(() => false))) {
    comboMotivo = context
      .locator('input.ui-select-data.ui-selectmenu-value[role="combobox"], input.ui-select-data[role="combobox"], input.ui-selectmenu-value[role="combobox"], input[id*="combo"]')
      .first();
  }
  if (!(await comboMotivo.isVisible().catch(() => false))) {
    comboMotivo = context.locator('xpath=.//following::input[@role="combobox"][1]').first();
  }
  await comboMotivo.waitFor({ state: 'visible', timeout: 10000 });
  const comboHabilitado = await esperarComboBizagiHabilitado(
    bizagiPage,
    comboMotivo,
    'Motivo Coincidencias OFAC',
    8000,
  );
  if (!comboHabilitado) {
    throw new Error("[CRITICO] 'Motivo Coincidencias OFAC' sigue deshabilitado; 'Descartar' no quedo aplicado.");
  }
  const botonMotivo = bizagiPage
    .locator('xpath=(//*[contains(normalize-space(.),"Motivo Coincidencias OFAC")]/following::div[contains(@class,"ui-selectmenu-btn") or contains(@class,"ui-selectmenu-button") or contains(@class,"ui-selectmenu-trigger")])[1]')
    .first();

  const valorActual = ((await comboMotivo.inputValue().catch(() => '')) || '').trim().toLowerCase();
  if (valorActual && !/seleccione|please select/i.test(valorActual)) {
    console.log(`[Cumplimiento][Bizagi] Motivo Coincidencias OFAC ya tenia valor='${valorActual}'.`);
    return;
  }

  let seleccionado = false;
  for (let intento = 1; intento <= 4 && !seleccionado; intento++) {
    await abrirComboSelectBizagi(bizagiPage, comboMotivo, bloqueMotivo).catch(() => {});
    if (await botonMotivo.isVisible().catch(() => false)) {
      await botonMotivo.click({ force: true }).catch(() => {});
    }
    const opciones = bizagiPage.locator('li, .ui-menuitem, .ui-selectmenu-item, [role="option"]');
    const hayOpciones = await opciones.first().waitFor({ state: 'visible', timeout: 2500 }).then(() => true).catch(() => false);
    if (!hayOpciones) {
      await comboMotivo.press('ArrowDown').catch(() => {});
      await bizagiPage.waitForTimeout(200);
      const hayOpcionesTrasTecla = await opciones.first().waitFor({ state: 'visible', timeout: 1500 }).then(() => true).catch(() => false);
      if (!hayOpcionesTrasTecla) {
        await bizagiPage.keyboard.press('Escape').catch(() => {});
        await bizagiPage.waitForTimeout(200);
        continue;
      }
    }

    const total = await opciones.count().catch(() => 0);
    if (total > 0) {
      let indiceObjetivo = -1;
      let textoObjetivo = '';
      const textoExactoGrabado = 'Coincidencia descartada por no corresponderse con la persona incluida en las Listas de Control Internas de Clientes';
      
      for (let idx = 0; idx < total; idx++) {
        const texto = ((await opciones.nth(idx).innerText().catch(() => '')) || '').trim();
        if (texto.includes(textoExactoGrabado)) {
            indiceObjetivo = idx;
            textoObjetivo = texto;
            break;
        }
      }

      if (indiceObjetivo === -1) {
          for (let idx = 0; idx < total; idx++) {
            const texto = ((await opciones.nth(idx).innerText().catch(() => '')) || '').trim();
            if (!/seleccione|please select/i.test(texto)) {
              indiceObjetivo = idx;
              textoObjetivo = texto;
              break;
            }
          }
      }

      if (indiceObjetivo === -1) indiceObjetivo = 0;

      await opciones.nth(indiceObjetivo).click({ force: true }).catch(() => {});
      await bizagiPage.waitForTimeout(250);
      await bizagiPage.locator('body').click({ position: { x: 5, y: 5 } }).catch(() => {});
      await bizagiPage.waitForTimeout(150);
      const valorFinal = ((await comboMotivo.inputValue().catch(() => '')) || '').trim().toLowerCase();
      seleccionado = !!valorFinal && !/seleccione|please select/i.test(valorFinal);
      if (seleccionado) {
        console.log(`[Cumplimiento][Bizagi] Motivo Coincidencias OFAC seleccionado='${textoObjetivo || valorFinal}'`);
      }
    }
  }

  console.log("[Cumplimiento][Bizagi] Motivo Coincidencias OFAC = index 1 (primer valor real)");
}

function obtenerConfiguracionAccionOfac() {
  return {
    nombre: 'Descartar',
    valor: '2',
    regex: /Descartar/i,
  };
}

async function seleccionarRadioAccionOfac(
  bizagiPage: Page,
  scope: Page | Locator,
  accionNombre: string,
  accionValor: string,
  accionRegex: RegExp,
) {
  let radioAccion = scope.locator(
    `xpath=(.//input[@type="radio" and contains(@id,"sidP_AccOFAC") and @value="${accionValor}"] | .//input[@type="radio" and @value="${accionValor}"])[1]`
  ).first();
  if (!(await radioAccion.count().catch(() => 0))) {
    radioAccion = bizagiPage.locator(
      `xpath=(//input[@type="radio" and contains(@id,"sidP_AccOFAC") and @value="${accionValor}"] | //input[@type="radio" and @value="${accionValor}"])[1]`
    ).first();
  }

  if (!(await radioAccion.count().catch(() => 0))) {
    throw new Error(`[CRITICO] No se encontro el radio '${accionNombre}' en 'Acción Coincidencias OFAC'.`);
  }

  const radioId = (await radioAccion.getAttribute('id').catch(() => null)) || '';
  let labelAccion = radioId
    ? bizagiPage.locator(`label[for="${radioId}"]`).filter({ hasText: accionRegex }).first()
    : scope.locator('label').filter({ hasText: accionRegex }).first();
  if (!(await labelAccion.isVisible().catch(() => false))) {
    labelAccion = bizagiPage.locator('label').filter({ hasText: accionRegex }).first();
  }

  const contenedorRadio = radioAccion
    .locator('xpath=ancestor::*[contains(@class,"ui-radio") or contains(@class,"ui-bizagi-render-radio-item")][1]')
    .first();

  const accionActiva = async () => {
    const inputChecked = await radioAccion.isChecked().catch(() => false);
    const ariaChecked = await radioAccion.getAttribute('aria-checked').catch(() => null);
    const claseRadio = await contenedorRadio.evaluate((el) => {
      const h = el as HTMLElement;
      return h.className || '';
    }).catch(() => '');
    const claseLabel = await labelAccion.evaluate((el) => {
      const h = el as HTMLElement;
      return h.className || '';
    }).catch(() => '');
    return inputChecked
      || ariaChecked === 'true'
      || /active|checked|highlight|filled|selected/i.test(claseRadio)
      || /active|checked|highlight|filled|selected/i.test(claseLabel);
  };

  let seleccionado = await accionActiva();
  for (let intento = 1; intento <= 4 && !seleccionado; intento++) {
    console.log(`[Cumplimiento][Bizagi] Marcando ${accionNombre} directo intento ${intento}/4`);
    if (await labelAccion.isVisible().catch(() => false)) {
      await labelAccion.scrollIntoViewIfNeeded().catch(() => {});
      await labelAccion.click({ force: true }).catch(() => {});
      await labelAccion.evaluate((el) => {
        const h = el as HTMLElement;
        h.click?.();
        h.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      }).catch(() => {});
      const box = await labelAccion.boundingBox().catch(() => null);
      if (box) {
        await bizagiPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2).catch(() => {});
      }
    }
    if (!seleccionado && await contenedorRadio.isVisible().catch(() => false)) {
      await contenedorRadio.scrollIntoViewIfNeeded().catch(() => {});
      await contenedorRadio.click({ force: true }).catch(() => {});
    }
    if (!seleccionado) {
      await radioAccion.click({ force: true }).catch(() => {});
      await radioAccion.check({ force: true }).catch(() => {});
      await radioAccion.evaluate((el) => {
        const input = el as HTMLInputElement;
        input.checked = true;
        input.setAttribute('checked', 'checked');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      }).catch(() => {});
    }
    await bizagiPage.waitForTimeout(300);
    seleccionado = await accionActiva();
  }

  if (!seleccionado) {
    throw new Error(`[CRITICO] No se pudo seleccionar '${accionNombre}' en 'Acción Coincidencias OFAC'.`);
  }

  console.log(`[Cumplimiento][Bizagi] Acción Coincidencias OFAC = ${accionNombre}`);
}

async function seleccionarRadioAccionOfacPorLabel(
  bizagiPage: Page,
  scope: Page | Locator | import('@playwright/test').Frame,
  accionNombre: string,
  accionValor: string,
  accionRegex: RegExp,
) {
  let labelAccion = scope.locator(
    `xpath=(.//label[contains(@for,"sidP_AccOFAC-${accionValor}") and normalize-space()="${accionNombre}"] | .//label[normalize-space()="${accionNombre}"])[1]`
  ).first();
  if (!(await labelAccion.isVisible().catch(() => false))) {
    labelAccion = bizagiPage.locator(
      `xpath=(//label[contains(@for,"sidP_AccOFAC-${accionValor}") and normalize-space()="${accionNombre}"] | //label[normalize-space()="${accionNombre}"])[1]`
    ).first();
  }
  if (!(await labelAccion.isVisible().catch(() => false))) {
    labelAccion = bizagiPage.locator('label').filter({ hasText: accionRegex }).first();
  }
  if (!(await labelAccion.isVisible().catch(() => false))) {
    throw new Error(`[CRITICO] No se encontro el label '${accionNombre}' en 'Acción Coincidencias OFAC'.`);
  }

  const leerInputAsociado = async () => {
    return await bizagiPage.evaluate(({ nombre, valor }) => {
      const normalizar = (texto?: string | null) => (texto || '').replace(/\u00A0/g, ' ').trim().toLowerCase();
      const nombreNormalizado = normalizar(nombre);
      const labels = Array.from(document.querySelectorAll('label'));
      const label = (labels.find((el) => {
        const htmlFor = (el as HTMLLabelElement).htmlFor || '';
        return normalizar(el.textContent) === nombreNormalizado && htmlFor.includes(`sidP_AccOFAC-${valor}`);
      }) || labels.find((el) => normalizar(el.textContent) === nombreNormalizado)) as HTMLLabelElement | undefined;
      const input = label?.htmlFor
        ? (document.getElementById(label.htmlFor) as HTMLInputElement | null)
        : ((label?.previousElementSibling as HTMLInputElement | null) ?? label?.closest('span')?.querySelector('input[type="radio"]') as HTMLInputElement | null);
      const contenedor = (label?.closest('.ui-radio, .ui-bizagi-render-radio-item, span') as HTMLElement | null)
        ?? (input?.closest('.ui-radio, .ui-bizagi-render-radio-item, span') as HTMLElement | null);
      return {
        checked: !!input?.checked,
        className: input?.className || '',
        labelClassName: label?.className || '',
        contenedorClassName: contenedor?.className || '',
      };
    }, { nombre: accionNombre, valor: accionValor }).catch(() => ({
      checked: false,
      className: '',
      labelClassName: '',
      contenedorClassName: '',
    }));
  };

  const contenedorRadio = labelAccion
    .locator('xpath=ancestor::*[contains(@class,"ui-radio") or contains(@class,"ui-bizagi-render-radio-item")][1]')
    .first();

  const marcarPorDom = async () => {
    return await bizagiPage.evaluate(({ nombre, valor }) => {
      const normalizar = (texto?: string | null) => (texto || '').replace(/\u00A0/g, ' ').trim().toLowerCase();
      const nombreNormalizado = normalizar(nombre);
      const labels = Array.from(document.querySelectorAll('label'));
      const label = (labels.find((el) => {
        const htmlFor = (el as HTMLLabelElement).htmlFor || '';
        return normalizar(el.textContent) === nombreNormalizado && htmlFor.includes(`sidP_AccOFAC-${valor}`);
      }) || labels.find((el) => normalizar(el.textContent) === nombreNormalizado)) as HTMLLabelElement | undefined;
      const input = label?.htmlFor
        ? (document.getElementById(label.htmlFor) as HTMLInputElement | null)
        : ((label?.previousElementSibling as HTMLInputElement | null) ?? label?.closest('span')?.querySelector('input[type="radio"]') as HTMLInputElement | null);
      const contenedor = (label?.closest('.ui-radio, .ui-bizagi-render-radio-item, span') as HTMLElement | null)
        ?? (input?.closest('.ui-radio, .ui-bizagi-render-radio-item, span') as HTMLElement | null);

      const dispararMouse = (nodo?: HTMLElement | null) => {
        if (!nodo) return;
        nodo.focus?.();
        nodo.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        nodo.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        nodo.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      };

      dispararMouse(contenedor);
      dispararMouse(label);
      input?.focus?.();
      input?.click?.();
      if (input) {
        input.checked = true;
        input.setAttribute('checked', 'checked');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      }
      return !!input?.checked;
    }, { nombre: accionNombre, valor: accionValor }).catch(() => false);
  };

  const accionActiva = async () => {
    const inputInfo = await leerInputAsociado();
    const claseRadio = await contenedorRadio.evaluate((el) => {
      const h = el as HTMLElement;
      return h.className || '';
    }).catch(() => '');
    const claseLabel = await labelAccion.evaluate((el) => {
      const h = el as HTMLElement;
      return h.className || '';
    }).catch(() => '');
    return inputInfo.checked
      || /active|checked|highlight|filled|selected/i.test(inputInfo.className || '')
      || /active|checked|highlight|filled|selected/i.test(inputInfo.labelClassName || '')
      || /active|checked|highlight|filled|selected/i.test(inputInfo.contenedorClassName || '')
      || /active|checked|highlight|filled|selected/i.test(claseRadio)
      || /active|checked|highlight|filled|selected/i.test(claseLabel);
  };

  let seleccionado = await accionActiva();
  for (let intento = 1; intento <= 4 && !seleccionado; intento++) {
    console.log(`[Cumplimiento][Bizagi] Marcando ${accionNombre} por label intento ${intento}/4`);
    await labelAccion.scrollIntoViewIfNeeded().catch(() => {});
    await labelAccion.click({ force: true }).catch(() => {});
    await labelAccion.evaluate((el) => {
      const label = el as HTMLLabelElement;
      const input = label.htmlFor
        ? (document.getElementById(label.htmlFor) as HTMLInputElement | null)
        : ((label.previousElementSibling as HTMLInputElement | null) ?? label.closest('span')?.querySelector('input[type="radio"]') as HTMLInputElement | null);
      label.click?.();
      label.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      if (input) {
        input.checked = true;
        input.setAttribute('checked', 'checked');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      }
    }).catch(() => {});
    const box = await labelAccion.boundingBox().catch(() => null);
    if (box) {
      await bizagiPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2).catch(() => {});
    }
    if (!seleccionado && await contenedorRadio.isVisible().catch(() => false)) {
      await contenedorRadio.click({ force: true }).catch(() => {});
    }
    if (!seleccionado) {
      await marcarPorDom().catch(() => false);
    }
    await bizagiPage.waitForTimeout(300);
    seleccionado = await accionActiva();
  }

  if (!seleccionado) {
    throw new Error(`[CRITICO] No se pudo seleccionar '${accionNombre}' en 'Acción Coincidencias OFAC' por label.`);
  }

  console.log(`[Cumplimiento][Bizagi] Acción Coincidencias OFAC = ${accionNombre} (label)`);
}

async function manejarCoincidenciasOfacDirectoBizagi(bizagiPage: Page) {
  // --- INTENTO PRIORITARIO (Basado en grabación exitosa) ---
  const btnDescartarGrabado = bizagiPage.getByText('Descartar').first();
  if (await btnDescartarGrabado.isVisible().catch(() => false)) {
      console.log('[Cumplimiento][Bizagi] Pulsando "Descartar" (grabado)...');
      await btnDescartarGrabado.click({ force: true }).catch(() => {});
      await bizagiPage.waitForTimeout(800); // Espera a que se habilite el combo de motivo
  }
  
  await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, bizagiPage);
  
  // El usuario grabo que despues de seleccionar el motivo debe dar Siguiente y Aceptar
  const btnSiguiente = bizagiPage.getByRole('button', { name: 'Siguiente' }).first();
  if (await btnSiguiente.isVisible().catch(() => false)) {
      console.log('[Cumplimiento][Bizagi] Haciendo click en Siguiente (secuencia OFAC)...');
      await btnSiguiente.click({ force: true }).catch(() => {});
      await bizagiPage.waitForTimeout(800);
      
      const btnAceptar = bizagiPage.getByRole('button', { name: 'Aceptar' }).first();
      if (await btnAceptar.isVisible().catch(() => false)) {
          console.log('[Cumplimiento][Bizagi] Haciendo click en Aceptar (confirmacion OFAC)...');
          await btnAceptar.click({ force: true }).catch(() => {});
      }
  }
  // ---------------------------------------------------------

  const radioDescartar = bizagiPage.locator('input[type="radio"][id*="sidP_AccOFAC"][value="2"]').first();
  const comboMotivo = bizagiPage
    .locator('xpath=(//*[contains(normalize-space(.),"Motivo Coincidencias OFAC")]/following::input[@role="combobox"] | //*[contains(normalize-space(.),"Motivo Coincidencias OFAC")]/following::input[contains(@class,"ui-select-data")])[1]')
    .first();

  let visibleRadioDescartar = !!(await radioDescartar.count().catch(() => 0)) || await bizagiPage.locator('label').filter({ hasText: /Descartar/i }).first().isVisible().catch(() => false);
  const visibleComboMotivo = await comboMotivo.isVisible().catch(() => false) || await bizagiPage.getByText(/Motivo Coincidencias OFAC/i).first().isVisible().catch(() => false);
  if (visibleComboMotivo && !visibleRadioDescartar) {
    const fin = Date.now() + 4000;
    while (Date.now() < fin && !visibleRadioDescartar) {
      await bizagiPage.waitForTimeout(200);
      visibleRadioDescartar = !!(await radioDescartar.count().catch(() => 0)) || await bizagiPage.locator('label').filter({ hasText: /Descartar/i }).first().isVisible().catch(() => false);
    }
  }
  console.log(`[Cumplimiento][Bizagi] OFAC visible radio=${visibleRadioDescartar} motivo=${visibleComboMotivo}`);
  if (!visibleRadioDescartar && !visibleComboMotivo) return false;

  const accionOfac = obtenerConfiguracionAccionOfac();
  await seleccionarRadioAccionOfacPorLabel(
    bizagiPage,
    bizagiPage,
    accionOfac.nombre,
    accionOfac.valor,
    accionOfac.regex,
  );
  await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, bizagiPage.locator('body').first());
  return true;
}

async function manejarCoincidenciasOfacConfirmarXpathBizagi(bizagiPage: Page) {
  const accionOfac = obtenerConfiguracionAccionOfac();
  
  // 1. Intentar encontrar el contexto (puede ser la página o un frame)
  let context: Page | import('@playwright/test').Frame = bizagiPage;
  const selectorsLabels = [
    'text=/Acci[oó]n Coincidencias OFAC/i',
    'text=/Motivo Coincidencias OFAC/i',
    'text=/Coincidencias OFAC/i'
  ];

  let foundInFrame = false;
  for (const frame of bizagiPage.frames()) {
    for (const sel of selectorsLabels) {
      if (await frame.locator(sel).first().isVisible().catch(() => false)) {
        context = frame;
        foundInFrame = true;
        break;
      }
    }
    if (foundInFrame) break;
  }

  console.log(`[Cumplimiento][Bizagi] Contexto detectado en ${foundInFrame ? 'Frame' : 'Main Page'}.`);

  // 2. Localizar radio 'Descartar' (o el configurado)
  const findAccion = async () => {
    const locators = [
        context.locator('input[type="radio"][value="2"]').first(),
        context.locator('input[type="radio"][id*="sidP_AccOFAC-2"]').first(),
        context.locator('label').filter({ hasText: /^Descartar$/i }).first(),
        context.locator('xpath=//label[contains(normalize-space(.),"Descartar")]').first()
    ];
    for (const loc of locators) {
        if (await loc.isVisible().catch(() => false)) return loc;
    }
    return null;
  };

  const radioDescartar = await findAccion();
  console.log(`[Cumplimiento][Bizagi] Radio Descartar detectado: ${!!radioDescartar}`);

  if (radioDescartar) {
      await radioDescartar.scrollIntoViewIfNeeded().catch(() => {});
      await radioDescartar.click({ force: true }).catch(() => {});
      await radioDescartar.check({ force: true }).catch(() => {});
  }

  // 3. Localizar y seleccionar Motivo
  const findCombo = async () => {
    const locators = [
        context.locator('input[role="combobox"][id*="combo"]').first(),
        context.locator('input.ui-select-data').first(),
        context.locator('xpath=(//*[contains(normalize-space(.),"Motivo Coincidencias OFAC")]/following::input[@role="combobox"])[1]').first(),
        context.locator('xpath=(//*[contains(normalize-space(.),"Motivo Coincidencias OFAC")]/following::input[contains(@class,"ui-select-data")])[1]').first()
    ];
    for (const loc of locators) {
        if (await loc.isVisible().catch(() => false)) return loc;
    }
    return null;
  };

  const comboMotivo = await findCombo();
  console.log(`[Cumplimiento][Bizagi] Combo Motivo detectado: ${!!comboMotivo}`);

  if (comboMotivo) {
      await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, context.locator('body').first());
  }

  // 4. Intentar Solicitar Aclaraciones = No
  try {
      await seleccionarNoSolicitarAclaracionesCumplimiento(context as any);
  } catch (e) {
      console.log(`[Cumplimiento][Bizagi][WARN] No se pudo marcar Aclaraciones: ${e.message}`);
  }

  // 5. Boton Siguiente (Final)
  const btnSiguiente = context.getByRole('button', { name: /Siguiente/i }).first();
  const visibleSiguiente = await btnSiguiente.isVisible().catch(() => false);
  console.log(`[Cumplimiento][Bizagi] Boton Siguiente visible: ${visibleSiguiente}`);

  if (visibleSiguiente) {
      await btnSiguiente.click({ force: true }).catch(() => {});
      return true;
  }

  return !!radioDescartar || !!comboMotivo;
}

async function seleccionarAccionesCumplimientoBizagi(bizagiPage: Page) {
  const manejoOfacConfirmar = await manejarCoincidenciasOfacConfirmarXpathBizagi(bizagiPage);
  if (manejoOfacConfirmar) return;

  const manejoOfacDirecto = await manejarCoincidenciasOfacDirectoBizagi(bizagiPage);
  if (manejoOfacDirecto) return;

  const radiosOfacAdjuntos = await bizagiPage.locator('input[type="radio"][id*="sidP_AccOFAC"]').count().catch(() => 0);
  const labelDescartarVisible = await bizagiPage.locator('label').filter({ hasText: /Descartar/i }).first().isVisible().catch(() => false);
  const labelMotivoVisible = await bizagiPage.getByText(/Motivo Coincidencias OFAC/i).first().isVisible().catch(() => false);
  if (radiosOfacAdjuntos > 0 || labelDescartarVisible || labelMotivoVisible) {
    const accionOfac = obtenerConfiguracionAccionOfac();
    console.log(`[Cumplimiento][Bizagi] OFAC detectado por red de seguridad. radios=${radiosOfacAdjuntos} descartarVisible=${labelDescartarVisible} motivoVisible=${labelMotivoVisible}`);
    await seleccionarRadioAccionOfacPorLabel(
      bizagiPage,
      bizagiPage,
      accionOfac.nombre,
      accionOfac.valor,
      accionOfac.regex,
    );
    await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, bizagiPage.locator('body').first());
    return;
  }

  const bloqueAccionOfacGlobal = bizagiPage
    .locator('xpath=//*[contains(normalize-space(.),"AcciÃ³n Coincidencias OFAC") or contains(normalize-space(.),"Accion Coincidencias OFAC")]/ancestor::*[self::div or self::td or self::section][1]')
    .first();
  const visibleAccionOfacGlobal = await bloqueAccionOfacGlobal.isVisible().catch(() => false);
  if (visibleAccionOfacGlobal) {
    console.log('[Cumplimiento][Bizagi] OFAC detectado por bloque global.');
    const radiosGlobal = bloqueAccionOfacGlobal.locator('input[type="radio"]');
    const totalRadiosGlobal = await radiosGlobal.count().catch(() => 0);
    if (totalRadiosGlobal > 0) {
      const radioDescartarGlobal = totalRadiosGlobal > 1 ? radiosGlobal.nth(1) : radiosGlobal.first();
      const labelDescartarGlobal = bloqueAccionOfacGlobal.locator('label').filter({ hasText: /Descartar/i }).first();
      const cajaDescartarGlobal = bloqueAccionOfacGlobal
        .locator('.ui-radiobutton-box, .p-radiobutton-box, [role="radio"]')
        .nth(totalRadiosGlobal > 1 ? 1 : 0);

      await labelDescartarGlobal.click({ force: true }).catch(() => {});
      await cajaDescartarGlobal.click({ force: true }).catch(() => {});
      await radioDescartarGlobal.check({ force: true }).catch(async () => {
        await radioDescartarGlobal.click({ force: true }).catch(() => {});
      });
      await radioDescartarGlobal.evaluate((el) => {
        const input = el as HTMLInputElement;
        input.checked = true;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }).catch(() => {});
      await bizagiPage.waitForTimeout(250);

      console.log('[Cumplimiento][Bizagi] Acción Coincidencias OFAC = Descartar (global)');
      await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, bizagiPage.locator('body').first());
      return;
    }
  }

  const seccionCoincidencias = bizagiPage
    .locator('xpath=//*[contains(normalize-space(.),"Gestión de Coincidencias") or (contains(normalize-space(.),"Gesti") and contains(normalize-space(.),"Coincidencias"))]/ancestor::*[self::div or self::section][1]')
    .first();
  await seccionCoincidencias.waitFor({ state: 'visible', timeout: 20000 });
  const seccionCoincidenciasEstrica = bizagiPage
    .locator('xpath=//*[contains(normalize-space(.),"Coincidencias") and (contains(normalize-space(.),"AcciÃ³n") or contains(normalize-space(.),"Solicitar Aclaraciones"))]/ancestor::*[self::div or self::section][1]')
    .first();
  const usarSeccionEstrica = await seccionCoincidenciasEstrica.isVisible().catch(() => false);
  const seccionObjetivo = usarSeccionEstrica ? seccionCoincidenciasEstrica : seccionCoincidencias;

  const bloqueAccionOfacFallback = seccionObjetivo
    .locator('xpath=//*[contains(normalize-space(.),"Acción Coincidencias OFAC") or contains(normalize-space(.),"Accion Coincidencias OFAC")]/ancestor::*[self::div or self::td or self::section][1]')
    .first();
  const bloqueAccionOfac2 = bloqueAccionOfacFallback;
  const visibleAccionOfacFallback = await bloqueAccionOfacFallback.isVisible().catch(() => false);
  if (visibleAccionOfacFallback) {
    const radiosOfac = bloqueAccionOfacFallback.locator('input[type="radio"]');
    const totalRadiosOfac = await radiosOfac.count().catch(() => 0);
    if (totalRadiosOfac > 0) {
      const radioDescartarOfac = totalRadiosOfac > 1 ? radiosOfac.nth(1) : radiosOfac.first();
      const labelDescartarOfac = bloqueAccionOfacFallback.locator('label').filter({ hasText: /Descartar/i }).first();
      const cajaDescartarOfac = bloqueAccionOfacFallback
        .locator('.ui-radiobutton-box, .p-radiobutton-box, [role="radio"]')
        .nth(totalRadiosOfac > 1 ? 1 : 0);
      const descartarActivo = async () => {
        const inputChecked = await radioDescartarOfac.isChecked().catch(() => false);
        const cajaActiva = await cajaDescartarOfac.evaluate((el) => {
          const h = el as HTMLElement;
          return /active|checked|highlight|filled/.test(h.className || '');
        }).catch(() => false);
        const ariaChecked = await cajaDescartarOfac.getAttribute('aria-checked').catch(() => null);
        return inputChecked || cajaActiva || ariaChecked === 'true';
      };
      let radioMarcado = await descartarActivo();

      for (let intento = 1; intento <= 4 && !radioMarcado; intento++) {
        if (await labelDescartarOfac.isVisible().catch(() => false)) {
          await labelDescartarOfac.scrollIntoViewIfNeeded().catch(() => {});
          await labelDescartarOfac.click({ force: true }).catch(() => {});
        }

        if (!radioMarcado && await cajaDescartarOfac.isVisible().catch(() => false)) {
          await cajaDescartarOfac.scrollIntoViewIfNeeded().catch(() => {});
          await cajaDescartarOfac.click({ force: true }).catch(() => {});
        }

        if (!radioMarcado && await radioDescartarOfac.isVisible().catch(() => false)) {
          await radioDescartarOfac.scrollIntoViewIfNeeded().catch(() => {});
          await radioDescartarOfac.check({ force: true }).catch(async () => {
            await radioDescartarOfac.click({ force: true }).catch(() => {});
          });
        }

        if (!radioMarcado) {
          await radioDescartarOfac.evaluate((el) => {
            const input = el as HTMLInputElement;
            input.checked = true;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            input.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          }).catch(() => {});
        }

        await bizagiPage.waitForTimeout(200);
        radioMarcado = await descartarActivo();
      }

      if (!radioMarcado) {
        throw new Error("[CRITICO] No se pudo seleccionar 'Descartar' en 'AcciÃ³n Coincidencias OFAC'.");
      }

      console.log("[Cumplimiento][Bizagi] AcciÃ³n Coincidencias OFAC = Descartar (radio)");
      await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, seccionObjetivo);
      return;
    }

    let comboOfac = bloqueAccionOfacFallback
      .locator('input.ui-select-data.ui-selectmenu-value[role="combobox"], input.ui-select-data[role="combobox"], input.ui-selectmenu-value[role="combobox"]')
      .first();
    if (!(await comboOfac.isVisible().catch(() => false))) {
      comboOfac = bloqueAccionOfacFallback.locator('xpath=.//following::input[@role="combobox"][1]').first();
    }
    await comboOfac.waitFor({ state: 'visible', timeout: 10000 });

    const valorActualOfac = ((await comboOfac.inputValue().catch(() => '')) || '').trim().toLowerCase();
    if (!valorActualOfac.includes('descartar')) {
      const seleccionadoOfac = await seleccionarOpcionComboBizagi(
        bizagiPage,
        comboOfac,
        /Descartar/i,
        /descartar/,
        { scope: bloqueAccionOfac2, descripcion: 'AcciÃ³n Coincidencias OFAC=Descartar' }
      );
      if (!seleccionadoOfac) {
        throw new Error("[CRITICO] No se pudo seleccionar 'Descartar' en 'AcciÃ³n Coincidencias OFAC'.");
      }
    }

    console.log("[Cumplimiento][Bizagi] AcciÃ³n Coincidencias OFAC = Descartar");
    await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, seccionObjetivo);
    return;
  }

  const bloqueAccionOfac = seccionObjetivo
    .locator('xpath=//*[contains(normalize-space(.),"Acción Coincidencias OFAC") or contains(normalize-space(.),"Accion Coincidencias OFAC")]/ancestor::*[self::div or self::td or self::section][1]')
    .first();
  const visibleAccionOfac = await bloqueAccionOfac.isVisible().catch(() => false);

  if (visibleAccionOfac) {
    let comboOfac = bloqueAccionOfac
      .locator('input.ui-select-data.ui-selectmenu-value[role="combobox"], input.ui-select-data[role="combobox"], input.ui-selectmenu-value[role="combobox"]')
      .first();
    if (!(await comboOfac.isVisible().catch(() => false))) {
      comboOfac = bloqueAccionOfac.locator('xpath=.//following::input[@role="combobox"][1]').first();
    }
    await comboOfac.waitFor({ state: 'visible', timeout: 10000 });

    const valorActualOfac = ((await comboOfac.inputValue().catch(() => '')) || '').trim().toLowerCase();
    if (!valorActualOfac.includes('descartar')) {
      const seleccionadoOfac = await seleccionarOpcionComboBizagi(
        bizagiPage,
        comboOfac,
        /Descartar/i,
        /descartar/,
        { scope: bloqueAccionOfac, descripcion: 'Acción Coincidencias OFAC=Descartar' }
      );
      if (!seleccionadoOfac) {
        throw new Error("[CRITICO] No se pudo seleccionar 'Descartar' en 'Acción Coincidencias OFAC'.");
      }
    }

    console.log("[Cumplimiento][Bizagi] Acción Coincidencias OFAC = Descartar");
    await seleccionarMotivoCoincidenciasOfacCumplimiento(bizagiPage, seccionObjetivo);
    return;
  }

  const combos = seccionObjetivo.locator(
    'input.ui-select-data.ui-selectmenu-value[role="combobox"], input.ui-select-data[role="combobox"], input.ui-selectmenu-value[role="combobox"]'
  );

  const totalCombos = await combos.count().catch(() => 0);
  console.log(`[Cumplimiento][Bizagi] combos de Acción detectados=${totalCombos}`);
  if (!totalCombos) {
    throw new Error("[CRITICO] No se encontraron campos editables en la columna 'Acción' de Cumplimiento.");
  }

  for (let i = 0; i < totalCombos; i++) {
    const combo = combos.nth(i);
    if (!(await combo.isVisible().catch(() => false))) continue;

    const valorActual = ((await combo.inputValue().catch(() => '')) || '').trim().toLowerCase();
    if (valorActual.includes('falso positivo') || valorActual === '1') {
      console.log(`[Cumplimiento][Bizagi] Acción fila ${i + 1}/${totalCombos} ya estaba en Falso Positivo.`);
      continue;
    }

    const fila = combo.locator('xpath=ancestor::tr[1]');
    let seleccionado = false;

    for (let intento = 1; intento <= 4 && !seleccionado; intento++) {
      await combo.scrollIntoViewIfNeeded().catch(() => {});
      await combo.click({ force: true }).catch(() => {});
      await combo.focus().catch(() => {});

      const btnCombo = fila.locator('.ui-selectmenu-btn, .ui-selectmenu-button, .ui-selectmenu-trigger, div.ui-selectmenu-btn').first();
      if (await btnCombo.isVisible().catch(() => false)) {
        await btnCombo.click({ force: true }).catch(() => {});
      } else {
        await combo.press('ArrowDown').catch(() => {});
      }

      const opcionFalsoPositivo = bizagiPage
        .locator('li, .ui-menuitem, .ui-selectmenu-item, [role="option"]')
        .filter({ hasText: /Falso Positivo/i })
        .first();
      const opcionVisible = await opcionFalsoPositivo
        .waitFor({ state: 'visible', timeout: 2500 })
        .then(() => true)
        .catch(() => false);

      if (!opcionVisible) {
        await combo.press('ArrowDown').catch(() => {});
        await bizagiPage.waitForTimeout(200);
        const opcionVisibleReintento = await opcionFalsoPositivo
          .waitFor({ state: 'visible', timeout: 1500 })
          .then(() => true)
          .catch(() => false);

        if (!opcionVisibleReintento) {
          console.log(`[Cumplimiento][Bizagi] La lista de Acción no abrió en fila ${i + 1}/${totalCombos} (intento ${intento}/4). Reintentando...`);
          await bizagiPage.keyboard.press('Escape').catch(() => {});
          await bizagiPage.waitForTimeout(250);
          continue;
        }
      }

      await opcionFalsoPositivo.click({ force: true }).catch(() => {});
      await bizagiPage.waitForTimeout(300);

      const valorFinal = ((await combo.inputValue().catch(() => '')) || '').trim().toLowerCase();
      seleccionado = valorFinal.includes('falso positivo') || valorFinal === '1';
    }

    if (!seleccionado) {
      throw new Error(`[CRITICO] No se pudo seleccionar 'Falso Positivo' en la fila ${i + 1} de Cumplimiento.`);
    }

    console.log(`[Cumplimiento][Bizagi] Acción fila ${i + 1}/${totalCombos} = Falso Positivo`);
  }
}

async function seleccionarNoSolicitarAclaracionesCumplimiento(bizagiPage: Page) {
  const bloqueAclaraciones = bizagiPage
    .locator('xpath=//*[contains(normalize-space(.),"Solicitar Aclaraciones") or contains(normalize-space(.),"Solicitar Aclaraciones?")]/ancestor::*[self::div or self::section or self::fieldset][1]')
    .first();
  await bloqueAclaraciones.waitFor({ state: 'visible', timeout: 20000 });

  const radioNo = bloqueAclaraciones.locator('input[type="radio"][value="false"]').first();
  const labelNo = bloqueAclaraciones.locator('label').filter({ hasText: /^No$/i }).first();
  const radioNoRole = bloqueAclaraciones.getByRole('radio', { name: /^No$/i }).first();

  if (await radioNo.isVisible().catch(() => false)) {
    await radioNo.scrollIntoViewIfNeeded().catch(() => {});
    await radioNo.check({ force: true }).catch(async () => {
      await radioNo.click({ force: true }).catch(() => {});
    });
  } else if (await radioNoRole.isVisible().catch(() => false)) {
    await radioNoRole.click({ force: true }).catch(() => {});
  } else if (await labelNo.isVisible().catch(() => false)) {
    await labelNo.click({ force: true }).catch(() => {});
  }

  if (!(await radioNo.isChecked().catch(() => false))) {
    const labelNo = bloqueAclaraciones.locator('label:has-text("No")').first();
    await labelNo.click({ force: true }).catch(() => {});
    await bizagiPage.waitForTimeout(200);
  }

  const checked = await radioNo.isChecked().catch(() => false);
  if (!checked) {
    console.log("[Cumplimiento][Bizagi][WARN] radioNo no marcado tras clics. Intentando via evaluate...");
    await radioNo.evaluate((el) => {
        const inp = el as HTMLInputElement;
        inp.checked = true;
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        inp.dispatchEvent(new Event('change', { bubbles: true }));
    }).catch(() => {});
  }

  const finalCheck = await radioNo.isChecked().catch(() => false);
  console.log(`[Cumplimiento][Bizagi] Solicitar Aclaraciones = No (checked=${finalCheck})`);
}

async function clickSiguienteCumplimiento(bizagiPage: Page) {
  const btnSiguiente = bizagiPage.getByRole('button', { name: /Siguiente/i }).first();
  await btnSiguiente.waitFor({ state: 'visible', timeout: 15000 });
  const enabled = await btnSiguiente.isEnabled().catch(() => false);
  if (!enabled) {
    throw new Error("[CRITICO] Botón 'Siguiente' no disponible en Cumplimiento Bizagi.");
  }
  await btnSiguiente.scrollIntoViewIfNeeded().catch(() => {});
  await btnSiguiente.click({ force: true }).catch(() => {});

  const modalConfirmacion = bizagiPage
    .locator('.ui-dialog:visible, [role="dialog"]:visible, .modal:visible')
    .filter({ hasText: /seguro.*continuar|deshacer esta acci/i })
    .first();
  const aparecioConfirmacion = await modalConfirmacion
    .waitFor({ state: 'visible', timeout: 2000 })
    .then(() => true)
    .catch(() => false);
  if (aparecioConfirmacion) {
    const confirmo = await confirmarPopupAceptar(bizagiPage, 5000).catch(() => false);
    if (!confirmo) {
      // Intento final con el selector grabado: getByRole('button', { name: 'Aceptar' })
      const btnAceptarGrabado = bizagiPage.getByRole('button', { name: 'Aceptar' }).first();
      if (await btnAceptarGrabado.isVisible().catch(() => false)) {
          await btnAceptarGrabado.click({ force: true }).catch(() => {});
      } else {
          throw new Error("[CRITICO] No se pudo confirmar con 'Aceptar' la continuacion de Cumplimiento en Bizagi.");
      }
    }
    console.log("[Cumplimiento][Bizagi] Confirmacion aceptada tras Siguiente");
  }

  await bizagiPage.waitForLoadState('domcontentloaded').catch(() => {});
  await bizagiPage.waitForTimeout(1500);
  console.log("[Cumplimiento][Bizagi] Click en Siguiente");
}

async function completarGestionCoincidenciasCumplimientoBizagi(bizagiPage: Page) {
  console.log('[Cumplimiento][Bizagi] completando acciones...');
  await seleccionarAccionesCumplimientoBizagi(bizagiPage);
  console.log('[Cumplimiento][Bizagi] completando solicitar aclaraciones...');
  await seleccionarNoSolicitarAclaracionesCumplimiento(bizagiPage);
  console.log('[Cumplimiento][Bizagi] haciendo click en Siguiente...');
  await clickSiguienteCumplimiento(bizagiPage);
}

async function pantallaCumplimientoAbiertaEstricaBizagi(bizagiPage: Page) {
  const tieneComboAccion = await bizagiPage.locator('input[role="combobox"], input.ui-select-data').first().isVisible().catch(() => false);
  const tieneAclaraciones = await bizagiPage.getByText(/Solicitar Aclaraciones/i).first().isVisible().catch(() => false);
  const tieneSiguiente = await bizagiPage.getByRole('button', { name: /Siguiente/i }).first().isVisible().catch(() => false);
  const tieneListas = await bizagiPage.getByText(/Listas OFAC|Lexis Nexis/i).first().isVisible().catch(() => false);
  const tieneDocs = await bizagiPage.getByText(/Documentos PKM|Coincidencias OFAC|Gesti[oó]n de Coincidencias/i).first().isVisible().catch(() => false);
  const tieneRegresar = await bizagiPage.getByText(/Regresar/i).first().isVisible().catch(() => false);
  return tieneComboAccion || tieneAclaraciones || tieneSiguiente || tieneListas || tieneDocs || tieneRegresar;
}

async function abrirSolicitudCumplimientoBizagiRapido(bizagiPage: Page, mpn: string) {
  const filasInbox = bizagiPage
    .locator('table#ui-bizagi-wp-app-inbox-grid-cases tbody tr[data-idworkflow]:visible');

  const candidatas = [
    {
      origen: 'Gestionar Coincidencias',
      fila: filasInbox.filter({ hasText: new RegExp(mpn, 'i') }).filter({ hasText: /Gestionar Coincidencias/i }).first(),
    },
    {
      origen: 'Verificar PLAFT',
      fila: filasInbox.filter({ hasText: new RegExp(mpn, 'i') }).filter({ hasText: /Verificar PLAFT/i }).first(),
    },
    {
      origen: 'Gestion Debida Diligencia PLAFT',
      fila: filasInbox.filter({ hasText: new RegExp(mpn, 'i') }).filter({ hasText: /Debida Diligencia PLAFT/i }).first(),
    },
    {
      origen: 'PLAFT',
      fila: filasInbox.filter({ hasText: new RegExp(mpn, 'i') }).filter({ hasText: /PLAFT/i }).first(),
    },
  ];

  let filaElegida: Locator | null = null;
  let origenElegido = '';
  for (const candidata of candidatas) {
    const visible = await candidata.fila.isVisible().catch(() => false);
    if (visible) {
      filaElegida = candidata.fila;
      origenElegido = candidata.origen;
      break;
    }
  }

  if (!filaElegida) {
    throw new Error(`[CRITICO] No se encontro en Bizagi una fila util de Cumplimiento para '${mpn}'.`);
  }

  console.log(`[Cumplimiento][Bizagi] intentando abrir fila origen='${origenElegido}' para ${mpn}`);
  await filaElegida.scrollIntoViewIfNeeded().catch(() => {});

  const xpathTextoMpn = `contains(normalize-space(translate(., '\u00A0', ' ')), '${mpn}')`;
  const selectorCeldaNumeroCaso = 'td._I_RadNumber, td._i_RadNumber, td.I_RadNumber, td.i_RadNumber, td[class*="RadNumber"]';
  const selectorSpanNumeroCaso = `${selectorCeldaNumeroCaso} span.ui-bizagi-wp-app-inbox-activity-name`;
  const xpathNumeroCasoExacto = `xpath=//table[@id='ui-bizagi-wp-app-inbox-grid-cases']//tr[@data-idworkflow and .//td[contains(normalize-space(.), '${origenElegido}') or .//*[contains(normalize-space(.), '${origenElegido}')]] and .//td[contains(@class,'RadNumber')]//span[${xpathTextoMpn}]]//td[contains(@class,'RadNumber')]//span[contains(@class,'ui-bizagi-wp-app-inbox-activity-name') and ${xpathTextoMpn}]`;
  const spanNumeroCasoXPath = bizagiPage.locator(xpathNumeroCasoExacto).first();
  const spanNumeroCaso = filaElegida.locator(
    `xpath=.//td[contains(@class,'RadNumber')]//span[contains(@class,'ui-bizagi-wp-app-inbox-activity-name') and ${xpathTextoMpn}]`
  ).first();
  const celdaNumeroCaso = filaElegida.locator(
    `xpath=.//td[contains(@class,'RadNumber') and .//span[${xpathTextoMpn}]]`
  ).first();
  const textoNumeroCaso = filaElegida.getByText(new RegExp(`^${mpn}$`, 'i')).first();
  const linkNumeroCaso = filaElegida.locator(`${selectorSpanNumeroCaso}, ${selectorCeldaNumeroCaso} a, ${selectorCeldaNumeroCaso}`).first();

  const candidatosNumeroCaso = [
    spanNumeroCasoXPath,
    spanNumeroCaso,
    textoNumeroCaso,
    celdaNumeroCaso,
    linkNumeroCaso,
  ];

  const candidatosFallback = [
    filaElegida,
  ];

  const pantallaCumplimientoAbierta = async () => {
    const tieneComboAccion = await bizagiPage.locator('input[role="combobox"], input.ui-select-data').first().isVisible().catch(() => false);
    const tieneAclaraciones = await bizagiPage.getByText(/Solicitar Aclaraciones/i).first().isVisible().catch(() => false);
    const tieneSiguiente = await bizagiPage.getByRole('button', { name: /Siguiente/i }).first().isVisible().catch(() => false);
    const tieneGestionCoincidencias = await bizagiPage.getByText(/Gesti[oó]n de Coincidencias/i).first().isVisible().catch(() => false);
    const tieneOtrasCoincidencias = await bizagiPage.getByText(/Otras Coincidencias/i).first().isVisible().catch(() => false);
    return tieneComboAccion || tieneAclaraciones || tieneSiguiente || tieneGestionCoincidencias || tieneOtrasCoincidencias;
  };

  const objetivoPrioritario = (await spanNumeroCasoXPath.isVisible().catch(() => false))
    ? spanNumeroCasoXPath
    : (await spanNumeroCaso.isVisible().catch(() => false))
    ? spanNumeroCaso
    : (await textoNumeroCaso.isVisible().catch(() => false))
      ? textoNumeroCaso
      : celdaNumeroCaso;
  await objetivoPrioritario.waitFor({ state: 'visible', timeout: 5000 }).catch(() => {});
  console.log(`[Cumplimiento][Bizagi] intento prioritario sobre numero de caso para ${mpn}`);
  for (const accionPrioritaria of [
    async () => { await objetivoPrioritario.click(); },
    async () => { await objetivoPrioritario.dblclick(); },
    async () => {
      const box = await objetivoPrioritario.boundingBox();
      if (!box) throw new Error('sin bounding box');
      await bizagiPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
    },
    async () => {
      const box = await objetivoPrioritario.boundingBox();
      if (!box) throw new Error('sin bounding box');
      await bizagiPage.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
    },
    async () => { await objetivoPrioritario.click({ force: true }); },
    async () => { await objetivoPrioritario.dblclick({ force: true }); },
    async () => {
      await objetivoPrioritario.evaluate((el) => {
        const h = el as HTMLElement;
        h.focus?.();
        h.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        h.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        h.click?.();
        h.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true, view: window }));
      });
    },
  ]) {
    await accionPrioritaria().catch(() => {});
    const inicioPrioritario = Date.now();
    while (Date.now() - inicioPrioritario < 1200) {
      if (await pantallaCumplimientoAbiertaEstricaBizagi(bizagiPage)) return;
      await bizagiPage.waitForTimeout(150);
    }
  }

  for (const candidatoDirecto of [spanNumeroCasoXPath, spanNumeroCaso, textoNumeroCaso, celdaNumeroCaso, linkNumeroCaso]) {
    const visibleDirecto = await candidatoDirecto.isVisible().catch(() => false);
    if (!visibleDirecto) continue;

    await candidatoDirecto.scrollIntoViewIfNeeded().catch(() => {});
    console.log(`[Cumplimiento][Bizagi] click directo en numero de caso para ${mpn}`);

    await candidatoDirecto.click({ force: true }).catch(() => {});
    let inicioDirecto = Date.now();
    while (Date.now() - inicioDirecto < 1500) {
      if (await pantallaCumplimientoAbiertaEstricaBizagi(bizagiPage)) return;
      await bizagiPage.waitForTimeout(200);
    }

    await candidatoDirecto.dblclick({ force: true }).catch(() => {});
    inicioDirecto = Date.now();
    while (Date.now() - inicioDirecto < 1500) {
      if (await pantallaCumplimientoAbiertaEstricaBizagi(bizagiPage)) return;
      await bizagiPage.waitForTimeout(200);
    }

    await candidatoDirecto.evaluate((el) => {
      const h = el as HTMLElement;
      h.focus?.();
      h.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true, view: window }));
      h.click?.();
    }).catch(() => {});
    inicioDirecto = Date.now();
    while (Date.now() - inicioDirecto < 1200) {
      if (await pantallaCumplimientoAbiertaEstricaBizagi(bizagiPage)) return;
      await bizagiPage.waitForTimeout(200);
    }
  }

  let indiceCandidato = 0;
  for (const candidato of [...candidatosNumeroCaso, ...candidatosFallback]) {
    indiceCandidato++;
    const visible = await candidato.isVisible().catch(() => false);
    if (!visible) continue;

    await candidato.scrollIntoViewIfNeeded().catch(() => {});
    const grupo = indiceCandidato <= candidatosNumeroCaso.length ? 'numero-caso' : 'fallback';
    console.log(`[Cumplimiento][Bizagi] probando candidato ${indiceCandidato}/${candidatosNumeroCaso.length + candidatosFallback.length} grupo=${grupo} para abrir ${origenElegido}`);

    const acciones: Array<() => Promise<void>> = [
      async () => { await candidato.click({ force: true }); },
      async () => { await candidato.dblclick({ force: true }); },
      async () => {
        const box = await candidato.boundingBox();
        if (!box) throw new Error('sin bounding box');
        await bizagiPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
      },
      async () => {
        await candidato.evaluate((el) => {
          const h = el as HTMLElement;
          h.focus?.();
          h.click?.();
          h.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        });
      },
      async () => {
        await candidato.focus().catch(() => {});
        await candidato.press('Enter').catch(() => {});
      },
    ];

    for (const accion of acciones) {
      await accion().catch(() => {});
      const inicio = Date.now();
      while (Date.now() - inicio < (grupo === 'numero-caso' ? 2200 : 1400)) {
        if (await pantallaCumplimientoAbiertaEstricaBizagi(bizagiPage)) return;
        await bizagiPage.waitForTimeout(250);
      }
    }
  }

  throw new Error(`[CRITICO] No se pudo abrir la solicitud de Cumplimiento para '${mpn}' desde la fila encontrada.`);
}

function extraerMpnDesdePortal(portalPage: Page) {
  return portalPage.locator('span.p-tag-value').filter({ hasText: /^MPN-\d+$/i }).first();
}

async function resolverMpnDesdePortal(portalPage: Page) {
  const tagMpn = extraerMpnDesdePortal(portalPage);
  const desdeTag = (await tagMpn.textContent().catch(() => '') || '').trim().toUpperCase();
  if (/^MPN-\d+$/i.test(desdeTag)) {
    console.log(`[Cumplimiento] MPN detectado desde tag visible: ${desdeTag}`);
    return desdeTag;
  }

  const bodyText = (await portalPage.locator('body').innerText().catch(() => '')) || '';
  const matchBody = bodyText.match(/MPN-\d+/i);
  if (matchBody?.[0]) {
    const mpnBody = matchBody[0].toUpperCase();
    console.log(`[Cumplimiento] MPN detectado desde bodyText: ${mpnBody}`);
    return mpnBody;
  }

  const html = (await portalPage.content().catch(() => '')) || '';
  const matchHtml = html.match(/MPN-\d+/i);
  if (matchHtml?.[0]) {
    const mpnHtml = matchHtml[0].toUpperCase();
    console.log(`[Cumplimiento] MPN detectado desde HTML: ${mpnHtml}`);
    return mpnHtml;
  }

  return '';
}

export async function abrirSolicitudCumplimientoEnBizagiDesdePortal(
  portalPage: Page,
  opts: { url?: string; usuario?: string; password?: string } = {}
) {
  const badgeCumplimiento = portalPage.getByText(/Verificaciones\s*-\s*Cumplimiento/i).first();
  const visibleCumplimiento = await badgeCumplimiento.isVisible().catch(() => false);
  if (!visibleCumplimiento) return null;

  const mpn = await resolverMpnDesdePortal(portalPage);

  if (!mpn) {
    console.log(`[Cumplimiento] Visible, pero no se pudo capturar el numero MPN. url=${portalPage.url()}`);
    return null;
  }

  const bizagiUrl = resolverBizagiUrl(opts.url);
  const bizagiUsuario = opts.usuario ?? process.env.BIZAGI_USER ?? 'domain\\admon';
  const bizagiPassword = opts.password ?? process.env.BIZAGI_PASS;

  const bizagiPage = await portalPage.context().newPage();
  await bizagiPage.goto(bizagiUrl, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await entrarABizagiSiHayLoginRobusto(bizagiPage, bizagiUsuario, bizagiPassword);
  // Priorizar el flujo de Admin si el buscador global suele fallar
  const casoEncontrado = await buscarCasoPorMpnEnBizagi(bizagiPage, mpn, { useAdminFallback: true, openIfFound: true });
  if (!casoEncontrado) {
    throw new Error(`[CRITICO] No se encontro ninguna fila en Bizagi para '${mpn}'.`);
  }
  await abrirSolicitudCumplimientoBizagiRapido(bizagiPage, mpn);
  await completarGestionCoincidenciasCumplimientoBizagi(bizagiPage);

  await portalPage.bringToFront().catch(() => {});
  return { mpn, bizagiPage };
}
