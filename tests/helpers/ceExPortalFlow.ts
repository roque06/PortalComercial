import path from 'node:path';
import fs from 'node:fs';
import {
  clickReintentarListaSiVisible,
  seleccionarDropdownFiltrableConReintentar,
} from './uiHelpers';

function escapeRegex(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export async function seleccionarInstrumentoRobusto(page: import("@playwright/test").Page) {
  // Localiza el dropdown de Instrumento usando multiples estrategias de localizacion.
  const encontrarDropdownInstrumento = async () => {
    // Estrategia 1: label con texto exacto "Instrumento"
    const labelExacto = page.locator('label').filter({ hasText: /^Instrumento$/i }).first();
    if (await labelExacto.isVisible().catch(() => false)) {
      const dd = labelExacto.locator('xpath=following::div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"][1]').first();
      if (await dd.isVisible().catch(() => false)) return dd;
      const ddAnc = labelExacto.locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]').first();
      if (await ddAnc.isVisible().catch(() => false)) return ddAnc;
    }
    // Estrategia 2: getByText con regex (incluye elementos no-label)
    const labelRegex = page.getByText(/^Instrumento$/i).first();
    if (await labelRegex.isVisible().catch(() => false)) {
      const dd = labelRegex.locator('xpath=following::div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"][1]').first();
      if (await dd.isVisible().catch(() => false)) return dd;
    }
    // Estrategia 3: cualquier dropdown cuyo label placeholder diga "Instrumento"
    const allDropdowns = page.locator('div.p-dropdown:visible, [data-pc-name="dropdown"]:visible');
    const count = await allDropdowns.count().catch(() => 0);
    for (let k = 0; k < count; k++) {
      const dd = allDropdowns.nth(k);
      const txt = ((await dd.locator('.p-dropdown-label, [data-pc-section="label"]').first().textContent().catch(() => '')) || '').trim();
      if (/instrumento/i.test(txt) || txt === '') {
        // Check surrounding text for "Instrumento"
        const parent = dd.locator('xpath=ancestor::*[self::div or self::fieldset][1]').first();
        const parentTxt = ((await parent.textContent().catch(() => '')) || '').toLowerCase();
        if (/instrumento/i.test(parentTxt)) return dd;
      }
    }
    return null;
  };

  // Verifica si el dropdown ya tiene un valor seleccionado (no placeholder).
  const tieneValor = async (dd: import("@playwright/test").Locator) => {
    const labelEl = dd.locator('.p-dropdown-label, [data-pc-section="label"]').first();
    const texto = ((await labelEl.textContent().catch(() => '')) || '').trim();
    const clases = (await labelEl.getAttribute('class').catch(() => '')) || '';
    return !!texto && !/^seleccione|^--/i.test(texto) && !clases.includes('p-placeholder');
  };

  // Abre el panel y selecciona el primer item con teclado (más fiable en Angular/PrimeNG).
  const seleccionarPrimerItem = async (dd: import("@playwright/test").Locator) => {
    await dd.scrollIntoViewIfNeeded().catch(() => {});
    await dd.click({ force: true }).catch(() => {});
    await page.waitForTimeout(400);

    // Obtener panel por aria-controls o fallback
    const combobox = dd.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute('aria-controls').catch(() => null);
    let panel: import("@playwright/test").Locator | null = null;
    if (panelId) {
      const byId = page.locator(`#${panelId}`);
      if (await byId.waitFor({ state: 'visible', timeout: 5000 }).then(() => true).catch(() => false)) panel = byId;
    }
    if (!panel) {
      const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
      if (await fallback.waitFor({ state: 'visible', timeout: 5000 }).then(() => true).catch(() => false)) panel = fallback;
    }
    if (!panel) {
      console.log('[Instrumento] Panel no visible tras click.');
      return false;
    }

    const items = panel.locator("li[role='option'], .p-dropdown-item, [data-pc-section='item']");
    await items.first().waitFor({ state: 'visible', timeout: 6000 }).catch(() => {});
    const count = await items.count().catch(() => 0);
    console.log(`[Instrumento] Items en panel: ${count}`);
    if (count === 0) {
      await page.keyboard.press('Escape').catch(() => {});
      return false;
    }

    // Estrategia 1: teclado — más fiable en Angular/PrimeNG (no depende de coordenadas)
    await page.keyboard.press('ArrowDown').catch(() => {});
    await page.waitForTimeout(150);
    await page.keyboard.press('Enter').catch(() => {});
    await page.waitForTimeout(300);

    // Si el teclado no funcionó, intentar click directo en el item
    const ddTrasKey = await tieneValor(dd).catch(() => false);
    if (!ddTrasKey) {
      console.log('[Instrumento] Teclado no seleccionó. Intentando click directo en item...');
      for (let k = 0; k < Math.min(count, 5); k++) {
        const item = items.nth(k);
        const txt = ((await item.textContent().catch(() => '')) || '').trim();
        if (txt && !/^seleccione|^--/i.test(txt)) {
          await item.scrollIntoViewIfNeeded().catch(() => {});
          await item.click({ force: true }).catch(() => {});
          await page.waitForTimeout(200);
          // Verificar que el panel cerró (señal de que el click fue aceptado)
          const panelCerro = !(await panel.isVisible().catch(() => false));
          if (panelCerro) return true;
          // Fallback: click via JS para bypassear posibles overlays
          const handle = await item.elementHandle().catch(() => null);
          if (handle) {
            await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => {});
            await page.waitForTimeout(200);
          }
          return true;
        }
      }
      // Último recurso: click en primer item sin filtrar
      const primerItem = items.first();
      await primerItem.click({ force: true }).catch(() => {});
    }

    return true;
  };

  for (let i = 1; i <= 5; i++) {
    await clickReintentarListaSiVisible(page, "Instrumento");

    const dd = await encontrarDropdownInstrumento();
    if (!dd) {
      console.log(`[Instrumento] Intento ${i}/5: dropdown no encontrado. Esperando...`);
      await page.waitForTimeout(1000);
      continue;
    }

    // Si ya tiene valor, no hacer nada
    if (await tieneValor(dd)) {
      console.log(`[Instrumento] Intento ${i}/5: ya tiene valor seleccionado.`);
      return;
    }

    // Intentar con la función existente primero
    try {
      await seleccionarDropdownFiltrableConReintentar(
        page,
        "Instrumento",
        { index: 1 },
        { maxIntentos: 3, esperaMs: 800, timeoutCampoMs: 8000, timeoutPanelMs: 6000, usarFiltro: false }
      );
    } catch (_) {
      // Si falla, usar estrategia directa
      console.log(`[Instrumento] Intento ${i}/5: seleccionarDropdownFiltrableConReintentar fallo. Usando estrategia directa...`);
      await seleccionarPrimerItem(dd).catch(() => {});
    }

    await page.waitForTimeout(400);
    if (await tieneValor(dd)) {
      console.log(`[Instrumento] Intento ${i}/5: valor seleccionado correctamente.`);
      return;
    }

    // Doble intento con estrategia directa
    console.log(`[Instrumento] Intento ${i}/5: sin valor tras primer intento. Reintentando con estrategia directa...`);
    await seleccionarPrimerItem(dd).catch(() => {});
    await page.waitForTimeout(400);
    if (await tieneValor(dd)) return;

    await page.waitForTimeout(800);
  }

  throw new Error("No se pudo seleccionar Instrumento tras varios intentos");
}

export async function seleccionarDropdownPorCampo(
  page: import("@playwright/test").Page,
  labelText: string | RegExp,
  index = 0
) {
  const pattern = labelText instanceof RegExp
    ? labelText
    : new RegExp(escapeRegex(String(labelText)).replace(/\\\?/g, "."), "i");

  const timeoutTotalMs = 30000;
  const inicio = Date.now();
  let ultimoError = "";

  while (Date.now() - inicio < timeoutTotalMs) {
    try {
      let label = page.locator("label").filter({ hasText: pattern }).first();
      const labelVisible = await label.isVisible().catch(() => false);
      if (!labelVisible) {
        label = page.getByText(pattern).first();
      }

      let field = label
        .locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]')
        .first();
      let fieldVisible = await field.isVisible().catch(() => false);
      if (!fieldVisible) {
        field = label
          .locator('xpath=following::*[self::div[contains(@class,"p-dropdown")] or @data-pc-name="dropdown"][1]')
          .first();
        fieldVisible = await field.isVisible().catch(() => false);
      }
      if (!fieldVisible) {
        await page.waitForTimeout(150);
        continue;
      }

      const dropdown = field.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();
      await dropdown.waitFor({ state: "visible", timeout: 2500 });
      await dropdown.click({ force: true });

      const combobox = dropdown.locator('[role="combobox"]').first();
      const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
      let panel: import("@playwright/test").Locator | null = null;
      if (panelId) {
        const byId = page.locator(`#${panelId}`);
        const byIdVisible = await byId.waitFor({ state: "visible", timeout: 1800 }).then(() => true).catch(() => false);
        if (byIdVisible) panel = byId;
      }
      if (!panel) {
        const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').last();
        const fallbackVisible = await fallback.waitFor({ state: "visible", timeout: 1800 }).then(() => true).catch(() => false);
        if (fallbackVisible) panel = fallback;
      }
      if (!panel) {
        await page.keyboard.press("Escape").catch(() => { });
        await page.waitForTimeout(120);
        continue;
      }

      const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
      const hasItems = await items.first().waitFor({ state: "visible", timeout: 2500 }).then(() => true).catch(() => false);
      if (!hasItems) {
        await page.mouse.click(10, 10).catch(() => { });
        await page.waitForTimeout(120);
        continue;
      }

      const count = await items.count().catch(() => 0);
      if (count > index) {
        await items.nth(index).click({ force: true });
        return;
      }

      await page.mouse.click(10, 10).catch(() => { });
      await page.waitForTimeout(120);
    } catch (e) {
      ultimoError = String(e);
      await page.waitForTimeout(150);
    }
  }

  throw new Error(`No se pudo seleccionar dropdown '${String(labelText)}' index ${index}. ${ultimoError}`);
}

export async function esperarYClickReintentarPaisIdentificacion(
  page: import("@playwright/test").Page,
  timeoutMs = 30000
) {
  const labels = page
    .locator("label")
    .filter({ hasText: /(Pa[ií]s|Tipo) de identificaci[oó]n/i });

  const total = await labels.count().catch(() => 0);
  if (!total) return;

  for (let j = 0; j < total; j++) {
    const slot = labels.nth(j).locator("xpath=following-sibling::*[1]").first();
    const btnRetry = slot
      .locator(
        'div.p-inputgroup:has-text("Reintentar buscar lista") button,' +
        'button:has-text("Reintentar buscar lista"),' +
        'button:has-text("Reintentar"),' +
        'button[aria-label*="Reintentar"]'
      )
      .first();
    const dropdown = slot.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();

    await Promise.race([
      btnRetry.waitFor({ state: "visible", timeout: timeoutMs }).catch(() => { }),
      dropdown.waitFor({ state: "visible", timeout: timeoutMs }).catch(() => { }),
    ]);

    for (let i = 1; i <= 6; i++) {
      const visibleRetry = await btnRetry.isVisible().catch(() => false);
      if (!visibleRetry) break;
      await btnRetry.click({ force: true });
      await page.waitForTimeout(1200);
    }
  }
}

export async function validarApnfdYSeleccionarNoSiVacio(page: import("@playwright/test").Page) {
  // Busca el bloque APNFD por titulo/texto caracteristico.
  const apnfdSection = page.locator(
    'xpath=//*[contains(normalize-space(.),"APNFD") or contains(normalize-space(.),"Requiere APNFD")]/ancestor::fieldset[1] | //fieldset[.//*[contains(normalize-space(.),"alguna de las actividades siguientes")]]'
  ).first();

  const visible = await apnfdSection.isVisible().catch(() => false);
  if (!visible) return false;

  // Cada pregunta suele estar en un grupo tipo selectbutton/radiogroup
  const grupos = apnfdSection.locator(
    '.p-selectbutton, [data-pc-name="selectbutton"], [role="radiogroup"]'
  );

  const total = await grupos.count().catch(() => 0);
  if (!total) return false;

  let cambios = 0;

  for (let i = 0; i < total; i++) {
    const grupo = grupos.nth(i);

    const yaSeleccionado = grupo.locator(
      '[role="radio"][aria-checked="true"], [aria-pressed="true"], .p-highlight'
    );
    const tieneSeleccion = (await yaSeleccionado.count().catch(() => 0)) > 0;
    if (tieneSeleccion) continue;

    const opcionNo = grupo.locator(
      '[role="radio"][aria-label="No"], button:has-text("No"), [data-pc-section="button"]:has-text("No")'
    ).first();

    const noVisible = await opcionNo.isVisible().catch(() => false);
    if (!noVisible) continue;

    await opcionNo.scrollIntoViewIfNeeded().catch(() => { });
    await opcionNo.click({ force: true });
    cambios++;
    await page.waitForTimeout(120);
  }

  return cambios > 0;
}

export async function validarGestionDocumentalSiRequerido(page: import("@playwright/test").Page) {
  const evidenciaPath = path.resolve(process.cwd(), 'data', 'Documentos', 'evidencia.pdf');
  const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
  const visibleBtn = await btnGestion.isVisible().catch(() => false);
  if (!visibleBtn) return false;

  const badge = btnGestion.locator('[data-pc-name="badge"], .p-badge, span[class*="badge"]').first();
  const badgeText = ((await badge.textContent().catch(() => "")) || "").trim();
  const ariaLabel = ((await btnGestion.getAttribute('aria-label').catch(() => "")) || "").trim();
  const textoBoton = ((await btnGestion.textContent().catch(() => "")) || "").trim();
  const sourceText = [badgeText, ariaLabel, textoBoton].join(" ");
  const matchPendientes = sourceText.match(/\b(\d+)\b/);
  const pendientes = matchPendientes ? Number.parseInt(matchPendientes[1], 10) : 0;
  const requiereDocs = Number.isFinite(pendientes) && pendientes > 0;
  if (!requiereDocs) return false;

  await btnGestion.scrollIntoViewIfNeeded().catch(() => { });
  await btnGestion.click().catch(() => { });
  await cargarDocumentoEnGestionDocumental(page, evidenciaPath).catch((e) => {
    console.log(`[GestionDoc][ce-ex] Carga documental no completada: ${String(e)}`);
  });

  return true;
}

export async function cargarDocumentoEnGestionDocumental(page: import("@playwright/test").Page, filePath: string) {
  console.log(`[GestionDoc][ce-ex] Iniciando carga: ${filePath}`);
  if (!fs.existsSync(filePath)) {
    throw new Error(`[GestionDoc][ce-ex] No existe archivo para adjuntar: ${filePath}`);
  }

  const modalDoc = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
    .first();
  await modalDoc.waitFor({ state: 'visible', timeout: 12000 });

  const uploadSelector =
    'button:has(i.pi-upload), button:has(.pi-upload), button:has([class*="pi-upload"]), button.p-button-icon-only.p-button-sm';
  const leerBadgePendientes = async () => {
    const badge = modalDoc.locator('[data-pc-name="badge"], .p-badge, span[class*="badge"]').first();
    const badgeText = ((await badge.textContent().catch(() => "")) || "").trim();
    const match = badgeText.match(/\b(\d+)\b/);
    return match ? Number.parseInt(match[1], 10) : 0;
  };
  const contarUploadsVisibles = async () => {
    const botones = modalDoc.locator(uploadSelector);
    const total = await botones.count().catch(() => 0);
    let visibles = 0;
    for (let i = 0; i < total; i++) {
      if (await botones.nth(i).isVisible().catch(() => false)) visibles++;
    }
    return visibles;
  };
  const esperarBotonUploadDisponible = async (timeoutMs: number) => {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const botones = modalDoc.locator(uploadSelector);
      const total = await botones.count().catch(() => 0);
      for (let i = 0; i < total; i++) {
        const boton = botones.nth(i);
        const visible = await boton.isVisible().catch(() => false);
        if (!visible) continue;
        const enabled = await boton.isEnabled().catch(() => false);
        if (enabled) return boton;
      }
      await page.waitForTimeout(350);
    }
    return null;
  };

  const confirmarPopupCargaSiVisible = async () => {
    const modalCarga = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Cargar documento/i })
      .last();
    const visibleCarga = await modalCarga.isVisible().catch(() => false);
    if (!visibleCarga) return;

    const btnCargar = modalCarga.getByRole('button', { name: /^Cargar$/i }).first();
    if (await btnCargar.isVisible().catch(() => false)) {
      await btnCargar.click({ force: true }).catch(() => { });
    }
    await modalCarga.waitFor({ state: 'hidden', timeout: 10000 }).catch(() => { });
  };

  const cerrarModalGestionDocumental = async () => {
    const btnCerrar = modalDoc
      .locator('button[aria-label="Cerrar"], button.p-dialog-header-close, [data-pc-section="closebutton"]')
      .first();
    const btnCancelar = modalDoc.getByRole('button', { name: /Cancelar|Cerrar/i }).first();

    if (await btnCerrar.isVisible().catch(() => false)) {
      await btnCerrar.click({ force: true }).catch(() => { });
    } else if (await btnCancelar.isVisible().catch(() => false)) {
      await btnCancelar.click({ force: true }).catch(() => { });
    } else {
      await page.keyboard.press('Escape').catch(() => { });
    }

    await modalDoc.waitFor({ state: 'hidden', timeout: 8000 }).catch(() => { });
  };

  const deadlineCargaTotal = Date.now() + 45000;
  while (Date.now() < deadlineCargaTotal) {
    const badgePendientes = await leerBadgePendientes();
    const uploadsVisibles = await contarUploadsVisibles();

    if (badgePendientes <= 0 && uploadsVisibles <= 0) break;

    const botonUpload = await esperarBotonUploadDisponible(uploadsVisibles > 0 ? 4000 : 15000);
    if (!botonUpload) {
      const badgeRevisado = await leerBadgePendientes();
      const uploadsRevisados = await contarUploadsVisibles();
      if (badgeRevisado <= 0 && uploadsRevisados <= 0) break;
      throw new Error(
        `[GestionDoc][ce-ex] Aun hay documentos pendientes pero no aparecio un boton de carga. badge=${badgeRevisado} uploads=${uploadsRevisados}`
      );
    }

    const chooserPromise = page.waitForEvent('filechooser', { timeout: 3500 }).catch(() => null);
    await botonUpload.click({ force: true }).catch(() => { });
    const chooser = await chooserPromise;
    if (chooser) {
      await chooser.setFiles(filePath);
      await confirmarPopupCargaSiVisible();
    } else {
      const inputFile = modalDoc.locator('input[type="file"]').first();
      if (await inputFile.isVisible().catch(() => false)) {
        await inputFile.setInputFiles(filePath);
        await confirmarPopupCargaSiVisible();
      } else {
        const modalCarga = page
          .locator('.p-dialog:visible, [role="dialog"]:visible')
          .filter({ hasText: /Cargar documento/i })
          .last();
        const visibleCarga = await modalCarga.isVisible().catch(() => false);
        if (visibleCarga) {
          const inputCarga = modalCarga.locator('input[type="file"]').first();
          if (await inputCarga.isVisible().catch(() => false)) {
            await inputCarga.setInputFiles(filePath);
          }
          await confirmarPopupCargaSiVisible();
        }
      }
    }

    await page.waitForTimeout(900);
  }

  await cerrarModalGestionDocumental();
}

export async function abrirBpmSiVerificacionConoceCliente(page: import("@playwright/test").Page) {
  const badgeVerificacion = page.getByText(/Verificaciones\s*-\s*Conoce tu cliente/i).first();
  const visibleVerificacion = await badgeVerificacion.isVisible().catch(() => false);
  if (!visibleVerificacion) return null;

  const tagMpn = page.locator('span.p-tag-value').filter({ hasText: /^MPN-\d+$/i }).first();
  let mpn = (await tagMpn.textContent().catch(() => "") || "").trim();
  if (!mpn) {
    const bodyText = (await page.locator("body").innerText().catch(() => "")) || "";
    const match = bodyText.match(/MPN-\d+/i);
    mpn = match ? match[0].toUpperCase() : "";
  }

  if (!mpn) {
    console.log("[Verificacion] Visible, pero no se pudo capturar el numero MPN.");
    return null;
  }

  console.log(`[Verificacion] MPN capturado: ${mpn}`);

  const bpmPage = await page.context().newPage();
  try {
    await bpmPage.goto("https://test-bscrd-santacruz.bizagi.com/", { waitUntil: "domcontentloaded", timeout: 60000 });
    console.log("[Verificacion] BPM abierto en nueva pestana.");
  } catch (e) {
    console.log(`[Verificacion] No se pudo abrir BPM: ${String(e)}`);
  }

  await page.bringToFront().catch(() => { });
  return { mpn, bpmPage };
}

export async function esperarPortalListoTrasLogin(
  page: import("@playwright/test").Page,
  opts: { timeoutMs?: number; stepMs?: number } = {}
) {
  const timeoutMs = opts.timeoutMs ?? 240000;
  const stepMs = opts.stepMs ?? 1500;
  const inicio = Date.now();
  let loginDetectado = false;
  let intentoPassword = false;

  while (Date.now() - inicio < timeoutMs) {
    const urlActual = page.url();
    const enLoginMicrosoft = /login\.microsoftonline\.com/i.test(urlActual);

    if (enLoginMicrosoft) {
      if (!loginDetectado) {
        console.log("Login Microsoft detectado. Esperando autenticacion manual...");
        loginDetectado = true;
      }

      // Si aparece, fuerza flujo por password para evitar push loop.
      if (!intentoPassword) {
        const usarPassword = page
          .locator('a:has-text("Use your password instead"), a:has-text("Usar su contrase"), a:has-text("Usa tu contrase")')
          .first();
        if (await usarPassword.isVisible().catch(() => false)) {
          await usarPassword.click({ force: true }).catch(() => { });
          intentoPassword = true;
        }
      }

      await page.waitForTimeout(stepMs);
      continue;
    }

    const inputNumero = page
      .locator("label")
      .filter({ hasText: /N[uú]mero de identificaci[oó]n/i })
      .first()
      .locator('xpath=following::input[1]');

    const inputNumeroVisible = await inputNumero.isVisible().catch(() => false);
    if (inputNumeroVisible || /portalcomercial\.contigo\.do/i.test(urlActual)) {
      return { loginDetectado, urlActual };
    }

    await page.waitForTimeout(stepMs);
  }

  throw new Error(`No se pudo abrir el portal luego del login. URL actual: ${page.url()}`);
}
