import type { RegistroExcel } from './ceExExcel';

export async function agregarRelacionadoSiAplica(
  page: import("@playwright/test").Page,
  registro: RegistroExcel
) {
  if (!registro.relacionado) return false;
  if (!registro.cedRelacionado) {
    throw new Error(`[CRITICO] Relacionado=Si pero Ced_Relacionado vacio para '${registro.identificacion}'.`);
  }

  const seccionRelacionados = page
    .locator(
      'xpath=//*[self::fieldset or self::div][.//*[contains(translate(normalize-space(.),"ÁÉÍÓÚáéíóú","AEIOUaeiou"),"Relacionados")]][1]'
    )
    .first();
  if (await seccionRelacionados.isVisible().catch(() => false)) {
    await seccionRelacionados.scrollIntoViewIfNeeded().catch(() => { });
  }

  const btnRelacionadosCandidatos = [
    page.getByRole('button', { name: /A[ñn]adir relacionado/i }).first(),
    page.locator('button:has-text("Añadir relacionado"), button:has-text("Anadir relacionado")').first(),
    page.locator('xpath=//button[contains(normalize-space(.),"Añadir relacionado") or contains(normalize-space(.),"Anadir relacionado")]').first(),
  ];

  let clicAnadir = false;
  for (let intento = 1; intento <= 6 && !clicAnadir; intento++) {
    for (const btn of btnRelacionadosCandidatos) {
      const visible = await btn.isVisible().catch(() => false);
      if (!visible) continue;
      await btn.scrollIntoViewIfNeeded().catch(() => { });
      await btn.click({ force: true }).catch(() => { });
      clicAnadir = true;
      break;
    }
    if (clicAnadir) break;
    await page.mouse.wheel(0, 500).catch(() => { });
    await page.waitForTimeout(500);
  }

  if (!clicAnadir) {
    throw new Error(`[CRITICO] No se encontro el boton 'Añadir relacionado'. URL: ${page.url()}`);
  }

  const modalRelacionado = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /A[ñn]adir relacionado/i })
    .first();
  await modalRelacionado.waitFor({ state: 'visible', timeout: 15000 });

  const labelNumero = modalRelacionado
    .locator("label")
    .filter({ hasText: /N[uú]mero de identificaci[oó]n/i })
    .first();
  const inputNumero = labelNumero.locator('xpath=following::input[1]').first();
  await inputNumero.waitFor({ state: 'visible', timeout: 20000 });

  // Mismas validaciones base de identificación: si está deshabilitado, reintenta lista.
  for (let i = 1; i <= 5; i++) {
    const enabled = await inputNumero.isEnabled().catch(() => false);
    if (enabled) break;

    const btnRetry = modalRelacionado
      .locator(
        'button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button[aria-label*="Reintentar"]'
      )
      .first();
    if (await btnRetry.isVisible().catch(() => false)) {
      await btnRetry.click({ force: true });
      await page.waitForTimeout(1000);
    } else {
      await page.waitForTimeout(700);
    }
  }

  const enabledFinal = await inputNumero.isEnabled().catch(() => false);
  if (!enabledFinal) {
    throw new Error("[CRITICO] 'Numero de identificacion' de relacionado sigue deshabilitado.");
  }

  await inputNumero.click({ force: true });
  await inputNumero.fill(registro.cedRelacionado);

  const btnBuscar = labelNumero.locator('xpath=following::button[1]').first();
  if (await btnBuscar.isVisible().catch(() => false)) {
    await btnBuscar.click({ force: true });
  } else {
    await page.keyboard.press('Enter').catch(() => { });
  }

  const spinnerConsulta = page.getByText(/Consultando datos del solicitante/i).first();
  await spinnerConsulta.waitFor({ state: 'visible', timeout: 6000 }).catch(() => { });
  await spinnerConsulta.waitFor({ state: 'hidden', timeout: 120000 }).catch(() => { });
  await page.waitForTimeout(500);

  const btnAceptar = modalRelacionado.getByRole('button', { name: /^Aceptar$/i }).first();
  await btnAceptar.waitFor({ state: 'visible', timeout: 10000 });
  await btnAceptar.click();
  await modalRelacionado.waitFor({ state: 'hidden', timeout: 15000 });

  const seleccionarDropdownIndexEnModal = async (
    modal: import("@playwright/test").Locator,
    labelRegex: RegExp,
    index: number,
    timeoutMs = 20000
  ) => {
    let anchor = modal.locator("label").filter({ hasText: labelRegex }).first();
    const hayLabelVisible = await anchor.isVisible().catch(() => false);
    if (!hayLabelVisible) {
      anchor = modal.getByText(labelRegex).first();
    }
    await anchor.waitFor({ state: "visible", timeout: timeoutMs });

    // Importante: usar el dropdown "inmediato" del label para evitar tomar el campo vecino.
    let dropdown = anchor
      .locator(
        'xpath=following::div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"][1]'
      )
      .first();
    if (!(await dropdown.isVisible().catch(() => false))) {
      const field = anchor
        .locator('xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]')
        .first();
      await field.waitFor({ state: "visible", timeout: timeoutMs });
      dropdown = field.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();
    }
    await dropdown.waitFor({ state: "visible", timeout: timeoutMs });
    await dropdown.scrollIntoViewIfNeeded().catch(() => { });
    await dropdown.click({ force: true });

    const combobox = dropdown.locator('[role="combobox"]').first();
    const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
    const panel = panelId
      ? page.locator(`#${panelId}`)
      : page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').first();
    await panel.waitFor({ state: "visible", timeout: 10000 });

    const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
    await items.first().waitFor({ state: "visible", timeout: 10000 });
    const count = await items.count();
    if (count <= index || count === 0) {
      throw new Error(`No hay item index ${index} para '${labelRegex}'.`);
    }

    const target = items.nth(index);
    await target.scrollIntoViewIfNeeded().catch(() => { });
    await target.click({ force: true });

    // Verifica que el valor quedo seleccionado (no placeholder vacio).
    const valorSeleccionado = dropdown.locator(".p-dropdown-label, [data-pc-section='label']").first();
    const texto = ((await valorSeleccionado.textContent().catch(() => "")) || "").trim();
    const clases = ((await valorSeleccionado.getAttribute("class").catch(() => "")) || "");
    const sinSeleccion = !texto || /^seleccione/i.test(texto) || clases.includes("p-placeholder");
    if (sinSeleccion) {
      throw new Error(`No se aplico seleccion para '${labelRegex}'.`);
    }
  };

  const clickReintentarPorLabelEnModal = async (modal: import("@playwright/test").Locator, labelRegex: RegExp) => {
    let label = modal.locator("label").filter({ hasText: labelRegex }).first();
    if ((await label.count().catch(() => 0)) === 0) {
      label = modal.getByText(labelRegex).first();
    }

    const candidatos = [
      label
        .locator(
          'xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-inputgroup")]][1]//div[contains(@class,"p-inputgroup")]//button[1]'
        )
        .first(),
      label
        .locator(
          'xpath=ancestor::*[self::div or self::span][.//button[contains(@class,"p-button-warning")]][1]//button[contains(@class,"p-button-warning")][1]'
        )
        .first(),
      modal
        .locator(
          'div.p-inputgroup:has(input[value*="Reintentar"]), div.p-inputgroup:has(input[placeholder*="Reintentar"]), div.p-inputgroup:has-text("Reintentar buscar lista")'
        )
        .first()
        .locator('button')
        .first(),
      modal.locator('button.p-button-warning, button:has(i.pi-refresh), button:has(i.pi-replay)').first(),
    ];

    for (const btnRetry of candidatos) {
      const visible = await btnRetry.isVisible().catch(() => false);
      if (!visible) continue;
      await btnRetry.scrollIntoViewIfNeeded().catch(() => { });
      await btnRetry.click({ force: true }).catch(() => { });
      await page.waitForTimeout(1200);
      return true;
    }

    return false;
  };

  // Asociar productos al relacionado recien agregado.
  const btnAsociarProducto = page
    .locator(
      'button:has-text("Asociar a productos"), button:has-text("Asociar a producto"), button:has-text("Asociar Producto"), button:has-text("Asociar productos")'
    )
    .first();
  await btnAsociarProducto.waitFor({ state: "visible", timeout: 20000 });
  await btnAsociarProducto.scrollIntoViewIfNeeded().catch(() => { });
  await btnAsociarProducto.click({ force: true });

  const modalAsociar = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /A[ñn]adir relacionado/i })
    .first();
  await modalAsociar.waitFor({ state: "visible", timeout: 15000 });

  const btnAnadir = modalAsociar.getByRole("button", { name: /^A[ñn]adir$/i }).first();
  await btnAnadir.waitFor({ state: "visible", timeout: 15000 });
  const rxTipoRelacionado = /^Tipo de relacionado$/i;
  const rxTipoRelacion = /^Tipo de relaci(?:ó|o)n$/i;

  const camposRelacionVisible = async () => {
    const tipoRelacionadoVisible = await modalAsociar
      .locator("label")
      .filter({ hasText: rxTipoRelacionado })
      .first()
      .isVisible()
      .catch(() => false);
    const tipoRelacionadoTextVisible = await modalAsociar.getByText(rxTipoRelacionado).first().isVisible().catch(() => false);
    const tipoRelacionVisible = await modalAsociar
      .locator("label")
      .filter({ hasText: rxTipoRelacion })
      .first()
      .isVisible()
      .catch(() => false);
    const tipoRelacionTextVisible = await modalAsociar.getByText(rxTipoRelacion).first().isVisible().catch(() => false);
    return (tipoRelacionadoVisible || tipoRelacionadoTextVisible) && (tipoRelacionVisible || tipoRelacionTextVisible);
  };

  let formularioRelacionListo = await camposRelacionVisible();
  for (let i = 1; i <= 6 && !formularioRelacionListo; i++) {
    await btnAnadir.scrollIntoViewIfNeeded().catch(() => { });
    await btnAnadir.click({ force: true }).catch(() => { });
    await page.waitForTimeout(1000);
    formularioRelacionListo = await camposRelacionVisible();
  }
  if (!formularioRelacionListo) {
    throw new Error("[CRITICO] No aparecieron los campos 'Tipo de relacionado/Tipo de relación' tras pulsar 'Añadir'.");
  }

  await seleccionarDropdownIndexEnModal(modalAsociar, rxTipoRelacionado, 0);

  let tipoRelacionSeleccionado = false;
  for (let i = 1; i <= 5; i++) {
    await clickReintentarPorLabelEnModal(modalAsociar, rxTipoRelacion).catch(() => { });
    const ok = await seleccionarDropdownIndexEnModal(modalAsociar, rxTipoRelacion, 0, 10000)
      .then(() => true)
      .catch(() => false);
    if (ok) {
      tipoRelacionSeleccionado = true;
      break;
    }
    await page.waitForTimeout(700);
  }
  if (!tipoRelacionSeleccionado) {
    throw new Error("[CRITICO] No se pudo seleccionar 'Tipo de relación' para el relacionado.");
  }

  const btnAceptarAsociacion = modalAsociar
    .locator('button.p-confirm-dialog-accept, button:has-text("Aceptar"), button[data-pc-name="acceptbutton"]')
    .last();
  await btnAceptarAsociacion.waitFor({ state: "visible", timeout: 15000 });
  await btnAceptarAsociacion.click({ force: true });
  await modalAsociar.waitFor({ state: "hidden", timeout: 15000 }).catch(() => { });

  return true;
}

