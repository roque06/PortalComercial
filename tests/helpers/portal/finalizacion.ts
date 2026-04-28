import type { Locator, Page } from '@playwright/test';

export function modalSolicitudFinalizadaLocator(page: Page) {
  return page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i })
    .last();
}

export async function cerrarModalFinalizadaRapido(page: Page) {
  const modal = modalSolicitudFinalizadaLocator(page);
  const visible = await modal.isVisible().catch(() => false);
  if (!visible) return true;

  const btnCerrar = modal.getByRole('button', { name: /^Cerrar$/i }).first();
  const btnAceptar = modal.getByRole('button', { name: /^Aceptar$/i }).first();
  const btnAny = modal.locator('button:has-text("Cerrar"), button:has-text("Aceptar")').first();

  for (const btn of [btnCerrar, btnAceptar, btnAny]) {
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

export async function cerrarModalSolicitudFinalizada(
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

export async function estadoSolicitudFinalizadaVisible(page: Page) {
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

export async function estadoPendienteFinalizarVisible(page: Page) {
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

export async function confirmarFinalizacionSiHayPopup(page: Page) {
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

export async function clickBotonFinalizarRobusto(page: Page) {
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

export async function irAPasoFinalDesdeWizard(
  page: Page,
  options: { stepNumber: number; esperarFinActualizandoSolicitud: (page: Page, timeoutMs?: number) => Promise<unknown> }
) {
  const candidatoPasoFinal = page.locator(
    `li.p-steps-item:has-text("${options.stepNumber}"), [data-pc-section="menuitem"]:has-text("${options.stepNumber}"), .p-steps-item:has(.p-steps-number:has-text("${options.stepNumber}"))`
  ).first();
  const visible = await candidatoPasoFinal.isVisible().catch(() => false);
  if (!visible) return false;

  await candidatoPasoFinal.click({ force: true }).catch(() => { });
  await options.esperarFinActualizandoSolicitud(page, 12000).catch(() => false);
  await page.waitForTimeout(220);
  return true;
}

export async function avanzarAPasoFinalSiHaceFalta(
  page: Page,
  options: {
    stepNumber: number;
    getBotonContinuar: (page: Page) => Locator;
    esperarFinActualizandoSolicitud: (page: Page, timeoutMs?: number) => Promise<unknown>;
    beforeContinuar?: (page: Page) => Promise<unknown>;
  }
) {
  const btnFinalizar = page.getByRole('button', { name: /^Finalizar$/i }).first();
  for (let i = 1; i <= 3; i++) {
    const finalizarVisible = await btnFinalizar.isVisible().catch(() => false);
    const finalizarEnabled = finalizarVisible ? await btnFinalizar.isEnabled().catch(() => false) : false;
    if (finalizarVisible && finalizarEnabled) return true;

    await irAPasoFinalDesdeWizard(page, {
      stepNumber: options.stepNumber,
      esperarFinActualizandoSolicitud: options.esperarFinActualizandoSolicitud,
    }).catch(() => false);
    const finalizaTrasPaso = await btnFinalizar.isVisible().catch(() => false);
    if (finalizaTrasPaso) return true;

    if (options.beforeContinuar) {
      await options.beforeContinuar(page).catch(() => false);
    }

    const btnContinuar = options.getBotonContinuar(page);
    const continuarVisible = await btnContinuar.isVisible().catch(() => false);
    const continuarEnabled = continuarVisible ? await btnContinuar.isEnabled().catch(() => false) : false;
    if (!continuarVisible || !continuarEnabled) break;

    await btnContinuar.scrollIntoViewIfNeeded().catch(() => { });
    await btnContinuar.click({ force: true, timeout: 2200 }).catch(() => { });
    await options.esperarFinActualizandoSolicitud(page, 18000).catch(() => false);
    await confirmarFinalizacionSiHayPopup(page).catch(() => false);
    await page.waitForTimeout(220);
  }

  return await btnFinalizar.isVisible().catch(() => false);
}

export async function finalizarSolicitudSiPendiente<TRegistro extends { identificacion: string }>(
  page: Page,
  registro: TRegistro,
  options: {
    timeoutMs: number;
    maxIntentos: number;
    stepNumber: number;
    getBotonContinuar: (page: Page) => Locator;
    esperarFinActualizandoSolicitud: (page: Page, timeoutMs?: number) => Promise<unknown>;
    beforeContinuar?: (page: Page) => Promise<unknown>;
    beforeEachIntent?: (page: Page) => Promise<unknown>;
    afterFinalizarClick?: (page: Page) => Promise<unknown>;
    onNoFinalizarVisible?: (page: Page) => Promise<unknown>;
  }
) {
  const inicio = Date.now();
  const modalFinalizadaVisibleInicial = await modalSolicitudFinalizadaLocator(page).isVisible().catch(() => false);
  if (modalFinalizadaVisibleInicial) {
    const cerroInicial = await cerrarModalFinalizadaRapido(page).catch(() => false);
    if (cerroInicial) return true;
  }

  const yaFinalizada = await estadoSolicitudFinalizadaVisible(page);
  if (yaFinalizada) return true;

  for (let intento = 1; intento <= options.maxIntentos; intento++) {
    if (Date.now() - inicio >= options.timeoutMs) break;
    console.log(`[Finalizacion][${registro.identificacion}] Intento ${intento}/${options.maxIntentos} para finalizar solicitud.`);

    const cerroModalPrevio = await cerrarModalSolicitudFinalizada(page, {
      timeoutMs: 5000,
      waitForAppearMs: 900,
    }).catch(() => false);
    if (cerroModalPrevio) return true;

    await avanzarAPasoFinalSiHaceFalta(page, {
      stepNumber: options.stepNumber,
      getBotonContinuar: options.getBotonContinuar,
      esperarFinActualizandoSolicitud: options.esperarFinActualizandoSolicitud,
      beforeContinuar: options.beforeContinuar,
    }).catch(() => false);

    if (options.beforeEachIntent) {
      await options.beforeEachIntent(page).catch(() => false);
    }

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

      if (options.afterFinalizarClick) {
        await options.afterFinalizarClick(page).catch(() => false);
      }

      const finalizadaInmediata = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
      if (finalizadaInmediata) return true;

      const modalVisibleTrasFinalizar = await modalSolicitudFinalizadaLocator(page).isVisible().catch(() => false);
      if (!modalVisibleTrasFinalizar) {
        await options.esperarFinActualizandoSolicitud(page, 4000).catch(() => false);
        const finalizadaTrasEsperaCorta = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
        if (finalizadaTrasEsperaCorta) return true;
      }

      const cerroModalTrasFinalizar = await cerrarModalSolicitudFinalizada(page, {
        timeoutMs: 12000,
        waitForAppearMs: 3500,
      }).catch(() => false);
      if (cerroModalTrasFinalizar) return true;
    } else if (options.onNoFinalizarVisible) {
      await options.onNoFinalizarVisible(page).catch(() => false);
    }

    const finalizadaSinModal = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
    if (finalizadaSinModal) return true;

    await options.esperarFinActualizandoSolicitud(page, 18000).catch(() => false);
    await confirmarFinalizacionSiHayPopup(page).catch(() => false);
    const cerroModalPostEspera = await cerrarModalSolicitudFinalizada(page, {
      timeoutMs: 10000,
      waitForAppearMs: 7000,
    }).catch(() => false);
    if (cerroModalPostEspera) return true;
    await cerrarModalFinalizadaRapido(page).catch(() => false);

    const finalizada = await estadoSolicitudFinalizadaVisible(page).catch(() => false);
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
