import { clickReintentarListaSiVisible } from './dropdowns-basic.js';

async function cerrarModalCancelarProcesoSiVisible(page) {
    const modalCancelarProceso = page
        .locator('.p-dialog:visible, [role="dialog"]:visible')
        .filter({ hasText: /Cancelar proceso de solicitud/i })
        .first();
    const visible = await modalCancelarProceso.isVisible().catch(() => false);
    if (!visible) return false;

    const btnMantenerProceso = modalCancelarProceso.getByRole('button', { name: /^Cancelar$/i }).first();
    const btnCerrar = modalCancelarProceso.locator('.p-dialog-header-close, [aria-label*="close" i], [aria-label*="cerr" i]').first();

    if (await btnMantenerProceso.isVisible().catch(() => false)) {
        await btnMantenerProceso.click({ force: true }).catch(() => { });
    } else if (await btnCerrar.isVisible().catch(() => false)) {
        await btnCerrar.click({ force: true }).catch(() => { });
    } else {
        await page.keyboard.press('Escape').catch(() => { });
    }

    await modalCancelarProceso.waitFor({ state: 'hidden', timeout: 4000 }).catch(() => { });
    return true;
}

export async function seleccionarDropdownConReintento(page, labelText, select, opts = {}) {
    const maxIntentos = opts.maxIntentos ?? 8;
    const esperaMs = opts.esperaMs ?? 1200;
    const timeoutPanelMs = opts.timeoutPanelMs ?? 15000;
    const timeoutDropdownMs = opts.timeoutDropdownMs ?? 60000;

    const label = page
        .locator(`xpath=//label[normalize-space(.)='${labelText}' or .//*[normalize-space(.)='${labelText}']]`)
        .first();

    await label.waitFor({ state: "visible", timeout: 30000 }).catch(() => { });
    if ((await label.count()) === 0) throw new Error(`âŒ No se encontrÃ³ el label '${labelText}'.`);

    let container = label.locator(
        'xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]'
    );

    if ((await container.count().catch(() => 0)) === 0) {
        container = label.locator(
            'xpath=ancestor::div[contains(@class,"flex") and contains(@class,"flex-col") and contains(@class,"gap-2")][1]'
        );
    }

    await container.waitFor({ state: "visible", timeout: 30000 });
    await container.scrollIntoViewIfNeeded().catch(() => { });

    const dropdownRoot = container.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();
    await dropdownRoot.waitFor({ state: "visible", timeout: timeoutDropdownMs });

    const combobox = dropdownRoot.locator('[role="combobox"]').first();

    let outer = container.locator(
        'xpath=ancestor::*[self::div or self::span][contains(@class,"p-inputgroup") or contains(@class,"grid") or contains(@class,"col")][1]'
    );

    if ((await outer.count().catch(() => 0)) === 0) outer = container;

    const btnReintentar = outer
        .locator('button.p-button-warning, button:has-text("Reintentar"), button:has-text("Reintentar buscar lista")')
        .first();

    const abrirPanel = async () => {
        await dropdownRoot.click({ force: true });

        const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
        if (!panelId) {
            const fallback = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').first();
            await fallback.waitFor({ state: "visible", timeout: timeoutPanelMs });
            return fallback;
        }

        const panel = page.locator(`#${panelId}`);
        await panel.waitFor({ state: "attached", timeout: timeoutPanelMs });
        await panel.waitFor({ state: "visible", timeout: timeoutPanelMs });
        return panel;
    };

    const esperarPanelEstable = async (panel) => {
        const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
        const emptyMsg = panel.locator('.p-dropdown-empty-message, [data-pc-section="emptyMessage"]');

        await Promise.race([
            items.first().waitFor({ state: "visible", timeout: 4000 }).catch(() => { }),
            emptyMsg.waitFor({ state: "visible", timeout: 4000 }).catch(() => { }),
        ]);

        return { items, emptyMsg };
    };

    const cerrarPanel = async () => {
        const panelVisible = await page.locator(".p-dropdown-panel:visible, [data-pc-section='panel']:visible").first().isVisible().catch(() => false);
        if (panelVisible) {
            await page.mouse.click(10, 10).catch(() => { });
        }
        await page.waitForTimeout(150);
    };

    for (let intento = 1; intento <= maxIntentos; intento++) {
        const panel = await abrirPanel();
        const { items, emptyMsg } = await esperarPanelEstable(panel);

        if (select.texto !== undefined) {
            const opt = items.filter({ hasText: select.texto }).first();

            if (await opt.isVisible().catch(() => false)) {
                await opt.scrollIntoViewIfNeeded().catch(() => { });
                await opt.click();
                return;
            }

            await emptyMsg.isVisible().catch(() => false);
        }

        if (typeof select.index === "number") {
            const count = await items.count().catch(() => 0);
            if (count > select.index) {
                await items.nth(select.index).click();
                return;
            }
        }

        await cerrarPanel();
        await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);

        const retryVisible = await btnReintentar.isVisible().catch(() => false);
        if (retryVisible) {
            console.log(`â†» '${labelText}': no aparece selecciÃ³n. Click Reintentar (${intento}/${maxIntentos})`);
            await btnReintentar.click({ force: true });
            await page.waitForTimeout(esperaMs);
        } else {
            console.log(`â†» '${labelText}': no aparece selecciÃ³n y NO hay botÃ³n naranja. Reabriendo... (${intento}/${maxIntentos})`);
            await page.waitForTimeout(esperaMs);
        }
    }

    throw new Error(`âŒ No se pudo seleccionar '${select.texto ?? select.index}' en '${labelText}' tras ${maxIntentos} intentos.`);
}

export async function asegurarIdentificacionHabilitada(page, opts = {}) {
    const maxIntentos = opts.maxIntentos ?? 8;
    const timeoutCampoMs = opts.timeoutCampoMs ?? 60000;
    const timeoutPanelMs = opts.timeoutPanelMs ?? 10000;
    const pausaMs = opts.pausaMs ?? 900;

    const labelPais = page.locator("label").filter({ hasText: /Pa[ií]s de identificaci[oó]n/i }).first();
    const labelTipo = page.locator("label").filter({ hasText: /Tipo de identificaci[oó]n/i }).first();

    let labelIdent = labelPais;
    const paisVisible = await labelPais.waitFor({ state: "visible", timeout: 5000 }).then(() => true).catch(() => false);
    if (!paisVisible) {
        await labelTipo.waitFor({ state: "visible", timeout: timeoutCampoMs });
        labelIdent = labelTipo;
    }

    const labelNumero = page.locator("label").filter({ hasText: /N[uú]mero de identificaci[oó]n/i }).first();
    const getInputNumeroId = () => labelNumero.locator("xpath=following::input[1]").first();

    const esperarHabilitado = async (locator, timeout = 15000) => {
        await locator.waitFor({ state: "visible", timeout });
        const handle = await locator.elementHandle();
        if (!handle) return false;
        return await page
            .waitForFunction((el) => !!el && el.disabled === false && el.readOnly !== true, handle, { timeout })
            .then(() => true)
            .catch(() => false);
    };

    const slotIdent = labelIdent.locator("xpath=following-sibling::*[1]").first();
    const getBtnRetryIdent = () =>
        slotIdent.locator(
            'div.p-inputgroup:has-text("Reintentar buscar lista") button,' +
            'button:has-text("Reintentar"),' +
            'button:has-text("Reintentar buscar lista"),' +
            'button[aria-label*="Reintentar"]'
        ).first();
    const getDropdownIdent = () => slotIdent.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();

    let inputNumeroId = getInputNumeroId();

    for (let i = 1; i <= maxIntentos; i++) {
        const enabled = await inputNumeroId.isEnabled().catch(() => false);
        if (enabled) return;

        const btnRetryIdent = getBtnRetryIdent();
        const retryVisible = await btnRetryIdent.isVisible().catch(() => false);
        if (retryVisible) {
            await btnRetryIdent.click({ force: true });
            await page.waitForTimeout(pausaMs);
        } else {
            const clickedPais = await clickReintentarListaSiVisible(page, "País de identificación", 1200).catch(() => false);
            if (!clickedPais) {
                await clickReintentarListaSiVisible(page, "Tipo de identificación", 1200).catch(() => false);
            }
        }

        const dropdownIdent = getDropdownIdent();
        const dropdownVisible = await dropdownIdent.isVisible().catch(() => false);
        const esPais = await labelIdent.filter({ hasText: /Pa[ií]s de identificaci[oó]n/i }).count().catch(() => 0);
        if (dropdownVisible && esPais > 0) {
            const labelDd = dropdownIdent.locator('.p-dropdown-label').first();
            const texto = ((await labelDd.textContent().catch(() => '')) || '').trim();
            const clases = ((await labelDd.getAttribute('class').catch(() => '')) || '');
            const sinSeleccion = !texto || /seleccione|reintentar/i.test(texto) || clases.includes('p-placeholder');

            if (sinSeleccion) {
                await dropdownIdent.click({ force: true }).catch(() => { });
                const panel = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').first();
                const panelVisible = await panel.waitFor({ state: 'visible', timeout: timeoutPanelMs }).then(() => true).catch(() => false);
                if (panelVisible) {
                    const optRD = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]')
                        .filter({ hasText: /Rep(ú|u)blica Dominicana/i })
                        .first();
                    if (await optRD.isVisible().catch(() => false)) {
                        await optRD.click().catch(() => { });
                    } else {
                        const panelOpen = await panel.isVisible().catch(() => false);
                        if (panelOpen) await page.mouse.click(10, 10).catch(() => { });
                    }
                }
            }
        }

        const ok = await esperarHabilitado(inputNumeroId, 7000);
        if (ok) return;
    }

    await page.reload();
    await page.waitForLoadState('networkidle');
    await labelIdent.waitFor({ state: 'visible', timeout: timeoutCampoMs });
    const clicPais = await clickReintentarListaSiVisible(page, 'País de identificación', 2000).catch(() => false);
    if (!clicPais) {
        await clickReintentarListaSiVisible(page, 'Tipo de identificación', 2000).catch(() => false);
    }
    inputNumeroId = getInputNumeroId();
    const okFinal = await esperarHabilitado(inputNumeroId, 20000);
    if (!okFinal) {
        throw new Error("❌ El campo 'Número de identificación' sigue deshabilitado tras reintentos en identificación.");
    }
}

export async function seleccionarDropdownFiltrableConReintentar(page, labelText, select, opts = {}) {
    const maxIntentos = opts.maxIntentos ?? 10;
    const esperaMs = opts.esperaMs ?? 900;
    const timeoutCampoMs = opts.timeoutCampoMs ?? 30000;
    const timeoutPanelMs = opts.timeoutPanelMs ?? 20000;
    const usarFiltro = opts.usarFiltro ?? true;

    const label = page.locator(
        `xpath=//label[normalize-space(.)='${labelText}' or .//*[normalize-space(.)='${labelText}']]`
    ).first();
    await label.waitFor({ state: "visible", timeout: timeoutCampoMs });

    let container = label.locator(
        'xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]'
    );
    if ((await container.count().catch(() => 0)) === 0) container = label.locator("xpath=ancestor::div[1]");
    await container.waitFor({ state: "visible", timeout: timeoutCampoMs });
    await container.scrollIntoViewIfNeeded().catch(() => { });

    const dropdownRoot = container.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();
    await dropdownRoot.waitFor({ state: "visible", timeout: timeoutCampoMs });

    const combobox = dropdownRoot.locator('[role="combobox"]').first();

    let outer = container.locator(
        'xpath=ancestor::*[self::div or self::span][contains(@class,"p-inputgroup") or contains(@class,"grid") or contains(@class,"col")][1]'
    );
    if ((await outer.count().catch(() => 0)) === 0) outer = container;

    const btnReintentar = outer.locator(
        'button.p-button-warning,' +
        'button[data-pc-severity="warning"],' +
        'button:has-text("Reintentar buscar lista"),' +
        'button:has-text("Reintentar"),' +
        'button[title*="Reintentar" i]'
    ).first();

    const abrirPanel = async () => {
        const targets = [
            dropdownRoot.locator('.p-dropdown-trigger, [data-pc-section="trigger"]').first(),
            combobox,
            dropdownRoot.locator('.p-dropdown-label, [data-pc-section="label"]').first(),
            dropdownRoot,
        ];

        for (const target of targets) {
            const visibleTarget = await target.isVisible().catch(() => false);
            if (!visibleTarget) continue;

            await target.click({ force: true }).catch(() => { });
            await page.waitForTimeout(80);
            const modalCancelacion = await cerrarModalCancelarProcesoSiVisible(page).catch(() => false);
            if (modalCancelacion) continue;

            await target.click({ force: true }).catch(() => { });

            const panelId = await combobox.getAttribute("aria-controls").catch(() => null);
            const quickTimeout = Math.min(timeoutPanelMs, 2500);

            if (panelId) {
                const panelById = page.locator(`#${panelId}`);
                const visibleById = await panelById.waitFor({ state: "visible", timeout: quickTimeout })
                    .then(() => true)
                    .catch(() => false);
                if (visibleById) return panelById;
            }

            const fallback = page.locator(".p-dropdown-panel:visible, [data-pc-section='panel']:visible").last();
            const visibleFallback = await fallback.waitFor({ state: "visible", timeout: quickTimeout })
                .then(() => true)
                .catch(() => false);
            if (visibleFallback) return fallback;
        }

        throw new Error("__DROPDOWN_PANEL_NOT_VISIBLE__");
    };

    const cerrarPanel = async () => {
        const panelVisible = await page.locator(".p-dropdown-panel:visible, [data-pc-section='panel']:visible").first().isVisible().catch(() => false);
        if (panelVisible) {
            await page.mouse.click(10, 10).catch(() => { });
        }
        await page.waitForTimeout(150);
    };

    for (let intento = 1; intento <= maxIntentos; intento++) {
        const panel = await abrirPanel().catch(() => null);
        if (!panel) {
            const retryVisibleNoPanel = await btnReintentar.isVisible().catch(() => false);
            if (retryVisibleNoPanel) {
                await btnReintentar.click({ force: true }).catch(() => { });
                await page.waitForTimeout(esperaMs);
            } else {
                await page.keyboard.press("Escape").catch(() => { });
                await page.waitForTimeout(Math.max(120, Math.floor(esperaMs / 2)));
            }
            continue;
        }

        const items = panel.locator("li[role='option'], .p-dropdown-item, [data-pc-section='item']");
        const emptyMsg = panel.locator(".p-dropdown-empty-message, [data-pc-section='emptyMessage']").first();
        const filterInput = panel.locator(
            "input.p-dropdown-filter, .p-dropdown-filter-container input, input.p-inputtext"
        ).first();

        if (usarFiltro && typeof select.texto === "string") {
            const fiVisible = await filterInput.isVisible().catch(() => false);
            if (fiVisible) {
                await filterInput.fill(select.texto).catch(() => { });
                await page.waitForTimeout(200);
            }
        }

        await Promise.race([
            items.first().waitFor({ state: "visible", timeout: 4000 }).catch(() => { }),
            emptyMsg.waitFor({ state: "visible", timeout: 4000 }).catch(() => { }),
        ]);

        if (select.texto !== undefined) {
            const opt = items.filter({ hasText: select.texto }).first();
            if (await opt.isVisible().catch(() => false)) {
                await opt.scrollIntoViewIfNeeded().catch(() => { });
                await opt.click();
                return;
            }

            const countTexto = await items.count().catch(() => 0);
            if (countTexto > 0) {
                await items.first().scrollIntoViewIfNeeded().catch(() => { });
                await items.first().click().catch(() => { });
                return;
            }
        }

        if (typeof select.index === "number") {
            const count = await items.count().catch(() => 0);
            if (count > select.index) {
                await items.nth(select.index).click();
                return;
            }
        }

        await cerrarPanel();

        const retryVisible = await btnReintentar.isVisible().catch(() => false);
        if (retryVisible) {
            console.log(`â†» '${labelText}': lista no cargÃ³ / sin opciÃ³n. Click Reintentar (${intento}/${maxIntentos})`);
            await btnReintentar.click({ force: true });
            await page.waitForTimeout(esperaMs);
        } else {
            console.log(`â†» '${labelText}': sin opciÃ³n y NO hay botÃ³n. Reintentando (${intento}/${maxIntentos})`);
            await page.waitForTimeout(esperaMs);
        }
    }

    throw new Error(`âŒ No se pudo seleccionar '${select.texto ?? select.index}' en '${labelText}' tras ${maxIntentos} intentos.`);
}
