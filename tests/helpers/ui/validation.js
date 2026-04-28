import { clickReintentarListaSiVisible, seleccionarDropdown } from './dropdowns-basic.js';

export async function corregirErroresPrimeVue(page, accionesPorError = {}) {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(350);

    const errores = await page.locator('div[role="alert"] .p-inline-message-text').allInnerTexts();
    if (!errores || errores.length === 0) {
        console.log("âœ” No hay mensajes de error visibles.");
        return false;
    }

    console.log("âš  ERRORES DETECTADOS:", errores);
    let algunaAccionEjecutada = false;

    for (const error of errores) {
        const accion = accionesPorError[error];
        if (accion) {
            console.log(`ðŸ”§ Ejecutando acciÃ³n para: "${error}"`);
            await accion();
            algunaAccionEjecutada = true;
        } else {
            console.log(`âš  No hay acciÃ³n definida para: "${error}"`);
        }
    }

    return algunaAccionEjecutada;
}

export async function seleccionarNoSiVacioSelectButton(page, tituloBloque) {
    const label = page.getByText(tituloBloque, { exact: true });
    await label.waitFor({ state: 'visible', timeout: 15000 });

    const bloque = label.locator('xpath=ancestor::div[contains(@class,"flex-col")][1]');
    const grupo = bloque.locator('.p-selectbutton, [data-pc-name="selectbutton"]').first();
    await grupo.waitFor({ state: 'visible', timeout: 15000 });

    const seleccionado = grupo.locator(
        '[aria-pressed="true"], [aria-checked="true"], .p-highlight, .p-button.p-highlight'
    );

    if (await seleccionado.count() === 0) {
        const candidatoNo = grupo.locator(
            'button, [role="button"], .p-button, .p-button-label'
        ).filter({ hasText: 'No' }).first();

        await candidatoNo.waitFor({ state: 'visible', timeout: 15000 });
        const clickable = await candidatoNo.evaluateHandle((el) => el.closest('button,[role="button"],.p-button') ?? el);
        await clickable.asElement().scrollIntoViewIfNeeded();
        await clickable.asElement().click({ force: true });
    }
}

export async function seleccionarNoSiVacioPorPregunta(page, textoPregunta, opcionNo = "No", afterAction) {
    const spanPregunta = page.locator('label span', { hasText: textoPregunta }).first();
    const preguntaVisible = await spanPregunta.waitFor({ state: "attached", timeout: 3000 })
        .then(() => true)
        .catch(() => false);
    if (!preguntaVisible) return false;
    await spanPregunta.scrollIntoViewIfNeeded();

    const label = spanPregunta.locator('xpath=ancestor::label[1]');
    const grupo = label.locator(
        'xpath=following-sibling::div[contains(@class,"p-selectbutton") or @data-pc-name="selectbutton"][1]'
    );
    await grupo.waitFor({ state: "attached", timeout: 5000 });
    await grupo.scrollIntoViewIfNeeded();

    const seleccionado = grupo.locator('[role="radio"][aria-checked="true"], [aria-pressed="true"], .p-highlight');
    if ((await seleccionado.count()) === 0) {
        const radioNo = grupo.locator(`[role="radio"][aria-label="${opcionNo}"]`).first();
        await radioNo.waitFor({ state: "attached", timeout: 15000 });
        await radioNo.scrollIntoViewIfNeeded();
        await radioNo.click({ force: true });
        await page.waitForTimeout(200);
    }

    if (typeof afterAction === "function") {
        await afterAction();
    }
}

export async function completarSiHayRequeridos(page, actions, opts = {}) {
    const maxRondas = opts.maxRondas ?? 4;
    const pausaMs = opts.pausaMs ?? 300;
    const debug = opts.debug ?? true;
    const root = opts.root ?? page;

    const requeridoGlobal = root.locator('text=Es requerido').first();
    await requeridoGlobal.waitFor({ state: "visible", timeout: 1500 }).catch(() => { });
    const hayRequeridos = await requeridoGlobal.isVisible().catch(() => false);
    if (!hayRequeridos) return false;

    console.log("âš ï¸ Se detectaron campos requeridos. Intentando completarlos...");
    let hizoAlgo = false;

    const findLabelNode = (labelText) => {
        const exact = root.locator(`xpath=//*[normalize-space(.)='${labelText}']`).first();
        const contains = root.locator(`xpath=//*[contains(normalize-space(.),'${labelText}')]`).first();
        return { exact, contains };
    };

    const findFieldContainer = (labelNode) => {
        const withRequired = labelNode.locator(
            `xpath=ancestor::*[self::div or self::section][.//*[normalize-space(.)='Es requerido']][1]`
        );
        const withControl = labelNode.locator(
            `xpath=ancestor::*[self::div or self::section][.//input or .//textarea or .//div[contains(@class,'p-dropdown')] or .//*[@data-pc-name='dropdown']][1]`
        );
        return { withRequired, withControl };
    };

    const isFieldRequired = async (container) => {
        const inside = container.locator(`text=Es requerido`).first();
        if (await inside.isVisible().catch(() => false)) return true;

        const sibling = container.locator(`xpath=following-sibling::*[1]//*[normalize-space(.)='Es requerido']`).first();
        if (await sibling.isVisible().catch(() => false)) return true;

        const invalid = container.locator(`[aria-invalid="true"], .p-invalid`).first();
        if (await invalid.isVisible().catch(() => false)) return true;

        return false;
    };

    for (let ronda = 1; ronda <= maxRondas; ronda++) {
        const sigueHabiendo = await requeridoGlobal.isVisible().catch(() => false);
        if (!sigueHabiendo) break;

        console.log(`ðŸ› ï¸ Ronda ${ronda}/${maxRondas} de completar requeridos...`);

        for (const [labelText, fn] of Object.entries(actions)) {
            const { exact, contains } = findLabelNode(labelText);
            const exactCount = await exact.count().catch(() => 0);
            const containsCount = await contains.count().catch(() => 0);

            const labelNode = exactCount ? exact : (containsCount ? contains : null);
            if (!labelNode) {
                if (debug) console.log(`   Â· '${labelText}': no encontrÃ© nodo con ese texto`);
                continue;
            }

            const { withRequired, withControl } = findFieldContainer(labelNode);
            const c1 = await withRequired.count().catch(() => 0);
            const container = c1 ? withRequired.first() : withControl.first();

            if (!(await container.count().catch(() => 0))) {
                if (debug) console.log(`   Â· '${labelText}': no encontrÃ© contenedor del campo`);
                continue;
            }

            await container.scrollIntoViewIfNeeded().catch(() => { });
            const required = await isFieldRequired(container);
            if (!required) {
                if (debug) console.log(`   Â· '${labelText}': estÃ¡ presente pero NO lo detecto como requerido ahora`);
                continue;
            }

            console.log(`âž¡ï¸ Completando requerido: ${labelText}`);
            await fn();
            hizoAlgo = true;
            await page.waitForTimeout(pausaMs);
        }

        await page.waitForTimeout(pausaMs);
    }

    return hizoAlgo;
}

export async function validarCorreoPredeterminadoYCorregir(page) {
    const msgCorreoPred = page.getByText(/Debe agregar un Correo Electr(?:o|\u00f3)nico como predeterminado/i);
    const msgVisible = await msgCorreoPred.waitFor({ state: "visible", timeout: 3000 })
        .then(() => true)
        .catch(() => false);
    if (!msgVisible) return false;

    const cardCorreo = page
        .locator('div.p-panel, div.p-fieldset, div.card')
        .filter({ has: page.getByText(/Correo Electr(?:o|\u00f3)nico/i) })
        .first();

    let btnLapiz = cardCorreo.locator('button:has([class*="pi-pencil"]), button.p-button-icon-only').first();
    let btnVisible = await btnLapiz.waitFor({ state: "visible", timeout: 4000 })
        .then(() => true)
        .catch(() => false);
    if (!btnVisible) {
        btnLapiz = page.locator('button:has([class*="pi-pencil"]):visible').first();
        btnVisible = await btnLapiz.waitFor({ state: "visible", timeout: 4000 })
            .then(() => true)
            .catch(() => false);
    }
    if (!btnVisible) return false;
    await btnLapiz.click({ force: true });

    const modal = page.locator('.p-dialog:visible, [role="dialog"]').first();
    const modalVisible = await modal.waitFor({ state: "visible", timeout: 7000 })
        .then(() => true)
        .catch(() => false);
    if (!modalVisible) return false;
    await clickReintentarListaSiVisible(page, "Tipo");

    const switchRoot = modal
        .getByText(/Predeterminado/i)
        .locator('xpath=following::*[contains(@class,"p-inputswitch")][1]');

    await switchRoot.waitFor({ state: "visible", timeout: 10000 });
    await switchRoot.scrollIntoViewIfNeeded();

    const inputSwitch = switchRoot.locator('input[type="checkbox"], input[type="radio"]');
    const isChecked = await inputSwitch.isChecked().catch(() => false);
    if (!isChecked) {
        await switchRoot.click({ force: true });
    }

    const modalContacto = page.locator('.p-dialog:visible, [role="dialog"]:visible').first();
    const labelTipo = modalContacto.getByText(/Tipo/i).first();
    const labelVisible = await labelTipo.isVisible().catch(() => false);
    if (labelVisible) {
        const fieldTipo = labelTipo.locator('xpath=ancestor::div[contains(@class,"field") or contains(@class,"flex")][1]');
        const dropdown = fieldTipo.locator('xpath=.//div[contains(@class,"p-dropdown")][1]');
        const labelEl = dropdown.locator(".p-dropdown-label");
        const btnRetryTipo = fieldTipo.locator(
            'button:has-text("Reintentar buscar lista"), button:has-text("Reintentar"), button.p-button-warning, button[data-pc-severity="warning"]'
        ).first();

        for (let i = 0; i < 4; i++) {
            const texto = (await labelEl.innerText().catch(() => "")).trim();
            const clases = (await labelEl.getAttribute("class").catch(() => "")) || "";
            const sinSeleccion = !texto || /^seleccione/i.test(texto) || clases.includes("p-placeholder");
            if (!sinSeleccion) break;

            const retryVisible = await btnRetryTipo.isVisible().catch(() => false);
            if (retryVisible) {
                await btnRetryTipo.click({ force: true });
                await page.waitForTimeout(700);
            }

            await seleccionarDropdown(page, "Tipo", { index: 1, root: modalContacto }).catch(() => { });
            await page.waitForTimeout(300);
        }
    }

    await page.getByRole("button", { name: "Aceptar" }).click();
    return true;
}
