import { getDropdown } from './shared.js';

export async function esperarDropdownCargado(page, label) {
    console.log(`â³ Esperando items del dropdown '${label}'...`);

    const root = await getDropdown(page, label);
    const combobox = root.locator('[role="combobox"]');
    const panelId = await combobox.getAttribute("aria-controls");

    if (!panelId) {
        throw new Error(`âŒ No se encontrÃ³ aria-controls para '${label}'.`);
    }

    const panel = page.locator(`#${panelId}`);
    await panel.locator(".p-dropdown-item").first().waitFor({ timeout: 15000 });

    console.log(`âœ” Items cargados para '${label}'.`);
}

export async function manejarReintentar(page) {
    const botonReintentar = page.getByRole('button', { name: 'Reintentar buscar lista' });

    for (let i = 0; i < 10; i++) {
        if (await botonReintentar.isVisible({ timeout: 1000 }).catch(() => false)) {
            console.log("ðŸ” BotÃ³n 'Reintentar buscar lista' detectado â†’ clic...");
            await botonReintentar.click();
            await page.waitForTimeout(1500);
        } else {
            break;
        }
    }
}

export async function seleccionarDropdown(page, label, opciones = {}) {
    const { texto, index, timeout = 15000, root } = opciones;
    const scope = root ?? page;

    console.log(`âž¡ï¸ Seleccionando [${label}]...`);

    const labelNode = scope.getByText(label, { exact: true }).first();
    const contenedor = labelNode.locator(
        "xpath=ancestor::div[contains(@class,'field') or contains(@class,'grid') or contains(@class,'col') or @data-slot='dt-dd' or contains(@class,'flex')][1]"
    );
    await contenedor.waitFor({ state: "visible", timeout });

    const dropdown = contenedor.locator('.p-dropdown');
    await dropdown.waitFor({ state: "visible" });

    const trigger = dropdown.locator('.p-dropdown-trigger');
    await trigger.click({ force: true });

    const combobox = dropdown.locator('[role="combobox"]');
    const panelId = await combobox.getAttribute("aria-controls");

    const panel = panelId
        ? page.locator(`#${panelId}`)
        : page.locator('.p-dropdown-panel.p-component')
            .filter({ has: page.locator('.p-dropdown-items') })
            .filter({ hasNot: page.locator('[style*="display: none"]') })
            .first();

    const panelVisible = await panel.waitFor({ state: "visible", timeout })
        .then(() => true)
        .catch(() => false);
    if (!panelVisible) {
        await trigger.click({ force: true });
        await panel.waitFor({ state: "visible", timeout });
    }

    const items = panel.locator('.p-dropdown-item');
    await items.first().waitFor({ timeout });

    const textos = await items.allInnerTexts();
    console.log(`   âœ” Items encontrados en '${label}':`, textos);

    let opcion;
    if (texto) {
        opcion = items.filter({ hasText: texto }).first();
    } else if (index !== undefined) {
        opcion = items.nth(index);
    } else {
        throw new Error(`âŒ Especifica texto o index para '${label}'`);
    }

    await opcion.waitFor({ state: "visible" });
    const handle = await opcion.elementHandle();
    await page.evaluate(el => el.click(), handle);

    console.log(`âœ” Seleccionado ${texto ?? index}`);
}

export async function seleccionarDefaultSiVacio(page, label, indexDefault = 0, accionDespues) {
    const dropdown = await getDropdown(page, label);
    const labelEl = dropdown.locator('.p-dropdown-label');
    const texto = (await labelEl.innerText().catch(() => '')).trim();
    const clases = (await labelEl.getAttribute('class').catch(() => '')) || '';

    const sinSeleccion = !texto || /^seleccione/i.test(texto) || clases.includes("p-placeholder");

    if (sinSeleccion) {
        console.log(`âš  '${label}' vacÃ­o â†’ seleccionando index ${indexDefault}`);
        await seleccionarDropdown(page, label, { index: indexDefault });

        if (typeof accionDespues === "function") {
            console.log(`â–¶ Ejecutando acciÃ³n despuÃ©s de seleccionar default para '${label}'`);
            await accionDespues();
        }

        return true;
    }

    console.log(`âœ” '${label}' ya tenÃ­a valor: "${texto}"`);
    return false;
}

export async function clickReintentarListaSiVisible(page, labelText, timeoutMs = 800) {
    let label = null;

    if (labelText instanceof RegExp) {
        label = page.locator('label').filter({ hasText: labelText }).first();
    } else {
        const texto = String(labelText);
        const escaped = texto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\\\?/g, '.');
        const pattern = new RegExp(escaped, 'i');
        label = page.locator('label').filter({ hasText: pattern }).first();

        if ((await label.count().catch(() => 0)) === 0) {
            label = page.getByText(pattern).first();
        }
    }

    if (label && (await label.count().catch(() => 0)) > 0) {
        let field = label.locator(
            'xpath=ancestor::div[@data-slot="dt-dd" or (contains(@class,"flex") and contains(@class,"flex-col")) or contains(@class,"field")][1]'
        );
        if ((await field.count().catch(() => 0)) === 0) {
            field = label.locator('xpath=ancestor::div[.//div[contains(@class,"p-inputgroup")]][1]');
        }

        const btnScoped = field.locator(
            'div.p-inputgroup:has-text("Reintentar buscar lista") button,' +
            'button:has-text("Reintentar buscar lista"),' +
            'button:has-text("Reintentar"),' +
            'button[aria-label*="Reintentar"]'
        ).first();

        const visibleScoped = await btnScoped.waitFor({ state: 'visible', timeout: timeoutMs })
            .then(() => true)
            .catch(() => false);
        if (visibleScoped) {
            await btnScoped.scrollIntoViewIfNeeded().catch(() => { });
            await btnScoped.click({ force: true });
            return true;
        }
    }

    const btnGlobal = page
        .locator('div.p-inputgroup:has-text("Reintentar buscar lista") button:visible')
        .first();
    const visibleGlobal = await btnGlobal.waitFor({ state: 'visible', timeout: timeoutMs })
        .then(() => true)
        .catch(() => false);
    if (!visibleGlobal) return false;

    await btnGlobal.scrollIntoViewIfNeeded().catch(() => { });
    await btnGlobal.click({ force: true });
    return true;
}

export async function seleccionarDropdownConReintentoYReintentarBtn(page, labelText, texto, maxIntentos = 8) {
    const label = page
        .locator(`xpath=//label[normalize-space(.)='${labelText}' or .//*[normalize-space(.)='${labelText}']]`)
        .first();

    const field = label.locator(
        'xpath=ancestor::div[contains(@class,"flex") and contains(@class,"flex-col") and contains(@class,"gap-2")][1]'
    );
    const dropdown = field.locator('div.p-dropdown,[data-pc-name="dropdown"]').first();

    for (let i = 1; i <= maxIntentos; i++) {
        await dropdown.click({ force: true });

        const panel = page.locator('.p-dropdown-panel:visible, [data-pc-section="panel"]:visible').first();
        await panel.waitFor({ state: "visible", timeout: 8000 });

        const items = panel.locator('li[role="option"], .p-dropdown-item, [data-pc-section="item"]');
        const emptyMsg = panel.locator('.p-dropdown-empty-message, [data-pc-section="emptyMessage"]').first();

        await Promise.race([
            items.first().waitFor({ state: "visible", timeout: 1500 }).catch(() => { }),
            emptyMsg.waitFor({ state: "visible", timeout: 1500 }).catch(() => { }),
        ]);

        const opt = items.filter({ hasText: texto }).first();
        if (await opt.isVisible().catch(() => false)) {
            await opt.click();
            return;
        }

        const panelVisible = await panel.isVisible().catch(() => false);
        if (panelVisible) await page.mouse.click(10, 10).catch(() => { });

        const clickeado = await clickReintentarListaSiVisible(page, labelText, 700);
        await page.waitForTimeout(clickeado ? 500 : 400);
    }

    throw new Error(`âŒ No se pudo seleccionar '${texto}' en '${labelText}' tras ${maxIntentos} intentos.`);
}
