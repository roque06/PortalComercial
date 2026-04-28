import { fieldContainerByLabel, inputInField } from './shared.js';

export async function llenarFecha(page, labelText, fecha = "15-08-2030", opts = {}) {
    const force = opts.force ?? false;
    const label = page.locator(`xpath=//*[normalize-space(text())='${labelText}']`);

    if (!(await label.count())) {
        throw new Error(`âŒ No se encontrÃ³ el label '${labelText}'.`);
    }

    const input = label.locator('xpath=following::input[contains(@class,"p-inputtext")][1]');
    await input.waitFor({ state: "visible" });

    const actual = (await input.inputValue()).trim();
    if (actual !== "" && !force) {
        console.log(`âœ” '${labelText}' ya tiene valor: ${actual}`);
        return;
    }

    console.log(`âŒ¨ Simulando escritura humana en '${labelText}' con '${fecha}'...`);
    await input.click({ clickCount: 1 });
    await page.keyboard.press("Control+A");
    await page.keyboard.press("Delete");
    await input.pressSequentially(fecha, { delay: 80 });
    await input.dispatchEvent("input");
    await input.dispatchEvent("change");
    await input.blur();
    console.log(`âœ” Fecha escrita correctamente: ${fecha}`);
}

export async function llenarFechaSiVisibleYVacia(page, labelOrRegex, fecha, opts = {}) {
    const debug = opts.debug ?? false;
    const labelLocator = labelOrRegex instanceof RegExp
        ? page.getByText(labelOrRegex)
        : page.getByText(String(labelOrRegex));

    let input = labelLocator
        .locator(
            'xpath=ancestor::div[contains(@class,"space-y") or contains(@class,"flex") or contains(@class,"grid") or contains(@class,"col")][1]//input[@role="combobox" or contains(@class,"p-inputtext")][1]'
        )
        .first();

    let visible = await input.isVisible().catch(() => false);
    if (!visible) {
        const fallbackText = (typeof labelOrRegex === "string" && labelOrRegex.trim()) ? labelOrRegex : "Tiempo de exclu";
        const esc = fallbackText.replace(/"/g, "\"");
        input = page.locator(
            `xpath=//*[contains(normalize-space(.),"${esc}")][1]/ancestor::div[contains(@class,"space-y") or contains(@class,"flex") or contains(@class,"grid") or contains(@class,"col")][1]//input[@role="combobox" or contains(@class,"p-inputtext")][1]`
        ).first();
        visible = await input.isVisible().catch(() => false);
        if (!visible) {
            const bySpaceY = page.locator(
                `xpath=//div[contains(@class,"space-y")][.//*[contains(normalize-space(.),"${esc}")]]//input[@role="combobox" or contains(@class,"p-inputtext")][1]`
            );
            input = bySpaceY.first();
            visible = await input.isVisible().catch(() => false);

            if (!visible) {
                const byFieldset = page.locator(
                    `xpath=//fieldset[.//*[contains(normalize-space(.),"Â¿Acepta publicidad") or contains(normalize-space(.),"Acepta publicidad")]]//input[@role="combobox" or contains(@class,"p-inputtext")][1]`
                );
                input = byFieldset.first();
                visible = await input.isVisible().catch(() => false);
            }

            if (!visible) {
                input = page.locator(
                    `xpath=//*[contains(normalize-space(.),"${esc}")][1]/following::input[@role="combobox" or contains(@class,"p-inputtext")][1]`
                ).first();
                visible = await input.isVisible().catch(() => false);
            }
        }
    }

    if (debug) {
        const countCombos = await page.locator('input[role="combobox"]').count().catch(() => 0);
        console.log(`[llenarFechaSiVisibleYVacia] visible=${visible} comboboxCount=${countCombos}`);
    }
    if (!visible) return false;

    const actual = (await input.inputValue().catch(() => "")).trim();
    if (actual) return false;

    await input.scrollIntoViewIfNeeded().catch(() => { });
    await input.click({ force: true });
    await page.keyboard.press('Control+A').catch(() => { });
    await page.keyboard.press('Delete').catch(() => { });
    await input.pressSequentially(fecha, { delay: 40 });
    await input.dispatchEvent('input');
    await input.dispatchEvent('change');
    await input.blur();

    let after = (await input.inputValue().catch(() => "")).trim();
    if (!after) {
        await input.evaluate((el, v) => {
            const inputEl = el instanceof HTMLInputElement ? el : null;
            if (!inputEl) return;
            inputEl.value = v;
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            inputEl.dispatchEvent(new Event('blur', { bubbles: true }));
        }, fecha);
        after = (await input.inputValue().catch(() => "")).trim();
    }

    return Boolean(after);
}

export async function clickSwitch(page, labelText) {
    const slider = page.locator(
        `xpath=//*[contains(normalize-space(.),"${labelText}")]/following::span[contains(@class,"p-inputswitch-slider")][1]`
    );
    await slider.waitFor({ state: "visible" });
    await slider.click({ force: true });
}

export async function llenarCampoSiVacio(page, labelText, valor, accionDespues) {
    console.log(`ðŸ”Ž Revisando campo: '${labelText}'`);

    const label = page.locator(`xpath=//*[normalize-space(text())="${labelText}"]`).first();
    if (!(await label.count())) {
        console.log(`âŒ No se encontrÃ³ el label '${labelText}'`);
        return false;
    }

    const input = label.locator(
        `xpath=following::input[contains(@class,"p-input") or contains(@class,"p-inputtext") or contains(@class,"p-inputnumber-input")][1]`
    );
    await input.waitFor({ state: "visible" });

    const valorActual = (await input.inputValue().catch(() => '')).trim();
    if (!valorActual) {
        console.log(`âš  Campo vacÃ­o â†’ llenando '${valor}'`);
        await input.click();
        await page.keyboard.press("Control+A");
        await page.keyboard.press("Delete");
        await input.pressSequentially(valor, { delay: 60 });
        await input.dispatchEvent("input");
        await input.dispatchEvent("change");
        await input.blur();

        if (typeof accionDespues === "function") {
            console.log(`â–¶ Ejecutando acciÃ³n despuÃ©s de llenar '${labelText}'`);
            await accionDespues();
        }
        return true;
    }

    console.log(`âœ” '${labelText}' ya tenÃ­a valor: '${valorActual}'`);
    return false;
}

export async function llenarInputNumber(page, labelText, valor, accionDespues = null) {
    console.log(`🔢 Llenando inputnumber '${labelText}' → ${valor}`);

    const timeoutMs = 20000;
    const maxIntentos = 4;
    const valorStr = String(valor ?? "");
    const valorNum = Number(valorStr.replace(/[^\d.-]/g, ""));

    const { label, field } = await fieldContainerByLabel(page, labelText);
    await label.waitFor({ state: "visible", timeout: timeoutMs });
    await field.waitFor({ state: "visible", timeout: timeoutMs });

    const input = field
        .locator('input.p-inputnumber-input, input.p-inputtext, input[type="text"], input[type="tel"], input')
        .first();
    await input.waitFor({ state: "visible", timeout: timeoutMs });

    const parseNum = (text) => {
        const cleaned = String(text ?? "").replace(/[^\d.-]/g, "");
        if (!cleaned) return Number.NaN;
        return Number(cleaned);
    };

    const coincideValor = async () => {
        const actual = (await input.inputValue().catch(() => "")).trim();
        if (!actual) return false;

        const actualNum = parseNum(actual);
        if (Number.isFinite(valorNum) && Number.isFinite(actualNum)) {
            return Math.abs(actualNum - valorNum) < 0.0001;
        }

        return actual.includes(valorStr);
    };

    for (let i = 1; i <= maxIntentos; i++) {
        await input.scrollIntoViewIfNeeded().catch(() => { });
        await input.click({ force: true });
        await page.keyboard.press("Control+A").catch(() => { });
        await page.keyboard.press("Delete").catch(() => { });
        await input.pressSequentially(valorStr, { delay: 25 }).catch(() => { });
        await input.dispatchEvent("input").catch(() => { });
        await input.dispatchEvent("change").catch(() => { });
        await input.blur().catch(() => { });
        await page.waitForTimeout(180);

        if (await coincideValor()) break;

        await input.evaluate((el, v) => {
            const inputEl = el instanceof HTMLInputElement ? el : null;
            if (!inputEl) return;
            inputEl.value = String(v);
            inputEl.dispatchEvent(new Event("input", { bubbles: true }));
            inputEl.dispatchEvent(new Event("change", { bubbles: true }));
            inputEl.dispatchEvent(new Event("blur", { bubbles: true }));
        }, valorStr).catch(() => { });

        await page.waitForTimeout(180);
        if (await coincideValor()) break;

        if (i === maxIntentos) {
            const actualFinal = (await input.inputValue().catch(() => "")).trim();
            throw new Error(`No se pudo llenar '${labelText}'. esperado='${valorStr}', actual='${actualFinal}'`);
        }
    }

    if (typeof accionDespues === "function") {
        await accionDespues();
    }

    return true;
}

export async function llenarInputNumberSiVacio(page, labelText, valor, accionDespues = null) {
    const { label, field } = await fieldContainerByLabel(page, labelText);
    const visible = await label.isVisible().catch(() => false);
    if (!visible) return false;

    const input = field
        .locator('input.p-inputnumber-input, input.p-inputtext, input[type="text"], input[type="tel"], input')
        .first();
    const inputVisible = await input.isVisible().catch(() => false);
    if (!inputVisible) return false;

    const actual = (await input.inputValue().catch(() => "")).trim();
    if (actual !== "") return false;

    await llenarInputNumber(page, labelText, valor, accionDespues);
    return true;
}

export async function llenarInputMask(page, labelText, telefono) {
    console.log(`ðŸ“ž Llenando telÃ©fono '${labelText}' â†’ ${telefono}`);

    const label = page.locator(`xpath=//*[normalize-space(text())="${labelText}"]`).first();
    if (!(await label.count())) {
        throw new Error(`âŒ No se encontrÃ³ el label '${labelText}'`);
    }

    const input = label.locator(`xpath=following::input[contains(@class,"p-inputmask")][1]`);
    await input.waitFor({ state: "visible" });

    await input.click({ force: true });
    await page.waitForTimeout(150);
    await input.click({ force: true });
    await page.waitForTimeout(150);
    await page.keyboard.press("Control+A");
    await page.keyboard.press("Delete");
    await page.waitForTimeout(150);
    await page.keyboard.type("8", { delay: 80 });
    await page.waitForTimeout(100);
    await page.keyboard.press("Control+A");
    await page.keyboard.press("Delete");
    await page.waitForTimeout(100);

    const numeroLimpio = telefono.replace(/\D/g, "");
    await input.type(numeroLimpio, { delay: 70 });
    await input.blur();
    await page.dispatchEvent(
        `xpath=//*[normalize-space(text())="${labelText}"]/following::input[1]`,
        "change"
    );

    console.log("âœ” TelÃ©fono llenado correctamente (VERIFICAR VISUAL)");
}

export async function ejecutarSiLabelVisible(page, labelText, action) {
    const label = page.locator(`label:has-text("${labelText}")`).first();
    const visible = await label.isVisible().catch(() => false);
    if (!visible) return;
    await action();
}

export async function llenarCampoPorLabel(page, labelText, value, opts = {}) {
    const timeoutMs = opts.timeoutMs ?? 20000;
    const maxIntentos = opts.maxIntentos ?? 4;
    const clearFirst = opts.clearFirst ?? true;
    const pressEnter = opts.pressEnter ?? false;
    const blur = opts.blur ?? true;
    const log = opts.log ?? true;

    const { label, field } = await fieldContainerByLabel(page, labelText);
    await label.waitFor({ state: 'visible', timeout: timeoutMs });
    await field.waitFor({ state: 'visible', timeout: timeoutMs });
    await field.scrollIntoViewIfNeeded().catch(() => { });

    const { inputNumber, inputText, anyInput, textarea } = inputInField(field);
    let target = inputNumber;
    if ((await target.count().catch(() => 0)) === 0) target = inputText;
    if ((await target.count().catch(() => 0)) === 0) target = anyInput;
    const hasInput = (await target.count().catch(() => 0)) > 0;

    if (!hasInput) {
        if ((await textarea.count().catch(() => 0)) === 0) {
            throw new Error(`âŒ No encontrÃ© input/textarea para el label '${labelText}'.`);
        }
        target = textarea;
    }

    for (let i = 1; i <= maxIntentos; i++) {
        await target.scrollIntoViewIfNeeded().catch(() => { });
        await target.waitFor({ state: 'visible', timeout: timeoutMs });

        const disabled = await target.isDisabled().catch(() => false);
        if (disabled) {
            if (log) console.log(`â†» '${labelText}' estÃ¡ disabled (intento ${i}/${maxIntentos})`);
            if (opts.onDisabled) {
                await opts.onDisabled(page);
            }
            await page.waitForTimeout(300);
            continue;
        }

        await target.waitFor({ state: 'editable', timeout: timeoutMs }).catch(() => { });

        if (clearFirst) {
            await target.click({ force: true });
            await page.keyboard.press('Control+A').catch(() => { });
            await page.keyboard.press('Backspace').catch(() => { });
            await page.waitForTimeout(30);
        }

        const txt = String(value);
        await target.fill(txt);

        if (pressEnter) await page.keyboard.press('Enter').catch(() => { });
        if (blur) await target.blur().catch(() => { });
        if (log) console.log(`âœ… Llenado '${labelText}' â†’ ${txt}`);
        return;
    }

    throw new Error(`âŒ No pude llenar '${labelText}' (sigue disabled o no editable) tras ${maxIntentos} intentos.`);
}

export async function llenarCampoYEnter(page, labelText, value, opts = {}) {
    return llenarCampoPorLabel(page, labelText, value, { ...opts, pressEnter: true });
}

export async function clickBotonPorLabel(page, labelText, opts = {}) {
    const timeoutMs = opts.timeoutMs ?? 15000;
    const { label, field } = await fieldContainerByLabel(page, labelText);

    await label.waitFor({ state: 'visible', timeout: timeoutMs });
    await field.waitFor({ state: 'visible', timeout: timeoutMs });
    await field.scrollIntoViewIfNeeded().catch(() => { });

    const candidates = [];
    if (opts.name) {
        candidates.push(field.getByRole('button', { name: opts.name }).first());
    }
    candidates.push(field.locator('.p-inputgroup button, button').first());
    candidates.push(label.locator('xpath=ancestor::*[self::div or self::span][.//button][1]//button[1]').first());

    for (const btn of candidates) {
        const visible = await btn.waitFor({ state: 'visible', timeout: Math.max(1200, Math.floor(timeoutMs / 3)) })
            .then(() => true)
            .catch(() => false);
        if (!visible) continue;
        await btn.scrollIntoViewIfNeeded().catch(() => { });
        await btn.click({ force: true });
        return;
    }

    throw new Error(`No se encontro un boton clickeable para '${String(labelText)}'.`);
}

function parseDDMMYYYY(s) {
    const [dd, mm, yyyy] = s.split("-").map(Number);
    return new Date(Date.UTC(yyyy, mm - 1, dd));
}

function formatDDMMYYYY(d) {
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const yyyy = d.getUTCFullYear();
    return `${dd}-${mm}-${yyyy}`;
}

function addDaysUTC(d, days) {
    const copy = new Date(d.getTime());
    copy.setUTCDate(copy.getUTCDate() + days);
    return copy;
}

async function setInputValueHard(input, value) {
    await input.click({ force: true });
    await input.press(process.platform === "darwin" ? "Meta+A" : "Control+A");
    await input.press("Backspace");
    await input.type(value, { delay: 20 });
    await input.press("Enter");
    await input.press("Tab");

    const after = (await input.inputValue().catch(() => ""))?.trim();
    if (after === value) return true;

    await input.evaluate((el, v) => {
        el.focus();
        el.value = v;
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
        el.dispatchEvent(new Event("blur", { bubbles: true }));
    }, value);

    const after2 = (await input.inputValue().catch(() => ""))?.trim();
    return after2 === value;
}

function clampDateByValidationMessage(texto, margenDias = 1) {
    const match = String(texto || "").match(/(m[i�]nima|m[a�]xima)\s+permitida\s+es\s*(\d{2}-\d{2}-\d{4})/i);
    if (!match) return null;

    const tipo = match[1].toLowerCase();
    const fechaLimite = parseDDMMYYYY(match[2]);
    if (tipo.includes("min")) {
        return formatDDMMYYYY(addDaysUTC(fechaLimite, margenDias));
    }
    return formatDDMMYYYY(fechaLimite);
}

export async function llenarFechaMinimaYDepurar(page, label = "Fecha de expiraciÃ³n", fechaDeseada, margenDias = 1) {
    const field = page.locator(
        `xpath=//*[normalize-space()="${label}"]/ancestor::div[contains(@class,"space-y-2")][1]`
    );
    const input = field.locator('input').first();
    const errorFecha = field.locator('text=/La fecha m[i�]nima permitida es\\s*\\d{2}-\\d{2}-\\d{4}|La fecha m[a�]xima permitida es\\s*\\d{2}-\\d{2}-\\d{4}/i').first();

    await input.waitFor({ state: "visible", timeout: 8000 });
    await setInputValueHard(input, fechaDeseada);

    const visible = await errorFecha.isVisible({ timeout: 800 }).catch(() => false);
    if (visible) {
        const txt = await errorFecha.innerText().catch(() => "");
        const valid = clampDateByValidationMessage(txt, margenDias);
        if (valid) {
            await setInputValueHard(input, valid);
        }
    }

    await errorFecha.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });

    const btn = page.getByRole("button", { name: "Depurar" });
    await btn.scrollIntoViewIfNeeded();
    await btn.click({ noWaitAfter: true });
}
