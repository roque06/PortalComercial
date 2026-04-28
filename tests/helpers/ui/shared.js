export async function getDropdown(page, labelText) {
    const asPattern = (value) => {
        if (value instanceof RegExp) return value;
        const escaped = String(value)
            .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
            .replace(/\\\?/g, ".");
        return new RegExp(escaped, "i");
    };

    const pattern = asPattern(labelText);
    let label = page.locator("label").filter({ hasText: pattern }).first();

    if ((await label.count().catch(() => 0)) === 0) {
        label = page.getByText(pattern).first();
    }

    if ((await label.count().catch(() => 0)) === 0) {
        throw new Error(`No se encontro el label '${String(labelText)}'.`);
    }

    const container = label.locator(
        'xpath=ancestor::*[self::div or self::span][.//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]'
    ).first();
    const dropdown = container.locator('div.p-dropdown, [data-pc-name="dropdown"]').first();

    if (await dropdown.count().catch(() => 0)) {
        return dropdown;
    }

    return label.locator(
        'xpath=following::div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"][1]'
    ).first();
}

export async function fieldContainerByLabel(page, labelText) {
    const pattern = labelText instanceof RegExp
        ? labelText
        : new RegExp(
            String(labelText).replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\\\?/g, "."),
            "i"
        );

    let label = page.locator("label").filter({ hasText: pattern }).first();
    const labelCount = await label.count().catch(() => 0);
    if (!labelCount) {
        label = page.getByText(pattern).first();
    }

    let field = label.locator(
        'xpath=ancestor::*[self::div or self::span][.//input or .//textarea or .//button or .//div[contains(@class,"p-dropdown") or @data-pc-name="dropdown"]][1]'
    ).first();
    if ((await field.count().catch(() => 0)) === 0) {
        field = label.locator('xpath=ancestor::div[1]').first();
    }

    return { label, field };
}

export function inputInField(field) {
    const inputNumber = field.locator('input.p-inputnumber-input').first();
    const inputText = field.locator('input.p-inputtext').first();
    const anyInput = field.locator('input').first();
    const textarea = field.locator('textarea').first();

    return { inputNumber, inputText, anyInput, textarea };
}
