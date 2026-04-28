import { chromium } from '@playwright/test';

(async () => {
    // Abrimos navegador real donde haremos login
    const browser = await chromium.launch({
        headless: false,
    });

    const context = await browser.newContext({
        ignoreHTTPSErrors: true
    });

    const page = await context.newPage();

    console.log("➡️ Abriendo portal...");
    await page.goto('https://srvqacgowb01.local.bsc.com:5000');

    console.log("➡️ Haz login manualmente y aprueba la MFA.");
    console.log("➡️ Cuando ya estés dentro del portal, CIERRA esta ventana del navegador.");

    // Espera que cierres la ventana
    page.on('close', async () => {
        console.log("➡️ Guardando sesión en auth.json …");
        await context.storageState({ path: 'auth.json' });
        console.log("✔ Sesión guardada exitosamente.");
        await browser.close();
    });
})();




