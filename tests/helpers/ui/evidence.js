import fs from "node:fs";
import path from "node:path";
import { PDFDocument } from "pdf-lib";

function sanitizeFileName(name) {
    return String(name ?? "")
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[\\/:*?"<>|]/g, "-")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 160);
}

function resolvePathFromCwd(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";
    return path.isAbsolute(raw) ? raw : path.resolve(process.cwd(), raw);
}

async function esperarPantallaListaParaCaptura(page, opts = {}) {
    const timeoutMs = opts.timeoutMs ?? 60000;
    const stableMs = opts.stableMs ?? 600;
    const inicio = Date.now();

    if (!page || page.isClosed?.()) return false;

    const overlays = page.locator(
        [
            ".p-blockui:visible",
            '[data-pc-name="blockui"]:visible',
            ".p-progressspinner:visible",
            ".p-progress-spinner:visible",
            '[role="progressbar"]:visible',
        ].join(", ")
    );

    const mensajes = [
        page.getByText(/Actualizando solicitud/i).first(),
        page.getByText(/Depurando solicitante|Consultando datos del solicitante/i).first(),
        page.getByText(/Obteniendo datos de certificado/i).first(),
    ];

    const hayMensajesLoadingVisibles = async () => {
        for (const m of mensajes) {
            if (await m.isVisible().catch(() => false)) return true;
        }
        return false;
    };

    while (Date.now() - inicio < timeoutMs) {
        const overlayCount = await overlays.count().catch(() => 0);
        const mensajeVisible = await hayMensajesLoadingVisibles().catch(() => false);

        if (overlayCount === 0 && !mensajeVisible) {
            await page.waitForTimeout(stableMs).catch(() => { });
            const overlayCount2 = await overlays.count().catch(() => 0);
            const mensajeVisible2 = await hayMensajesLoadingVisibles().catch(() => false);
            if (overlayCount2 === 0 && !mensajeVisible2) return true;
        }

        await page.waitForTimeout(250).catch(() => { });
    }

    return false;
}

export async function capturarCuentaComoPNG(page, identificacion, tipoCuenta, opts = {}) {
    const dirFromEnv = resolvePathFromCwd(process.env.PW_EVIDENCIAS_DIR);
    const dir = dirFromEnv || path.resolve(process.cwd(), "artifacts", "evidencias_tmp");
    fs.mkdirSync(dir, { recursive: true });

    const fileName = `${sanitizeFileName(identificacion)}__${sanitizeFileName(tipoCuenta)}.png`;
    const filePath = path.join(dir, fileName);

    const listo = opts.skipSpinnerWait ? true : await esperarPantallaListaParaCaptura(page).catch(() => false);
    if (!listo) {
        console.log("⚠️ Screenshot: timeout esperando spinner/overlay. Se captura de todos modos.");
    }
    await page.screenshot({ path: filePath, fullPage: true });
    console.log("ðŸ–¼ï¸ Screenshot guardado:", filePath);

    return filePath;
}

export async function unirPNGsEnUnPDF(pngPaths, outPdfName = "Evidencias_Cuentas.pdf") {
    const outDirFromEnv = resolvePathFromCwd(process.env.PW_EVIDENCIAS_PDF_DIR);
    const outDir = outDirFromEnv || path.resolve(process.cwd(), "artifacts", "evidencias");
    fs.mkdirSync(outDir, { recursive: true });

    const safeName = sanitizeFileName(outPdfName).toLowerCase().endsWith(".pdf")
        ? sanitizeFileName(outPdfName)
        : `${sanitizeFileName(outPdfName)}.pdf`;

    const outPath = path.join(outDir, safeName);
    const pdfDoc = await PDFDocument.create();

    for (const pngPath of pngPaths) {
        const bytes = fs.readFileSync(pngPath);
        const img = await pdfDoc.embedPng(bytes);
        const { width, height } = img.scale(1);
        const pdfPage = pdfDoc.addPage([width, height]);
        pdfPage.drawImage(img, { x: 0, y: 0, width, height });
    }

    fs.writeFileSync(outPath, await pdfDoc.save());
    console.log("ðŸ“„ PDF final generado:", outPath);

    return outPath;
}
