import path from 'node:path';
import fs from 'node:fs';

export async function validarGestionDocumentalSiRequerido(page: import("@playwright/test").Page) {
  const evidenciaPath = path.resolve(process.cwd(), 'data', 'Documentos', 'evidencia.pdf');
  const btnGestion = page.getByRole('button', { name: /Gesti(?:o|\u00f3)n Documental/i }).first();
  const visibleBtn = await btnGestion.isVisible().catch(() => false);
  if (!visibleBtn) return false;

  // Si ya se finalizó la solicitud, no intentar gestión documental
  const modalFinalizada = page.locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Solicitud finalizada|solicitud ha sido finalizada/i }).first();
  if (await modalFinalizada.isVisible().catch(() => false)) return false;

  const badge = btnGestion.locator('[data-pc-name="badge"], .p-badge, span[class*="badge"]').first();
  const badgeVisible = await badge.isVisible().catch(() => false);
  const badgeText = badgeVisible ? ((await badge.textContent().catch(() => "")) || "").trim() : "";
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
    console.log(`[GestionDoc] Carga documental no completada: ${String(e)}`);
  });

  return true;
}

export async function cargarDocumentoEnGestionDocumental(page: import("@playwright/test").Page, filePath: string) {
  console.log(`[GestionDoc] Iniciando carga: ${filePath}`);
  if (!fs.existsSync(filePath)) {
    throw new Error(`[GestionDoc] No existe archivo para adjuntar: ${filePath}`);
  }

  const modalDoc = page
    .locator('.p-dialog:visible, [role="dialog"]:visible')
    .filter({ hasText: /Gesti(?:o|\u00f3)n documental/i })
    .first();
  await modalDoc.waitFor({ state: 'visible', timeout: 15000 });

  const panelDoc = modalDoc.locator('.p-panel-content, [data-pc-section="content"]').first();
  await panelDoc.waitFor({ state: 'visible', timeout: 10000 }).catch(() => { });
  const docScope = modalDoc;
  let envioIniciado = false;

  const obtenerPendientesModal = async () => {
    const badgeModal = modalDoc.locator('[data-pc-name="badge"], .p-badge').first();
    const visible = await badgeModal.isVisible().catch(() => false);
    if (!visible) return 0;
    const badgeTxt = ((await badgeModal.textContent().catch(() => "")) || "").trim();
    const match = badgeTxt.match(/\d+/);
    if (!match) return 0;
    const n = Number.parseInt(match[0], 10);
    return Number.isFinite(n) ? n : 0;
  };
  const pendientesObjetivo = await obtenerPendientesModal().catch(() => 0);

  const contarBotonesUploadPendienteVisibles = async () =>
    modalDoc
      .locator('button:visible:has(i.pi-upload), button:visible:has(.pi-upload), button:visible:has([class*="pi-upload"]), button:visible:has([class*="ph-upload"]), button:visible:has([class*="ph--upload"]), button:visible[aria-label*="cargar" i], button:visible[aria-label*="subir" i], button:visible[aria-label*="Cargar" i], button[aria-label*="adjuntar" i]:visible')
      .count()
      .catch(() => 0);

  const contarBotonesEyeVisibles = async () =>
    modalDoc
      .locator('button.p-button-success:visible, button:visible:has(i.pi-eye), button:visible:has(.pi-eye), button:visible:has([class*="pi-eye"]), button:visible:has([class*="ph--eye"]), button:visible:has([class*="icon-[ph--eye]"])')
      .count()
      .catch(() => 0);

  const hayDocumentoEnCargados = async () => {
    const result = await page.evaluate(() => {
      const botones = Array.from(document.querySelectorAll('button'));
      let tieneBotonEye = false;
      const clasesEye: string[] = [];
      for (const btn of botones) {
        const todasClases = [btn.className, ...Array.from(btn.querySelectorAll('[class]')).map(e => e.className)].join(' ');
        if (/eye/i.test(todasClases)) {
          const rect = btn.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) {
            tieneBotonEye = true;
            clasesEye.push(todasClases.replace(/\s+/g, ' ').substring(0, 120));
          }
        }
      }
      const innerText = (document.body as HTMLElement).innerText || '';
      const cargadosIdx = innerText.search(/cargados/i);
      const lastCargadosIdx = innerText.toLowerCase().lastIndexOf('cargados');
      const trasTras = lastCargadosIdx >= 0 ? innerText.substring(lastCargadosIdx + 8).trim() : '';
      return {
        tieneBotonEye,
        clasesEye: clasesEye.slice(0, 2),
        innerTextLen: innerText.length,
        primeraOcurrencia: cargadosIdx,
        ultimaOcurrencia: lastCargadosIdx,
        trasCargados: trasTras.substring(0, 60),
      };
    }).catch(() => null);

    if (!result) {
      console.log('[GestionDoc][Debug] hayDocumentoEnCargados: evaluate fallo -> false');
      return false;
    }
    const hayContenido = result.trasCargados.length > 3;
    console.log(
      `[GestionDoc][Debug] hayDocumentoEnCargados: eye=${result.tieneBotonEye}, ` +
      `innerLen=${result.innerTextLen}, cargados1=${result.primeraOcurrencia}, ` +
      `cargadosN=${result.ultimaOcurrencia}, trasCargados="${result.trasCargados}" -> ${result.tieneBotonEye || hayContenido}` +
      (result.clasesEye.length ? ` | clases=[${result.clasesEye.join(' | ')}]` : '')
    );
    return result.tieneBotonEye || hayContenido;
  };

  const modalDocumentalCompleto = async () => {
    // Solo marcar como completo cuando el estado REAL del modal confirma 0 pendientes
    // y 0 botones de carga visibles. No confiar en `pendientesObjetivo` porque el
    // servidor puede generar documentos adicionales despues de que abrimos el modal
    // (p.ej. FORMULARIO CONOZCA SU CLIENTE aparece tras subir los primeros 2).
    const pendientesActuales = await obtenerPendientesModal().catch(() => 0);
    const uploadsActuales = await contarBotonesUploadPendienteVisibles().catch(() => 0);
    if (pendientesActuales > 0) return false;
    if (uploadsActuales > 0) return false;

    const botonesEye = await contarBotonesEyeVisibles();
    const textoModal = ((await modalDoc.innerText().catch(() => "")) || "").trim();
    const esperandoConfirmacion = /Esperando confirmaci[oÃ³]n/i.test(textoModal);
    if (esperandoConfirmacion) return false;

    const pendientesMatch = textoModal.match(/Pendientes\s*([\s\S]*?)Cargados/i);
    const cargadosMatch = textoModal.match(/Cargados\s*([\s\S]*)$/i);
    const pendientesTexto = (pendientesMatch?.[1] ?? "").replace(/\s+/g, " ").trim();
    const cargadosTexto = (cargadosMatch?.[1] ?? "").replace(/\s+/g, " ").trim();
    return botonesEye > 0 && pendientesTexto.length === 0 && cargadosTexto.length > 3;
  };

  // Espera a que aparezca al menos un botón de upload o el badge llegue a 0.
  // Se usa cuando badge>0 pero aún no se ha renderizado el botón del siguiente doc.
  const esperarBotonUploadOBadgeCero = async (timeoutMs = 60000) => {
    const inicio = Date.now();
    while (Date.now() - inicio < timeoutMs) {
      const uploads = await contarBotonesUploadPendienteVisibles().catch(() => 0);
      if (uploads > 0) return true;
      const pendientes = await obtenerPendientesModal().catch(() => 0);
      if (pendientes <= 0) return true;
      await page.waitForTimeout(1000);
    }
    return false;
  };

  const esperarCambioPendientes = async (pendientesAntes: number, uploadsAntes: number, timeoutMs = 12000) => {
    const inicio = Date.now();
    while (Date.now() - inicio < timeoutMs) {
      const enviandoVisible = await modalDoc.getByText(/Enviando documento/i).first().isVisible().catch(() => false);
      if (enviandoVisible) {
        envioIniciado = true;
        return true;
      }

      const actual = await obtenerPendientesModal();
      const uploadsActuales = await contarBotonesUploadPendienteVisibles();
      if (actual < pendientesAntes) return true;
      if (uploadsActuales < uploadsAntes) return true;
      if (actual <= 0 && uploadsActuales <= 0) return true;
      await page.waitForTimeout(200);
    }
    return false;
  };

  const esperarEstadoDocumentalEstable = async (timeoutMs = 20000) => {
    const inicio = Date.now();
    let pendientes = await obtenerPendientesModal().catch(() => 0);
    let uploads = await contarBotonesUploadPendienteVisibles().catch(() => 0);
    let completo = await modalDocumentalCompleto().catch(() => false);
    let hayCargados = await hayDocumentoEnCargados().catch(() => false);
    let esperandoConfirmacion = await modalDoc.getByText(/Esperando confirmaci[oó]n/i).first().isVisible().catch(() => false);

    while (Date.now() - inicio < timeoutMs) {
      pendientes = await obtenerPendientesModal().catch(() => pendientes);
      uploads = await contarBotonesUploadPendienteVisibles().catch(() => uploads);
      completo = await modalDocumentalCompleto().catch(() => completo);
      hayCargados = await hayDocumentoEnCargados().catch(() => hayCargados);
      esperandoConfirmacion = await modalDoc.getByText(/Esperando confirmaci[oó]n/i).first().isVisible().catch(() => false);

      if (completo) break;
      if (pendientes <= 0 && uploads <= 0) break;  // badge=0 sin uploads pendientes → terminado
      if (uploads > 0) break;                       // hay boton disponible → salir a subirlo
      if (pendientes <= 0 && !esperandoConfirmacion) break;

      await page.waitForTimeout(500);
    }

    return { pendientes, uploads, esperandoConfirmacion, completo, hayCargados };
  };

  const cerrarModalPrincipal = async () => {
    const btnCerrarModal = modalDoc.locator(
      '[data-pc-section="closebutton"], [data-pc-group-section="headericon"], button.p-dialog-header-close, button[aria-label="Cerrar"], button[aria-label="Close"], button:has-text("×")'
    ).first();

    for (let intento = 1; intento <= 4; intento++) {
      const modalVisibleAntes = await modalDoc.isVisible().catch(() => false);
      const cerrarVisible = await btnCerrarModal.isVisible().catch(() => false);
      console.log(`[GestionDoc][Close] intento=${intento} modalVisible=${modalVisibleAntes} cerrarVisible=${cerrarVisible}`);
      if (!modalVisibleAntes) return;

      if (cerrarVisible) {
        await btnCerrarModal.scrollIntoViewIfNeeded().catch(() => { });
        const clicked = await btnCerrarModal.click({ force: true, timeout: 2500 }).then(() => true).catch(() => false);
        if (!clicked) {
          const handle = await btnCerrarModal.elementHandle().catch(() => null);
          if (handle) {
            await page.evaluate((el) => (el as HTMLElement).click(), handle).catch(() => { });
          }
        }
      } else {
        await page.keyboard.press('Escape').catch(() => { });
      }

      const hidden = await modalDoc.waitFor({ state: 'hidden', timeout: 2500 }).then(() => true).catch(() => false);
      console.log(`[GestionDoc][Close] intento=${intento} hidden=${hidden}`);
      if (hidden) return;

      await page.keyboard.press('Escape').catch(() => { });
      const hiddenTrasEscape = await modalDoc.waitFor({ state: 'hidden', timeout: 1500 }).then(() => true).catch(() => false);
      console.log(`[GestionDoc][Close] intento=${intento} hiddenTrasEscape=${hiddenTrasEscape}`);
      if (hiddenTrasEscape) return;

      await page.waitForTimeout(400);
    }

    const sigueVisible = await modalDoc.isVisible().catch(() => false);
    console.log(`[GestionDoc][Close][WARN] modal sigue visible tras reintentos=${sigueVisible}`);
  };

  const intentarSetArchivoEnInputs = async (
    scope: import("@playwright/test").Locator,
    pendientesAntes: number,
    uploadsAntes: number,
    scopeNombre: string
  ) => {
    const inputs = scope.locator('input[type="file"]');
    const total = await inputs.count().catch(() => 0);
    if (!total) return false;

    for (let i = 0; i < total; i++) {
      const input = inputs.nth(i);
      const disabled = await input.isDisabled().catch(() => false);
      if (disabled) continue;

      const setOk = await input.setInputFiles(filePath).then(() => true).catch(() => false);
      if (!setOk) continue;

      await page.waitForTimeout(250);
      console.log(`[GestionDoc] Archivo asignado via input file (${scopeNombre}#${i + 1}).`);
      return true;
    }
    return false;
  };

  const obtenerBotonesUploadPendiente = () =>
    modalDoc.locator(
      [
        'button:has(i.pi-upload)',
        'button:has(.pi-upload)',
        'button:has([class*="pi-upload"])',
        'button:has([class*="ph-upload"])',
        'button:has([class*="ph--upload"])',
        'button[aria-label*="cargar" i]',
        'button[aria-label*="subir" i]',
        'button[aria-label*="adjuntar" i]'
      ].join(', ')
    );

  const hayArchivoListoEnDialogoCarga = async (dialogCarga: import("@playwright/test").Locator) => {
    const nombreArchivo = dialogCarga.getByText(/\.pdf$|evidencia\.pdf|cedula/i).first();
    const iconoEliminar = dialogCarga.locator(
      'button:visible:has(i.pi-trash), button:visible:has(.pi-trash), [class*="pi-trash"]'
    ).first();
    const iconoVer = dialogCarga.locator(
      'button:visible:has(i.pi-eye), button:visible:has(.pi-eye), [class*="pi-eye"]'
    ).first();

    const nombreVisible = await nombreArchivo.isVisible().catch(() => false);
    const trashVisible = await iconoEliminar.isVisible().catch(() => false);
    const eyeVisible = await iconoVer.isVisible().catch(() => false);
    return nombreVisible || trashVisible || eyeVisible;
  };

  const esperarArchivoListoEnDialogoCarga = async (dialogCarga: import("@playwright/test").Locator, timeoutMs = 8000) => {
    const inicio = Date.now();
    while (Date.now() - inicio < timeoutMs) {
      const listo = await hayArchivoListoEnDialogoCarga(dialogCarga).catch(() => false);
      if (listo) return true;
      await page.waitForTimeout(250);
    }
    return false;
  };

  const confirmarCargaEnDialogo = async (
    dialogCarga: import("@playwright/test").Locator,
    pendientesAntes: number,
    uploadsAntes: number
  ) => {
    const btnCargarPreciso = dialogCarga.locator(
      'button[aria-label="Cargar"][data-pc-severity="primary"], button[aria-label="Cargar"].p-button-primary'
    ).first();
    const btnCargarFallback = dialogCarga.getByRole('button', { name: /^Cargar$/i }).last();

    const precisoVisible = await btnCargarPreciso.isVisible().catch(() => false);
    const btnCargarFinal = precisoVisible ? btnCargarPreciso : btnCargarFallback;
    const visible = await btnCargarFinal.isVisible().catch(() => false);
    if (!visible) return false;

    await btnCargarFinal.scrollIntoViewIfNeeded().catch(() => { });

    let clicked = await btnCargarFinal.click({ timeout: 2500 }).then(() => true).catch(() => false);
    if (!clicked) {
      clicked = await btnCargarFinal.click({ force: true, timeout: 2500 }).then(() => true).catch(() => false);
    }

    if (!clicked) {
      const box = await btnCargarFinal.boundingBox().catch(() => null);
      if (box) {
        await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
        clicked = true;
      }
    }

    if (!clicked) return false;

    console.log('[GestionDoc] Click en "Cargar" final del popup de carga.');

    const modalCerro = await dialogCarga.waitFor({ state: 'hidden', timeout: 10000 })
      .then(() => true)
      .catch(() => false);
    if (modalCerro) {
      console.log('[GestionDoc] Modal secundario "Cargar documento" cerro tras click en Cargar.');
    }

    const cambioPendientes = await esperarCambioPendientes(pendientesAntes, uploadsAntes, 12000).catch(() => false);
    const hayCargado = await hayDocumentoEnCargados().catch(() => false);
    return Boolean(modalCerro || cambioPendientes || hayCargado || envioIniciado);
  };

  const inputsInline = await intentarSetArchivoEnInputs(docScope, 0, 0, 'modalDoc').catch(() => false);
  if (inputsInline) {
    console.log('[GestionDoc] Archivo asignado directamente al input visible del modal.');
    return;
  }

  let pendientesAntes = await obtenerPendientesModal().catch(() => 0);
  let uploadsAntes = await contarBotonesUploadPendienteVisibles().catch(() => 0);
  const botonesUpload = obtenerBotonesUploadPendiente();
  const totalUpload = await botonesUpload.count().catch(() => 0);
  if (!totalUpload) {
    throw new Error('[GestionDoc] No se encontraron botones de carga para documentos pendientes.');
  }

  for (let i = 0; i < await botonesUpload.count().catch(() => 0); i++) {
    const btnUpload = botonesUpload.nth(i);
    const visible = await btnUpload.isVisible().catch(() => false);
    if (!visible) continue;

    const chooserPromise = page.waitForEvent('filechooser', { timeout: 2500 }).catch(() => null);
    await btnUpload.scrollIntoViewIfNeeded().catch(() => { });
    await btnUpload.click({ force: true }).catch(() => { });

    const chooser = await chooserPromise;
    if (chooser) {
      await chooser.setFiles(filePath).catch(() => { });
    }

    const dialogCarga = page
      .locator('.p-dialog:visible, [role="dialog"]:visible')
      .filter({ hasText: /Cargar documento|cargar archivo|seleccionar archivo/i })
      .last();

    const dialogVisible = await dialogCarga.waitFor({ state: 'visible', timeout: 5000 }).then(() => true).catch(() => false);
    if (!dialogVisible) {
      const setViaPageInput = await intentarSetArchivoEnInputs(page.locator('body'), pendientesAntes, uploadsAntes, 'body').catch(() => false);
      if (setViaPageInput) {
        await page.waitForTimeout(500);
        pendientesAntes = await obtenerPendientesModal().catch(() => pendientesAntes);
        uploadsAntes = await contarBotonesUploadPendienteVisibles().catch(() => uploadsAntes);
        continue;
      }
      continue;
    }

    const inputsDialog = await intentarSetArchivoEnInputs(dialogCarga, pendientesAntes, uploadsAntes, 'dialogCarga').catch(() => false);
    if (!inputsDialog) {
      const btnUploadDialog = dialogCarga.locator(
          'button:has(i.pi-upload), button:has(.pi-upload), button:has([class*="pi-upload"]), button:has([class*="ph-upload"]), button:has([class*="ph--upload"]), button[aria-label*="cargar" i], button[aria-label*="subir" i], button[aria-label*="adjuntar" i], button.p-button-primary'
      ).first();
      if (await btnUploadDialog.isVisible().catch(() => false)) {
          const chooserDialogPromise = page.waitForEvent('filechooser', { timeout: 3500 }).catch(() => null);
          await btnUploadDialog.scrollIntoViewIfNeeded().catch(() => {});
          await btnUploadDialog.click({ force: true }).catch(() => {});
          
          const chooserDialog = await chooserDialogPromise;
          if (chooserDialog) {
            await chooserDialog.setFiles(filePath).catch(() => { });
            console.log('[GestionDoc] Archivo provisto via filechooser en el dialog de carga.');
          }
      } else {
          // If no specific button, just wait to see if it opened anyway
          const chooserDialog = await page.waitForEvent('filechooser', { timeout: 1200 }).catch(() => null);
          if (chooserDialog) {
            await chooserDialog.setFiles(filePath).catch(() => { });
          }
      }
    }

    await esperarArchivoListoEnDialogoCarga(dialogCarga, 4000).catch(() => false);
    
    // If the modal changed and requires a secondary confirmation:
    // But since "Cargar" was already clicked to open the filechooser, we check if there's a different button now or if we just wait.
    let confirmo = false;
    const dialogSigueVisible = await dialogCarga.isVisible().catch(() => false);
    if (dialogSigueVisible) {
        const btnListoCargar = dialogCarga.locator('button[aria-label*="Cargar"], button:has-text("Cargar")').last();
        // If the file is ready in the DOM (e.g. name is shown), clicking Cargar again might be needed
        const fileInDOM = await hayArchivoListoEnDialogoCarga(dialogCarga).catch(() => false);
        if (fileInDOM && await btnListoCargar.isVisible().catch(() => false)) {
            confirmo = await confirmarCargaEnDialogo(dialogCarga, pendientesAntes, uploadsAntes).catch(() => false);
        } else {
             // Let it process automatically
             const modalCerro = await dialogCarga.waitFor({ state: 'hidden', timeout: 10000 }).then(() => true).catch(() => false);
             confirmo = modalCerro || await esperarCambioPendientes(pendientesAntes, uploadsAntes, 12000).catch(() => false);
        }
    } else {
        confirmo = true;
    }
    
    if (confirmo) {
      console.log('[GestionDoc] Documento cargado correctamente.');
      await page.waitForTimeout(800);
      console.log('[GestionDoc][PostCarga] leyendo estado inmediato...');
      const botonesEyeActuales = await contarBotonesEyeVisibles().catch(() => 0);
      const uploadsInmediatos = await contarBotonesUploadPendienteVisibles().catch(() => 0);
      const pendientesInmediatos = await obtenerPendientesModal().catch(() => 0);
      const esperandoTrasCarga = await modalDoc.getByText(/Esperando confirmaci[oÃ³]n/i).first().isVisible().catch(() => false);
      console.log(`[GestionDoc][PostCarga] objetivoInicial=${pendientesObjetivo} eyes=${botonesEyeActuales} badge=${pendientesInmediatos} uploads=${uploadsInmediatos} esperando=${esperandoTrasCarga}`);
      // Nunca cerrar basado en `pendientesObjetivo` inicial: el servidor puede
      // generar documentos nuevos despues. Requerir estado real: 0 pendientes + 0 uploads.
      console.log('[GestionDoc][PostCarga] Entrando a espera estable para confirmar estado real...');
      const estadoEstable = await esperarEstadoDocumentalEstable(45000);
      pendientesAntes = estadoEstable.pendientes;
      uploadsAntes = estadoEstable.uploads;
      console.log(
        `[GestionDoc] Pendientes restantes tras carga: uploads=${uploadsAntes} badge=${pendientesAntes} esperando=${estadoEstable.esperandoConfirmacion} completo=${estadoEstable.completo}`
      );
      // Solo cerrar cuando realmente no quedan pendientes ni uploads visibles.
      if (uploadsAntes <= 0 && pendientesAntes <= 0 && estadoEstable.completo) {
        await cerrarModalPrincipal();
        console.log('[GestionDoc] Modal principal cerrado. Todos los documentos cargados.');
        return;
      }
      if (uploadsAntes <= 0 && pendientesAntes > 0) {
        console.log('[GestionDoc] Aun hay pendientes sin boton visible; esperando que se genere el siguiente documento...');
        await esperarBotonUploadOBadgeCero(90000);
      }
      // Reiniciar el loop desde el primer botón (CDK Portal re-renderiza la lista)
      i = -1; // i++ lo lleva a 0 en la siguiente iteración
      continue;
    }

    pendientesAntes = await obtenerPendientesModal().catch(() => pendientesAntes);
    uploadsAntes = await contarBotonesUploadPendienteVisibles().catch(() => uploadsAntes);
  }

  let uploadsFinales = await contarBotonesUploadPendienteVisibles().catch(() => 0);
  let pendientesFinales = await obtenerPendientesModal().catch(() => 0);

  // Si quedan pendientes en el badge pero ya no hay botones visibles, dar otra
  // oportunidad al servidor a que genere el siguiente documento antes de
  // cerrar o fallar. Esto cubre el caso de FORMULARIO CONOZCA SU CLIENTE
  // que aparece despues de subir los primeros documentos.
  if (uploadsFinales <= 0 && pendientesFinales > 0) {
    console.log('[GestionDoc][EndLoop] badge>0 sin uploads visibles. Esperando generacion del siguiente documento antes de decidir cierre...');
    await esperarBotonUploadOBadgeCero(90000);
    uploadsFinales = await contarBotonesUploadPendienteVisibles().catch(() => 0);
    pendientesFinales = await obtenerPendientesModal().catch(() => 0);

    // Si aparecio un nuevo boton de upload, subir el archivo y volver a esperar estado estable.
    if (uploadsFinales > 0) {
      console.log('[GestionDoc][EndLoop] boton de upload aparecio; reintentando carga del documento restante.');
      const btnUploadNuevo = obtenerBotonesUploadPendiente().first();
      const visibleNuevo = await btnUploadNuevo.isVisible().catch(() => false);
      if (visibleNuevo) {
        const chooserPromise = page.waitForEvent('filechooser', { timeout: 2500 }).catch(() => null);
        await btnUploadNuevo.scrollIntoViewIfNeeded().catch(() => {});
        await btnUploadNuevo.click({ force: true }).catch(() => {});
        const chooser = await chooserPromise;
        if (chooser) await chooser.setFiles(filePath).catch(() => {});

        const dialogCargaNuevo = page
          .locator('.p-dialog:visible, [role="dialog"]:visible')
          .filter({ hasText: /Cargar documento|cargar archivo|seleccionar archivo/i })
          .last();
        const dialogVisibleNuevo = await dialogCargaNuevo.waitFor({ state: 'visible', timeout: 5000 }).then(() => true).catch(() => false);
        if (dialogVisibleNuevo) {
          await intentarSetArchivoEnInputs(dialogCargaNuevo, pendientesFinales, uploadsFinales, 'dialogCargaEndLoop').catch(() => false);
          await esperarArchivoListoEnDialogoCarga(dialogCargaNuevo, 4000).catch(() => false);
          await confirmarCargaEnDialogo(dialogCargaNuevo, pendientesFinales, uploadsFinales).catch(() => false);
        }
        await esperarEstadoDocumentalEstable(45000).catch(() => null);
        uploadsFinales = await contarBotonesUploadPendienteVisibles().catch(() => 0);
        pendientesFinales = await obtenerPendientesModal().catch(() => 0);
      }
    }
  }

  const hayCargadosFinal = await hayDocumentoEnCargados().catch(() => false);
  const completoFinal = await modalDocumentalCompleto().catch(() => false);
  if (completoFinal || (uploadsFinales <= 0 && pendientesFinales <= 0)) {
    await cerrarModalPrincipal();
    console.log('[GestionDoc] Modal principal cerrado al agotar el loop. Todos los documentos quedaron cargados.');
    return;
  }

  // Si aun asi quedan pendientes visibles, NO cerrar prematuramente: lanzar error para que
  // el retry superior (finalizarSolicitudSiPendiente) reintente desde cero.
  if (pendientesFinales > 0) {
    throw new Error(`[GestionDoc] Quedaron ${pendientesFinales} documentos pendientes tras espera extendida.`);
  }

  // Ultimo recurso: badge=0 con docs en Cargados pero completo=false (anomalia), cerrar.
  if (hayCargadosFinal) {
    console.log(`[GestionDoc][WARN] Cerrando modal con badge=${pendientesFinales} y docs en Cargados.`);
    await cerrarModalPrincipal();
    return;
  }

  throw new Error('[GestionDoc] No se pudo completar la carga documental.');
}
