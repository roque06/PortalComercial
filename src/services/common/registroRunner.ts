import path from 'node:path';
import type { Page } from '@playwright/test';
import type { BrowserSession, RegistroErrorInfo, RegistroSkipInfo } from '../../core/types/runner';

export interface RunRegistrosOptions<TRegistro> {
  page: Page;
  registros: TRegistro[];
  browserSession: BrowserSession;
  onRegistro: (registro: TRegistro) => Promise<void>;
  onRegistroError?: (
    registro: TRegistro,
    errorMsg: string,
    state: { failedRegistros: RegistroErrorInfo[]; skippedRegistros: RegistroSkipInfo[] },
  ) => Promise<'continue' | 'throw'>;
  onRegistroFinally?: (registro: TRegistro) => Promise<void>;
}

function getIdentificacion(registro: unknown) {
  if (!registro || typeof registro !== 'object') return 'sin-identificacion';
  return String((registro as { identificacion?: string }).identificacion ?? 'sin-identificacion');
}

async function takeErrorScreenshot(page: Page, identificacion: string) {
  if (page.isClosed()) return;

  await page
    .screenshot({
      path: path.resolve(
        process.cwd(),
        'artifacts',
        'evidencias_tmp',
        `error-${identificacion.replace(/[^\d-]/g, '')}.png`,
      ),
      fullPage: true,
    })
    .catch(() => {});
}

export async function runRegistros<TRegistro>(
  options: RunRegistrosOptions<TRegistro>,
): Promise<{ failedRegistros: RegistroErrorInfo[]; skippedRegistros: RegistroSkipInfo[] }> {
  const { page, registros, browserSession, onRegistro, onRegistroError, onRegistroFinally } = options;
  const failedRegistros: RegistroErrorInfo[] = [];
  const skippedRegistros: RegistroSkipInfo[] = [];

  for (const registro of registros) {
    const identificacion = getIdentificacion(registro);
    console.log(`\n\nINICIANDO PROCESO PARA: ${identificacion}\n`);

    try {
      await onRegistro(registro);
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      failedRegistros.push({ identificacion, error: errorMsg });
      console.log(`[ERROR][${identificacion}] ${errorMsg}`);
      await takeErrorScreenshot(page, identificacion);

      if (browserSession.pauseOnRegistroError) {
        await page.pause().catch(() => {});
      }

      if (onRegistroError) {
        const action = await onRegistroError(registro, errorMsg, { failedRegistros, skippedRegistros });
        if (action === 'continue') continue;
      }

      throw error;
    } finally {
      if (onRegistroFinally) {
        console.log(`[Loop][${identificacion}] Cleanup de fin de registro.`);
        await onRegistroFinally(registro);
      }
    }
  }

  return { failedRegistros, skippedRegistros };
}

export async function closeBrowserSession(session: BrowserSession, failedRegistros: RegistroErrorInfo[]) {
  if (failedRegistros.length && session.keepBrowserOpenOnFailures) {
    console.log('Se detectaron errores. El navegador queda abierto (PW_KEEP_BROWSER_ON_FAIL=1).');
    if (session.holdOnFailure) {
      const paused = await session.page.pause().then(() => true).catch(() => false);
      if (!paused) {
        await session.page.waitForTimeout(300000).catch(() => {});
      }
    }
    return;
  }

  await session.context.close();
}
