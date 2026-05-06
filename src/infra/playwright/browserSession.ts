import { chromium } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import type { BrowserSession, BrowserChannel } from '../../core/types/runner';

export interface LaunchPortalSessionOptions {
  browserChannel?: BrowserChannel;
  defaultContinueOnError?: boolean;
  extraArgs?: string[];
  initialUrl?: string;
  preferNewPage?: boolean;
}

function resolveProfileDirectory(profileRoot: string, configuredProfileDirectory: string) {
  let profileDirectory = configuredProfileDirectory;
  const profileConfiguredPath = path.join(profileRoot, configuredProfileDirectory);

  if (!fs.existsSync(profileConfiguredPath)) {
    const availableProfiles = fs.existsSync(profileRoot)
      ? fs
        .readdirSync(profileRoot, { withFileTypes: true })
        .filter((entry) => entry.isDirectory() && /^(Default|Profile \d+)$/i.test(entry.name))
        .map((entry) => entry.name)
      : [];

    if (availableProfiles.length > 0) {
      profileDirectory = availableProfiles.includes('Default')
        ? 'Default'
        : availableProfiles.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0];
      console.log(
        `[Sesion] Perfil solicitado '${configuredProfileDirectory}' no existe en '${profileRoot}'. ` +
        `Usando '${profileDirectory}'.`
      );
    }
  }

  return profileDirectory;
}

async function maximizeWindow(page: BrowserSession['page']) {
  const cdp = await page.context().newCDPSession(page);
  const { windowId } = await cdp.send('Browser.getWindowForTarget');

  try {
    await cdp.send('Browser.setWindowBounds', {
      windowId,
      bounds: { windowState: 'maximized' },
    });
  } catch {
    await cdp.send('Browser.setWindowBounds', {
      windowId,
      bounds: { windowState: 'normal' },
    });
    await cdp.send('Browser.setWindowBounds', {
      windowId,
      bounds: { windowState: 'maximized' },
    });
  }
}

export async function launchPortalSession(
  options: LaunchPortalSessionOptions = {},
): Promise<BrowserSession> {
  const defaultRobotEdgeUserData = path.resolve(
    process.cwd(),
    'artifacts',
    'profiles',
    'profile-edge-robot-temp',
  );
  const profileSuffix = process.env.PW_EDGE_PROFILE_SUFFIX ?? '';
  const profileRoot = process.env.PW_PROFILE_DIR ?? (profileSuffix ? path.resolve(process.cwd(), 'artifacts', 'browser-profiles', `profile-edge-robot${profileSuffix}`) : defaultRobotEdgeUserData);
  const configuredProfileDirectory = process.env.PW_CHROME_PROFILE ?? 'Default';
  const profileDirectory = resolveProfileDirectory(profileRoot, configuredProfileDirectory);
  const browserChannel =
    options.browserChannel ??
    ((process.env.PW_BROWSER_CHANNEL as BrowserChannel | undefined) ?? 'msedge');

  const isEdge = browserChannel === 'msedge';

  console.log(
    `[Sesion] Browser=${browserChannel} | ProfileRoot=${profileRoot} | ProfileDir=${profileDirectory}`,
  );

  const launchArgs = [
    '--start-maximized',
    '--window-position=0,0',
    '--no-first-run',
    ...(isEdge ? ['--edge-skip-compat-layer-relaunch'] : []),
    ...(options.extraArgs ?? []),
  ];

  if (profileDirectory) {
    launchArgs.push(`--profile-directory=${profileDirectory}`);
  }

  console.log(`[BrowserSession] Lanzando ${isEdge ? 'Edge' : browserChannel} persistent context`);
  console.log(`[BrowserSession] browserChannel=${browserChannel}`);
  console.log(`[BrowserSession] userDataDir=${profileRoot}`);
  console.log(`[BrowserSession] executablePath=${browserChannel === 'msedge' ? 'msedge (channel)' : 'default'}`);
  console.log(`[BrowserSession] headless=false`);
  console.log(`[BrowserSession] args=${JSON.stringify(launchArgs)}`);

  let context;
  try {
    context = await chromium.launchPersistentContext(profileRoot, {
      channel: browserChannel,
      headless: false,
      ignoreHTTPSErrors: true,
      viewport: null,
      args: launchArgs,
    });
    console.log('[BrowserSession] Persistent context creado correctamente');
  } catch (error) {
    console.error('[BrowserSession] Error al crear persistent context:', error.message);
    console.error('[BrowserSession] Stack:', error.stack);
    console.error(`[BrowserSession] userDataDir=${profileRoot}`);
    console.error(`[BrowserSession] executablePath=default`);
    console.error(`[BrowserSession] args=${JSON.stringify(launchArgs)}`);
    throw error;
  }
  const existingPages = context.pages();
  const page =
    options.preferNewPage
      ? await context.newPage()
      : existingPages[0] ?? (await context.newPage());

  await maximizeWindow(page);

  if (options.initialUrl) {
    await page.goto(options.initialUrl, { waitUntil: 'domcontentloaded' }).catch(() => { });
  }

  return {
    context,
    page,
    pauseOnRegistroError: (process.env.PW_PAUSE_ON_ERROR ?? '0') === '1',
    keepBrowserOpenOnFailures: (process.env.PW_KEEP_BROWSER_ON_FAIL ?? '1') === '1',
    holdOnFailure: (process.env.PW_HOLD_ON_FAIL ?? '1') === '1',
    continueOnRegistroError:
      (process.env.PW_CONTINUE_ON_ERROR ?? String(Number(options.defaultContinueOnError ?? true))) ===
      '1',
  };
}
