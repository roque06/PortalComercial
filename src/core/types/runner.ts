import type { BrowserContext, Page } from '@playwright/test';

export type BrowserChannel = 'chrome' | 'msedge';

export interface BrowserSession {
  context: BrowserContext;
  page: Page;
  keepBrowserOpenOnFailures: boolean;
  holdOnFailure: boolean;
  pauseOnRegistroError: boolean;
  continueOnRegistroError: boolean;
}

export interface RegistroErrorInfo {
  identificacion: string;
  error: string;
}

export interface RegistroSkipInfo {
  identificacion: string;
  razon: string;
}

export interface RegistroRunnerState {
  failedRegistros: RegistroErrorInfo[];
  skippedRegistros: RegistroSkipInfo[];
}
