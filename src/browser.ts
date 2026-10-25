import fs from 'node:fs';
import { chromium, type BrowserContext, type Page, type Request } from 'playwright-core';
import { CliError, profileDir, type Status } from './sessions.js';

// OWA vive en outlook.cloud.microsoft; outlook.office.com redirige allí.
export const OWA_URL = 'https://outlook.cloud.microsoft/mail/';
export const OWA_HOSTS = ['outlook.office.com', 'outlook.office365.com', 'outlook.cloud.microsoft'];
export const LOGIN_HOSTS = ['login.microsoftonline.com', 'login.live.com', 'login.microsoft.com'];

export const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

export const hostOf = (url: string) => {
  try {
    return new URL(url).hostname;
  } catch {
    return ''; // about:blank
  }
};

/**
 * Si al agotar la espera la página está en una web que no es OWA, se da por caducada: con un
 * proveedor de identidad federado (ADFS, Okta...) el login acaba en su propio host, no en Microsoft.
 */
export const stuckOutsideOwa = (url: string) => {
  const host = hostOf(url);
  return host !== '' && !OWA_HOSTS.includes(host);
};

export async function launch(name: string, headless: boolean): Promise<BrowserContext> {
  const dir = profileDir(name);
  fs.mkdirSync(dir, { recursive: true });
  try {
    return await chromium.launchPersistentContext(dir, {
      channel: 'msedge',
      headless,
      viewport: headless ? { width: 1400, height: 900 } : null,
      args: ['--no-first-run', '--no-default-browser-check'],
      // Playwright añade --no-sandbox si no se pide expresamente; Edge abre correo ajeno, así que con sandbox.
      chromiumSandbox: true,
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    // Con el perfil en uso, Edge delega en el proceso existente y sale: Playwright lo ve como cierre.
    if (/has been closed|exitCode|ProcessSingleton|already in use|SingletonLock/i.test(msg)) {
      throw new CliError(`la sesión ${name} está en uso (otro owa o Edge tiene abierto su perfil)`, 1);
    }
    throw new CliError(`No se pudo lanzar Edge: ${msg.split('\n')[0]}`, 1);
  }
}

/** Lee el payload de un JWT. Nunca se devuelve ni se guarda el token entero. */
export function decodeJwtPayload(token: string): Record<string, unknown> | null {
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  try {
    return JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
  } catch {
    return null;
  }
}

export function bearerOf(authorization: string | undefined): string | null {
  const m = authorization?.match(/^Bearer\s+(\S+)$/i);
  return m ? m[1] : null;
}

export const isOwaHost = (url: string) => {
  try {
    return OWA_HOSTS.includes(new URL(url).hostname);
  } catch {
    return false;
  }
};

export const isServiceResponse = (url: string, status: number) => {
  try {
    const u = new URL(url);
    return status === 200 && OWA_HOSTS.includes(u.hostname) && /\/owa\/(service\.svc|startupdata\.ashx)/i.test(u.pathname);
  } catch {
    return false;
  }
};

/** Observa las peticiones a OWA y recuerda la cuenta (upn) que aparezca en un Bearer. */
export function watchAccount(context: BrowserContext): { get: () => string | null } {
  let account: string | null = null;
  context.on('request', (req: Request) => {
    if (account || !isOwaHost(req.url())) return;
    const token = bearerOf(req.headers()['authorization']);
    if (!token) return;
    const p = decodeJwtPayload(token);
    const a = p && (p.upn ?? p.preferred_username ?? p.unique_name);
    if (typeof a === 'string') account = a;
  });
  return { get: () => account };
}

export interface CheckResult {
  status: Status;
  account: string | null;
  error?: string;
}

/**
 * Abre OWA en headless y decide si la sesión sirve. Con `keepOpenMs` mantiene la página abierta
 * tras un OK para que MSAL renueve sus tokens (keepalive).
 */
export async function headlessCheck(name: string, keepOpenMs = 0): Promise<CheckResult> {
  const context = await launch(name, true);
  try {
    const acct = watchAccount(context);
    const page = context.pages()[0] ?? (await context.newPage());
    let serviceOk = false;
    page.on('response', (r) => {
      if (isServiceResponse(r.url(), r.status())) serviceOk = true;
    });
    let gotoError: string | undefined;
    page.goto(OWA_URL, { waitUntil: 'domcontentloaded', timeout: 60_000 }).catch((e) => {
      gotoError = String(e?.message ?? e).split('\n')[0];
    });

    const deadline = Date.now() + 60_000;
    let status: Status = 'unknown';
    while (Date.now() < deadline) {
      await sleep(500);
      let host = '';
      try {
        host = new URL(page.url()).hostname;
      } catch {
        /* about:blank */
      }
      if (LOGIN_HOSTS.includes(host)) {
        status = 'expired';
        break;
      }
      const treeVisible = OWA_HOSTS.includes(host) && (await page.locator('[role="tree"]').first().isVisible().catch(() => false));
      if (serviceOk || treeVisible) {
        status = 'ok';
        break;
      }
      if (gotoError && !page.url().startsWith('http')) {
        return { status: 'error', account: null, error: gotoError };
      }
    }
    if (status === 'unknown' && stuckOutsideOwa(page.url())) status = 'expired';
    if (status === 'ok') {
      // Margen corto a que aparezca un Bearer para conocer la cuenta.
      const until = Date.now() + 3000;
      while (!acct.get() && Date.now() < until) await sleep(300);
      if (keepOpenMs > 0) await sleep(keepOpenMs);
    }
    return { status, account: acct.get() };
  } catch (e) {
    return { status: 'error', account: null, error: e instanceof Error ? e.message.split('\n')[0] : String(e) };
  } finally {
    await context.close().catch(() => {});
  }
}

/** Login visible: espera a que el usuario cierre la ventana (o al timeout). */
export async function interactiveLogin(name: string, timeoutMs: number): Promise<{ account: string | null; timedOut: boolean }> {
  const context = await launch(name, false);
  const acct = watchAccount(context);
  let timedOut = false;
  const closed = new Promise<void>((resolve) => context.on('close', () => resolve()));
  const trackPage = (p: Page) => {
    p.on('close', () => {
      // Edge puede seguir vivo sin ventanas: si no queda ninguna página, cerramos nosotros.
      if (context.pages().length === 0) context.close().catch(() => {});
    });
  };
  context.pages().forEach(trackPage);
  context.on('page', trackPage);
  const page = context.pages()[0] ?? (await context.newPage());
  page.goto(OWA_URL).catch(() => {});
  const timer = setTimeout(() => {
    timedOut = true;
    context.close().catch(() => {});
  }, timeoutMs);
  await closed;
  clearTimeout(timer);
  return { account: acct.get(), timedOut };
}

export async function remoteLogout(name: string): Promise<void> {
  const context = await launch(name, true);
  try {
    const page = context.pages()[0] ?? (await context.newPage());
    await page.goto('https://login.microsoftonline.com/common/oauth2/v2.0/logout', { waitUntil: 'load', timeout: 30_000 }).catch(() => {});
  } finally {
    await context.close().catch(() => {});
  }
}
