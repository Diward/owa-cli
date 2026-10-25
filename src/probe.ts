import { launch, OWA_URL, bearerOf, decodeJwtPayload, isServiceResponse, sleep } from './browser.js';

const PROBE_SUFFIXES = ['office.com', 'office365.com', 'office.net', 'microsoft.com', 'microsoftonline.com'];

interface Entry {
  method: string;
  host: string;
  path: string;
  action: string | null;
  count: number;
  auth: Set<string>;
  xHeaders: Set<string>;
  cookies: boolean;
  statuses: Set<number>;
}

/** Diagnóstico: qué peticiones hace OWA y cómo se autentican. Nunca emite valores de token. */
export async function probe(name: string, seconds: number): Promise<unknown> {
  const context = await launch(name, true);
  try {
    const entries = new Map<string, Entry>();
    const audiences = new Map<string, { aud: unknown; scp: unknown; appid: unknown; exp: string | null }>();
    const pending: Promise<void>[] = [];
    let serviceOk = false;

    context.on('response', (resp) => {
      pending.push(
        (async () => {
          const req = resp.request();
          const u = new URL(req.url());
          if (!PROBE_SUFFIXES.some((s) => u.hostname === s || u.hostname.endsWith('.' + s))) return;
          const headers = await req.allHeaders().catch(() => req.headers());
          const action = u.searchParams.get('action');
          const key = [req.method(), u.hostname, u.pathname, action ?? ''].join(' ');
          let e = entries.get(key);
          if (!e) {
            e = { method: req.method(), host: u.hostname, path: u.pathname, action, count: 0, auth: new Set(), xHeaders: new Set(), cookies: false, statuses: new Set() };
            entries.set(key, e);
          }
          e.count++;
          e.auth.add(headers['authorization']?.split(/\s+/)[0] ?? 'none');
          for (const h of Object.keys(headers)) if (h.startsWith('x-')) e.xHeaders.add(h);
          if (headers['cookie']) e.cookies = true;
          e.statuses.add(resp.status());
          if (isServiceResponse(req.url(), resp.status())) serviceOk = true;

          const token = bearerOf(headers['authorization']);
          const p = token ? decodeJwtPayload(token) : null;
          if (p) {
            const aKey = `${String(p.aud)}|${String(p.scp)}`;
            if (!audiences.has(aKey)) {
              audiences.set(aKey, {
                aud: p.aud ?? null,
                scp: p.scp ?? null,
                appid: p.appid ?? p.azp ?? null,
                exp: typeof p.exp === 'number' ? new Date(p.exp * 1000).toISOString() : null,
              });
            }
          }
        })().catch(() => {}),
      );
    });

    const page = context.pages()[0] ?? (await context.newPage());
    page.goto(OWA_URL, { waitUntil: 'domcontentloaded', timeout: 60_000 }).catch(() => {});
    await sleep(seconds * 1000);
    await Promise.allSettled(pending);

    const treeVisible = await page.locator('[role="tree"]').first().isVisible().catch(() => false);
    return {
      finalUrl: page.url().split(/[?#]/)[0], // sin query/fragment: pueden llevar códigos de auth
      mailboxLoaded: serviceOk || treeVisible,
      seconds,
      requests: [...entries.values()]
        .sort((a, b) => a.host.localeCompare(b.host) || a.path.localeCompare(b.path))
        .map((e) => ({
          method: e.method,
          host: e.host,
          path: e.path,
          action: e.action,
          count: e.count,
          auth: [...e.auth],
          xHeaders: [...e.xHeaders].sort(),
          cookies: e.cookies,
          status: [...e.statuses].sort(),
        })),
      tokenAudiences: [...audiences.values()],
    };
  } finally {
    await context.close().catch(() => {});
  }
}
