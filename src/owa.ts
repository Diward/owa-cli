import { randomUUID } from 'node:crypto';
import type { Page } from 'playwright-core';
import { bearerOf, decodeJwtPayload, hostOf, isOwaHost, launch, OWA_URL, sleep, stuckOutsideOwa, LOGIN_HOSTS } from './browser.js';
import { CliError } from './sessions.js';

/*
 * Transporte de service.svc: primero el JSON en el cuerpo del POST, que es lo que acepta Exchange
 * Online hoy; si el servidor lo rechaza, se manda en la cabecera x-owa-urlpostdata (URL-encoded) con
 * cuerpo vacío, como hace a veces el propio OWA.
 */

const AUDIENCES = ['https://outlook.office.com', 'https://outlook.office365.com', 'https://outlook.cloud.microsoft'];
// Cabeceras del propio OWA que se reutilizan tal cual (los valores se quedan en memoria).
const COPY_HEADERS = ['x-anchormailbox', 'x-owa-sessionid', 'x-req-source', 'x-owa-actionsource', 'x-client-version', 'x-tenantid', 'x-owa-hosted-ux', 'x-js-clientid'];

export const EWS_VERSION = 'V2018_01_08';

export function requestHeader(): unknown {
  return {
    __type: 'JsonRequestHeaders:#Exchange',
    RequestServerVersion: EWS_VERSION,
    TimeZoneContext: {
      __type: 'TimeZoneContext:#Exchange',
      TimeZoneDefinition: { __type: 'TimeZoneDefinitionType:#Exchange', Id: 'UTC' },
    },
  };
}

export interface OwaClient {
  /** Llama a una acción de service.svc con el Body EWS y devuelve el primer mensaje de respuesta. */
  call(action: string, body: Record<string, unknown>): Promise<any>;
  /** Cuenta (upn) del token, si se conoce. */
  account: string | null;
  /** Origen de OWA en esta ejecución (p. ej. https://outlook.cloud.microsoft). */
  origin: string;
}

interface Captured {
  headers: Record<string, string>;
  account: string | null;
  origin: string;
}

function ewsError(msg: any): CliError {
  const code = msg?.ResponseCode ?? 'Desconocido';
  const text = msg?.MessageText ?? '';
  return new CliError(`Exchange rechazó la petición: ${code}${text ? ` — ${text}` : ''}`, 1);
}

async function postInPage(
  page: Page,
  url: string,
  headers: Record<string, string>,
  json: string,
  mode: 'body' | 'urlpostdata',
): Promise<{ status: number; text: string }> {
  return page.evaluate(
    async ({ url, headers, json, mode }) => {
      const h: Record<string, string> = { ...headers };
      let body: string | undefined;
      if (mode === 'body') {
        h['content-type'] = 'application/json; charset=utf-8';
        body = json;
      } else {
        h['x-owa-urlpostdata'] = encodeURIComponent(json);
        h['content-type'] = 'application/json; charset=utf-8';
      }
      const r = await fetch(url, { method: 'POST', headers: h, body, credentials: 'include' });
      return { status: r.status, text: await r.text() };
    },
    { url, headers, json, mode },
  );
}

/**
 * Abre OWA en headless, captura (solo en memoria) el Bearer y las cabeceras de las peticiones del propio
 * OWA a service.svc, y ejecuta `fn` con un cliente que llama a service.svc desde dentro de la página.
 */
export async function withOwa<T>(sessionName: string, fn: (c: OwaClient) => Promise<T>): Promise<T> {
  const context = await launch(sessionName, true);
  try {
    let captured: Captured | null = null;
    context.on('request', (req) => {
      if (captured || !isOwaHost(req.url())) return;
      if (!/\/owa\/service\.svc/i.test(new URL(req.url()).pathname)) return;
      req.allHeaders().then((h) => {
        if (captured) return;
        const token = bearerOf(h['authorization']);
        const p = token ? decodeJwtPayload(token) : null;
        if (!token || !p || !AUDIENCES.includes(String(p.aud).replace(/\/$/, ''))) return;
        const keep: Record<string, string> = { authorization: h['authorization'] };
        for (const k of COPY_HEADERS) if (h[k]) keep[k] = h[k];
        const a = p.upn ?? p.preferred_username ?? p.unique_name;
        captured = { headers: keep, account: typeof a === 'string' ? a : null, origin: new URL(req.url()).origin };
      }).catch(() => {});
    });

    const page = context.pages()[0] ?? (await context.newPage());
    page.goto(OWA_URL, { waitUntil: 'domcontentloaded', timeout: 60_000 }).catch(() => {});
    const deadline = Date.now() + 60_000;
    while (!captured && Date.now() < deadline) {
      await sleep(150);
      let host = '';
      try { host = new URL(page.url()).hostname; } catch { /* about:blank */ }
      if (LOGIN_HOSTS.includes(host)) {
        throw new CliError(`La sesión ${sessionName} ha caducado: hace falta que una persona ejecute "owa login --session ${sessionName}"`, 3);
      }
    }
    if (!captured && stuckOutsideOwa(page.url())) {
      throw new CliError(`La sesión ${sessionName} ha caducado (la página quedó en ${hostOf(page.url())}): hace falta que una persona ejecute "owa login --session ${sessionName}"`, 3);
    }
    if (!captured) throw new CliError('Tiempo agotado (60 s) esperando a que OWA cargue y emita su token. Prueba "owa status".', 1);
    const cap: Captured = captured;

    let mode: 'body' | 'urlpostdata' = 'body';
    const client: OwaClient = {
      account: cap.account,
      origin: cap.origin,
      async call(action, body) {
        const url = `${cap.origin}/owa/service.svc?action=${encodeURIComponent(action)}&app=Mail`;
        const json = JSON.stringify({ __type: `${action}JsonRequest:#Exchange`, Header: requestHeader(), Body: body });
        const headers = { ...cap.headers, action, 'x-owa-correlationid': randomUUID() };
        let res = await postInPage(page, url, headers, json, mode);
        if (mode === 'body' && (res.status === 400 || res.status === 500 || res.status === 411) && process.env.OWA_DEBUG) {
          console.error(`[debug] body transport -> ${res.status}`);
        }
        if (mode === 'body' && res.status >= 400 && res.status !== 401 && res.status !== 403) {
          const alt = await postInPage(page, url, headers, json, 'urlpostdata');
          if (alt.status < 400) { mode = 'urlpostdata'; res = alt; }
        }
        if (res.status === 401 || res.status === 403) {
          throw new CliError(`OWA rechazó el token (HTTP ${res.status}); prueba "owa status" o "owa login --session ${sessionName}"`, 3);
        }
        if (res.status >= 400) throw new CliError(`service.svc ${action}: HTTP ${res.status} ${res.text.slice(0, 200)}`, 1);
        let data: any;
        try { data = JSON.parse(res.text); } catch { throw new CliError(`service.svc ${action}: respuesta no JSON`, 1); }
        const items = data?.Body?.ResponseMessages?.Items;
        if (Array.isArray(items)) {
          const first = items[0];
          if (first && first.ResponseClass && first.ResponseClass !== 'Success') throw ewsError(first);
          return first;
        }
        if (data?.Body?.ResponseCode && data.Body.ResponseCode !== 'NoError') throw ewsError(data.Body);
        return data?.Body;
      },
    };
    return await fn(client);
  } finally {
    await context.close().catch(() => {});
  }
}
