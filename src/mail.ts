import fs from 'node:fs';
import type { OwaClient } from './owa.js';
import { CliError } from './sessions.js';

/*
 * Correo sobre service.svc (EWS en JSON).
 *
 * Comportamiento medido contra Exchange Online:
 *  - El JSON va en el CUERPO del POST; no hizo falta x-owa-urlpostdata (queda como respaldo en owa.ts).
 *  - FindItem + Restriction funciona. FindItem + QueryString (AQS) responde "NoError" pero con 0
 *    resultados SIEMPRE (crea una carpeta de búsqueda en el servidor que nunca se llena), y mezclar
 *    QueryString con Restriction da HTTP 400. Por eso TODOS los filtros, texto incluido, se traducen
 *    a Restriction (Contains sobre Subject/Body/From/DisplayTo) y --query es un subconjunto de AQS
 *    que se interpreta aquí, no en el servidor.
 *  - FindItem no devuelve ToRecipients: en la lista, "to" son nombres (DisplayTo).
 */

const P = (u: string) => ({ __type: 'PropertyUri:#Exchange', FieldURI: u });
const FLAG_STATUS = { __type: 'ExtendedPropertyUri:#Exchange', PropertyTag: '0x1090', PropertyType: 'Integer' };
const constant = (v: string) => ({ __type: 'FieldURIOrConstantType:#Exchange', Item: { __type: 'Constant:#Exchange', Value: v } });

export const DISTINGUISHED = ['inbox', 'drafts', 'sentitems', 'deleteditems', 'junkemail', 'archive', 'outbox'];
const REF_PREFIX = 60; // los ids de un buzón comparten este prefijo; `ref` es el resto

// ---------------------------------------------------------------- carpetas

export interface Folder {
  id: string;
  displayName: string;
  path: string;
  unread: number;
  total: number;
  distinguished: string | null;
}

export async function fetchFolders(c: OwaClient): Promise<Folder[]> {
  const raw: any[] = [];
  for (let offset = 0; ; offset += 500) {
    const r = await c.call('FindFolder', {
      __type: 'FindFolderRequest:#Exchange',
      FolderShape: {
        __type: 'FolderResponseShape:#Exchange',
        BaseShape: 'Default',
        AdditionalProperties: [P('ParentFolderId'), P('FolderClass'), P('DistinguishedFolderId')],
      },
      ParentFolderIds: [{ __type: 'DistinguishedFolderId:#Exchange', Id: 'msgfolderroot' }],
      Traversal: 'Deep',
      Paging: { __type: 'IndexedPageView:#Exchange', BasePoint: 'Beginning', Offset: offset, MaxEntriesReturned: 500 },
    });
    const items: any[] = r.RootFolder?.Folders ?? [];
    raw.push(...items);
    if (r.RootFolder?.IncludesLastItemInRange !== false || !items.length) break;
  }
  const byId = new Map<string, any>(raw.map((f) => [f.FolderId.Id, f]));
  const pathOf = (f: any, depth = 0): string => {
    const parent = f.ParentFolderId && byId.get(f.ParentFolderId.Id);
    return parent && depth < 20 ? `${pathOf(parent, depth + 1)}/${f.DisplayName}` : f.DisplayName;
  };
  return raw
    .filter((f) => !f.FolderClass || /^IPF\.Note(\.|$)/.test(f.FolderClass))
    .map((f) => ({
      id: f.FolderId.Id,
      displayName: f.DisplayName,
      path: pathOf(f),
      unread: f.UnreadCount ?? 0,
      total: f.TotalCount ?? 0,
      distinguished: f.DistinguishedFolderId ?? null,
    }))
    .sort((a, b) => a.path.localeCompare(b.path));
}

export interface FolderTarget {
  /** Para el Body EWS. */
  ref: Record<string, unknown>;
  label: string;
}

/** Distinguido (sin llamar), id, o nombre/ruta de `folders` (sin distinguir mayúsculas). */
export async function resolveFolder(c: OwaClient, arg: string): Promise<FolderTarget> {
  const low = arg.trim().toLowerCase();
  if (DISTINGUISHED.includes(low)) return { ref: { __type: 'DistinguishedFolderId:#Exchange', Id: low }, label: low };
  if (/^[A-Za-z0-9+/=_-]{40,}$/.test(arg.trim())) return { ref: { __type: 'FolderId:#Exchange', Id: arg.trim() }, label: 'id' };
  const folders = await fetchFolders(c);
  const exactPath = folders.filter((f) => f.path.toLowerCase() === low);
  const byName = folders.filter((f) => f.displayName.toLowerCase() === low);
  // Final de la ruta ("Clientes/Acme"): no depende del nombre localizado de la bandeja de entrada.
  const bySuffix = folders.filter((f) => f.path.toLowerCase().endsWith('/' + low));
  const hit = exactPath.length ? exactPath : byName.length ? byName : bySuffix;
  if (hit.length === 1) return { ref: { __type: 'FolderId:#Exchange', Id: hit[0].id }, label: hit[0].path };
  if (hit.length > 1) {
    throw new CliError(`Carpeta ambigua "${arg}". Candidatas (usa la ruta completa):\n` + hit.map((f) => `  ${f.path}`).join('\n'), 2);
  }
  const near = folders.filter((f) => f.path.toLowerCase().includes(low)).slice(0, 8);
  throw new CliError(
    `No existe la carpeta "${arg}".` + (near.length ? ` ¿Quizá?\n${near.map((f) => `  ${f.path}`).join('\n')}` : ' Ver: owa folders'),
    2,
  );
}

// ---------------------------------------------------------------- filtros

export interface Filters {
  read?: boolean; // true = leídos, false = no leídos
  flagged?: boolean;
  hasAttachments?: boolean;
  importance?: 'High' | 'Normal' | 'Low';
  categories: string[];
  from: string[];
  to: string[];
  subject: string[];
  body: string[];
  any: string[];
  since?: Date; // inclusive
  until?: Date; // exclusivo
}

export const emptyFilters = (): Filters => ({ categories: [], from: [], to: [], subject: [], body: [], any: [] });

/** YYYY-MM-DD como fecha local a las 00:00. */
export function parseLocalDate(s: string, flag: string): Date {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  const d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  if (!m || !d || d.getMonth() !== Number(m[2]) - 1) throw new CliError(`--${flag} debe ser YYYY-MM-DD (recibido "${s}")`, 2);
  return d;
}

export const addDays = (d: Date, n: number) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);

function yesNo(v: string, key: string): boolean {
  const l = v.toLowerCase();
  if (['yes', 'true', 'sí', 'si'].includes(l)) return true;
  if (['no', 'false'].includes(l)) return false;
  throw new CliError(`--query: "${key}:" espera yes o no (recibido "${v}")`, 2);
}

/**
 * Subconjunto de AQS interpretado en el cliente (el servidor no devuelve nada con QueryString):
 * from: to: subject: body: category: importance: hasattachments: isread: isflagged:
 * received:/received>=/received<=/received>/received< con YYYY-MM-DD, "frases" y palabras sueltas.
 */
export function parseQuery(q: string, f: Filters): void {
  const re = /(?:([A-Za-z]+)\s*(>=|<=|:|>|<|=)\s*)?(?:"([^"]*)"|(\S+))/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(q))) {
    const key = m[1]?.toLowerCase();
    const op = m[2];
    const val = m[3] ?? m[4];
    if (!key) {
      f.any.push(val);
      continue;
    }
    switch (key) {
      case 'from': f.from.push(val); break;
      case 'to': f.to.push(val); break;
      case 'subject': f.subject.push(val); break;
      case 'body': f.body.push(val); break;
      case 'category': case 'categories': f.categories.push(val); break;
      case 'hasattachments': f.hasAttachments = yesNo(val, key); break;
      case 'isread': f.read = yesNo(val, key); break;
      case 'isflagged': f.flagged = yesNo(val, key); break;
      case 'importance': f.importance = parseImportance(val); break;
      case 'kind':
        if (!['email', 'mail', 'emails'].includes(val.toLowerCase())) throw new CliError(`--query: kind:${val} no soportado`, 2);
        break;
      case 'received': case 'date': {
        const d = parseLocalDate(val, 'query received');
        if (op === '>=') f.since = d;
        else if (op === '>') f.since = addDays(d, 1);
        else if (op === '<=') f.until = addDays(d, 1);
        else if (op === '<') f.until = d;
        else { f.since = d; f.until = addDays(d, 1); }
        break;
      }
      default:
        throw new CliError(
          `--query: "${key}:" no está soportado. Admite: from to subject body category importance hasattachments isread isflagged received kind, "frases" y palabras sueltas`,
          2,
        );
    }
  }
}

export function parseImportance(v: string): 'High' | 'Normal' | 'Low' {
  const l = v.toLowerCase();
  if (l === 'high' || l === 'alta') return 'High';
  if (l === 'normal') return 'Normal';
  if (l === 'low' || l === 'baja') return 'Low';
  throw new CliError(`--importance debe ser high, normal o low (recibido "${v}")`, 2);
}

const cmp = (type: string, path: unknown, value: string) => ({ __type: `${type}:#Exchange`, Item: path, FieldURIOrConstant: constant(value) });
const SENDER_NAME = { __type: 'ExtendedPropertyUri:#Exchange', PropertyTag: '0x0C1A', PropertyType: 'String' };
const containsPath = (path: unknown, v: string) => ({
  __type: 'Contains:#Exchange',
  ContainmentMode: 'Substring',
  ContainmentComparison: 'IgnoreCase',
  Item: path,
  Constant: { __type: 'Constant:#Exchange', Value: v },
});
const contains = (field: string, v: string) => ({
  __type: 'Contains:#Exchange',
  ContainmentMode: 'Substring',
  ContainmentComparison: 'IgnoreCase',
  Item: P(field),
  Constant: { __type: 'Constant:#Exchange', Value: v },
});

export function buildRestriction(f: Filters): unknown | undefined {
  const items: unknown[] = [];
  if (f.read !== undefined) items.push(cmp('IsEqualTo', P('IsRead'), String(f.read)));
  if (f.flagged !== undefined) items.push(cmp(f.flagged ? 'IsEqualTo' : 'IsNotEqualTo', FLAG_STATUS, '2'));
  if (f.hasAttachments !== undefined) items.push(cmp('IsEqualTo', P('HasAttachments'), String(f.hasAttachments)));
  if (f.importance) items.push(cmp('IsEqualTo', P('Importance'), f.importance));
  for (const v of f.categories) items.push(contains('Categories', v));
  // Contains sobre From solo mira la dirección; el nombre visible está en PidTagSenderName.
  for (const v of f.from) items.push({ __type: 'Or:#Exchange', Items: [contains('From', v), containsPath(SENDER_NAME, v)] });
  for (const v of f.to) items.push(contains('DisplayTo', v));
  for (const v of f.subject) items.push(contains('Subject', v));
  for (const v of f.body) items.push(contains('Body', v));
  for (const v of f.any) {
    items.push({ __type: 'Or:#Exchange', Items: [...['Subject', 'From', 'DisplayTo', 'Body'].map((fl) => contains(fl, v)), containsPath(SENDER_NAME, v)] });
  }
  if (f.since) items.push(cmp('IsGreaterThanOrEqualTo', P('DateTimeReceived'), f.since.toISOString()));
  if (f.until) items.push(cmp('IsLessThan', P('DateTimeReceived'), f.until.toISOString()));
  if (!items.length) return undefined;
  return { __type: 'RestrictionType:#Exchange', Item: items.length === 1 ? items[0] : { __type: 'And:#Exchange', Items: items } };
}

// ---------------------------------------------------------------- listar

export interface Person {
  name: string | null;
  address: string | null;
}

export interface MessageSummary {
  id: string;
  ref: string;
  receivedAt: string | null;
  from: Person | null;
  to: Person[];
  subject: string;
  preview: string;
  isRead: boolean;
  hasAttachments: boolean;
  flagged: boolean;
  importance: string;
  categories: string[];
  untrusted: true;
}

const person = (m: any): Person | null => {
  const mb = m?.Mailbox ?? m;
  if (!mb || (!mb.EmailAddress && !mb.Name)) return null;
  return { name: mb.Name ?? null, address: mb.EmailAddress ?? null };
};

const iso = (s: unknown) => (typeof s === 'string' && !Number.isNaN(Date.parse(s)) ? new Date(s).toISOString() : null);
export const toRef = (id: string) => id.slice(REF_PREFIX);

export interface ListResult {
  folder: string;
  total: number;
  offset: number;
  limit: number;
  hasMore: boolean;
  messages: MessageSummary[];
}

export async function listMessages(c: OwaClient, target: FolderTarget, f: Filters, limit: number, offset: number): Promise<ListResult> {
  const body: Record<string, unknown> = {
    __type: 'FindItemRequest:#Exchange',
    ItemShape: {
      __type: 'ItemResponseShape:#Exchange',
      BaseShape: 'IdOnly',
      AdditionalProperties: ['DateTimeReceived', 'Subject', 'IsRead', 'HasAttachments', 'Importance', 'Categories', 'Flag', 'From', 'DisplayTo', 'Preview'].map(P),
    },
    ParentFolderIds: [target.ref],
    Traversal: 'Shallow',
    Paging: { __type: 'IndexedPageView:#Exchange', BasePoint: 'Beginning', Offset: offset, MaxEntriesReturned: limit },
    SortOrder: [{ __type: 'SortResults:#Exchange', Order: 'Descending', Path: P('DateTimeReceived') }],
  };
  const restriction = buildRestriction(f);
  if (restriction) body.Restriction = restriction;
  const r = await c.call('FindItem', body);
  const items: any[] = r.RootFolder?.Items ?? [];
  const messages = items.map((i): MessageSummary => ({
    id: i.ItemId.Id,
    ref: toRef(i.ItemId.Id),
    receivedAt: iso(i.DateTimeReceived),
    from: person(i.From),
    // FindItem solo da los nombres de los destinatarios (DisplayTo); las direcciones, en `read`.
    to: typeof i.DisplayTo === 'string' && i.DisplayTo ? i.DisplayTo.split(';').map((n: string) => ({ name: n.trim(), address: null })) : [],
    subject: i.Subject ?? '',
    preview: (i.Preview ?? '').replace(/\s+/g, ' ').trim(),
    isRead: !!i.IsRead,
    hasAttachments: !!i.HasAttachments,
    flagged: i.Flag?.FlagStatus === 'Flagged',
    importance: i.Importance ?? 'Normal',
    categories: Array.isArray(i.Categories) ? i.Categories : [],
    untrusted: true,
  }));
  return {
    folder: target.label,
    total: r.RootFolder?.TotalItemsInView ?? messages.length,
    offset,
    limit,
    hasMore: r.RootFolder?.IncludesLastItemInRange === false && offset + messages.length < (r.RootFolder?.TotalItemsInView ?? 0),
    messages,
  };
}

// ---------------------------------------------------------------- leer

/** Un ref corto (sin el prefijo común del buzón) se completa con el prefijo de cualquier mensaje del buzón. */
export async function expandId(c: OwaClient, arg: string): Promise<string> {
  const id = arg.trim();
  if (id.length >= 40) return id;
  if (!/^[A-Za-z0-9+/=_-]{6,}$/.test(id)) throw new CliError(`Identificador de mensaje no válido: "${id}"`, 2);
  for (const folder of ['inbox', 'sentitems', 'archive', 'drafts']) {
    const r = await c.call('FindItem', {
      __type: 'FindItemRequest:#Exchange',
      ItemShape: { __type: 'ItemResponseShape:#Exchange', BaseShape: 'IdOnly' },
      ParentFolderIds: [{ __type: 'DistinguishedFolderId:#Exchange', Id: folder }],
      Traversal: 'Shallow',
      Paging: { __type: 'IndexedPageView:#Exchange', BasePoint: 'Beginning', Offset: 0, MaxEntriesReturned: 1 },
    });
    const tpl: string | undefined = r.RootFolder?.Items?.[0]?.ItemId?.Id;
    if (tpl) return tpl.slice(0, REF_PREFIX) + id;
  }
  throw new CliError('No se pudo completar el identificador corto; usa el id completo (--json).', 1);
}

export interface Attachment {
  name: string;
  size: number | null;
  contentType: string | null;
}

export interface MessageFull {
  id: string;
  untrusted: true;
  subject: string;
  from: Person | null;
  to: Person[];
  cc: Person[];
  bcc: Person[];
  receivedAt: string | null;
  sentAt: string | null;
  isRead: boolean;
  importance: string;
  categories: string[];
  internetMessageId: string | null;
  conversationId: string | null;
  attachments: Attachment[];
  body: { type: 'text' | 'html'; untrusted: true; content: string };
}

const people = (l: unknown): Person[] => (Array.isArray(l) ? (l.map(person).filter(Boolean) as Person[]) : []);

export async function readMessage(c: OwaClient, idArg: string, html: boolean): Promise<MessageFull> {
  const id = await expandId(c, idArg);
  let r: any;
  try {
    r = await c.call('GetItem', {
      __type: 'GetItemRequest:#Exchange',
      ItemShape: {
        __type: 'ItemResponseShape:#Exchange',
        BaseShape: 'IdOnly',
        BodyType: html ? 'HTML' : 'Text',
        AdditionalProperties: [
          'Subject', 'Body', 'From', 'ToRecipients', 'CcRecipients', 'BccRecipients', 'DateTimeReceived', 'DateTimeSent', 'IsRead',
          'Importance', 'Categories', 'InternetMessageId', 'ConversationId', 'Attachments', 'HasAttachments',
        ].map(P),
      },
      ItemIds: [{ __type: 'ItemId:#Exchange', Id: id }],
    });
  } catch (e) {
    if (e instanceof CliError && /ErrorInvalidId|ErrorItemNotFound|ErrorInvalidIdMalformed/.test(e.message)) {
      throw new CliError('No existe ese mensaje (identificador no válido o borrado)', 1);
    }
    throw e;
  }
  const i = r.Items?.[0];
  if (!i) throw new CliError('GetItem no devolvió el mensaje', 1);
  return {
    id: i.ItemId?.Id ?? id,
    untrusted: true,
    subject: i.Subject ?? '',
    from: person(i.From),
    to: people(i.ToRecipients),
    cc: people(i.CcRecipients),
    bcc: people(i.BccRecipients),
    receivedAt: iso(i.DateTimeReceived),
    sentAt: iso(i.DateTimeSent),
    isRead: !!i.IsRead,
    importance: i.Importance ?? 'Normal',
    categories: Array.isArray(i.Categories) ? i.Categories : [],
    internetMessageId: i.InternetMessageId ?? null,
    conversationId: i.ConversationId?.Id ?? null,
    attachments: (Array.isArray(i.Attachments) ? i.Attachments : []).map((a: any) => ({
      name: a.Name ?? '',
      size: typeof a.Size === 'number' ? a.Size : null,
      contentType: a.ContentType ?? null,
    })),
    body: { type: html ? 'html' : 'text', untrusted: true, content: i.Body?.Value ?? '' },
  };
}

// ---------------------------------------------------------------- borrador

const ADDR_RE = /^[^\s@<>(),;:"]+@[^\s@<>(),;:"]+\.[^\s@<>(),;:"]+$/;

export function parseAddresses(raw: string | undefined, flag: string): string[] {
  if (!raw) return [];
  const list = raw.split(/[,;]/).map((s) => s.trim()).filter(Boolean);
  for (const a of list) if (!ADDR_RE.test(a)) throw new CliError(`--${flag}: dirección no válida "${a}"`, 2);
  return list;
}

export interface DraftInput {
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  body: string;
  html: boolean;
}

export function readBodyArg(body: string | undefined, bodyFile: string | undefined): string {
  if (body !== undefined && bodyFile !== undefined) throw new CliError('Usa --body o --body-file, no ambos', 2);
  if (bodyFile !== undefined) {
    try {
      return fs.readFileSync(bodyFile, 'utf8').replace(/^﻿/, '');
    } catch (e) {
      throw new CliError(`No se pudo leer --body-file: ${(e as Error).message}`, 2);
    }
  }
  if (body === undefined) throw new CliError('Falta el cuerpo: --body "..." o --body-file ruta', 2);
  return body;
}

export async function createDraft(c: OwaClient, d: DraftInput): Promise<{ id: string; link: string }> {
  const rcpt = (l: string[]) => l.map((a) => ({ __type: 'EmailAddressWrapper:#Exchange', EmailAddress: a, RoutingType: 'SMTP', MailboxType: 'OneOff' }));
  const r = await c.call('CreateItem', {
    __type: 'CreateItemRequest:#Exchange',
    MessageDisposition: 'SaveOnly', // nunca SendOnly/SendAndSaveCopy: owa no envía
    SavedItemFolderId: { __type: 'TargetFolderId:#Exchange', BaseFolderId: { __type: 'DistinguishedFolderId:#Exchange', Id: 'drafts' } },
    Items: [
      {
        __type: 'Message:#Exchange',
        Subject: d.subject,
        Body: { __type: 'BodyContentType:#Exchange', BodyType: d.html ? 'HTML' : 'Text', Value: d.body },
        ToRecipients: rcpt(d.to),
        CcRecipients: rcpt(d.cc),
        BccRecipients: rcpt(d.bcc),
      },
    ],
  });
  const id: string | undefined = r.Items?.[0]?.ItemId?.Id;
  if (!id) throw new CliError('CreateItem no devolvió el id del borrador', 1);
  return { id, link: `${c.origin}/mail/drafts/id/${encodeURIComponent(id)}` };
}
