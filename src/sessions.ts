import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export type Status = 'ok' | 'expired' | 'error' | 'unknown';

export interface SessionMeta {
  name: string;
  account: string | null;
  createdAt: string;
  lastCheckAt: string | null;
  lastOkAt: string | null;
  lastStatus: Status;
}

/** Error con código de salida propio. */
export class CliError extends Error {
  constructor(message: string, public code: number = 1) {
    super(message);
  }
}

const NAME_RE = /^[a-z0-9][a-z0-9_-]{0,31}$/;

/**
 * Fuera de AppData a propósito: las apps MSIX (la de Claude, entre ellas) virtualizan AppData para
 * todo lo que lanzan, y una sesión creada por un agente no existiría para una tarea programada.
 */
export function homeDir(): string {
  const env = process.env.OWA_CLI_HOME;
  if (env) return path.resolve(env);
  return path.join(os.homedir(), '.owa-cli');
}

export function sessionsRoot(): string {
  return path.join(homeDir(), 'sessions');
}

export function validateName(name: string): string {
  if (!NAME_RE.test(name)) {
    throw new CliError(`Nombre de sesión no válido: "${name}" (debe cumplir ${NAME_RE.source})`, 2);
  }
  return name;
}

export const sessionDir = (name: string) => path.join(sessionsRoot(), validateName(name));
export const profileDir = (name: string) => path.join(sessionDir(name), 'profile');
const metaFile = (name: string) => path.join(sessionDir(name), 'session.json');

export function readMeta(name: string): SessionMeta | null {
  try {
    return JSON.parse(fs.readFileSync(metaFile(name), 'utf8')) as SessionMeta;
  } catch {
    return null;
  }
}

export function writeMeta(meta: SessionMeta): void {
  fs.mkdirSync(sessionDir(meta.name), { recursive: true });
  fs.writeFileSync(metaFile(meta.name), JSON.stringify(meta, null, 2) + '\n');
}

/** Devuelve la sesión o lanza CliError código 4. */
export function requireSession(name: string): SessionMeta {
  const meta = readMeta(name);
  if (!meta) throw new CliError(`La sesión "${name}" no existe. Créala con: owa login --session ${name}`, 4);
  return meta;
}

export function newMeta(name: string): SessionMeta {
  return { name, account: null, createdAt: new Date().toISOString(), lastCheckAt: null, lastOkAt: null, lastStatus: 'unknown' };
}

export function listSessions(): SessionMeta[] {
  const root = sessionsRoot();
  if (!fs.existsSync(root)) return [];
  const out: SessionMeta[] = [];
  for (const entry of fs.readdirSync(root, { withFileTypes: true })) {
    if (!entry.isDirectory() || !NAME_RE.test(entry.name)) continue;
    const meta = readMeta(entry.name);
    if (meta) out.push(meta);
  }
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

/** Aplica el resultado de una comprobación a los metadatos y los guarda. */
export function recordCheck(meta: SessionMeta, status: Status, account: string | null): SessionMeta {
  const now = new Date().toISOString();
  meta.lastCheckAt = now;
  meta.lastStatus = status;
  if (status === 'ok') meta.lastOkAt = now;
  if (account) meta.account = account;
  writeMeta(meta);
  return meta;
}

/** Borra la carpeta de una sesión, comprobando antes que cuelga directamente de la raíz de sesiones. */
export function deleteSession(name: string): void {
  const dir = path.resolve(sessionDir(name));
  const root = path.resolve(sessionsRoot());
  const rel = path.relative(root, dir);
  if (!rel || rel.startsWith('..') || path.isAbsolute(rel) || path.dirname(dir) !== root) {
    throw new CliError(`Ruta fuera de la raíz de sesiones, no se borra: ${dir}`, 1);
  }
  fs.rmSync(dir, { recursive: true, force: true });
}
