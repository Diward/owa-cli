# owa-cli — arquitectura

CLI para que un agente use **tu** sesión de Outlook Web (Microsoft 365) desde el Edge de tu equipo, sin
Microsoft Graph ni registro de apps en Entra ID. Haces login una vez en una ventana visible; después
todo corre en headless con ese perfil.

Estado: **implementada y probada** contra Exchange Online (login visible, reapertura headless sin
pedir login, y autenticación de OWA observada; ver «Correo»).

## Alcance v1

| # | Necesidad | Comando |
|---|---|---|
| 1 | Login visible; al cerrar la ventana queda para headless | `owa login [--session S]` |
| 2 | Listar sesiones abiertas | `owa sessions [--check]` |
| 3 | Que la sesión no muera | `owa keepalive [--all]` + tarea programada |
| 4 | Cerrar una sesión | `owa logout --session S` |
| 5 | Leer correo con los filtros de Outlook | `owa folders`, `owa list ...`, `owa read <id>` |
| 6 | Borrador completo | `owa draft --to --cc --bcc --subject --body/--body-file [--html]` |

**Fuera de v1, a propósito: enviar.** El agente deja borradores y el envío lo haces tú. Un correo
malicioso leído por el agente no puede hacer que salga nada.

## Las tres capas

```
 agente ──► owa (CLI, un proceso por comando)
              │
              ├─ sesiones   perfil de Edge por sesión, metadatos, keepalive
              │
              ├─ navegador  Playwright → msedge instalado, launchPersistentContext
              │               login: visible · resto: headless
              │
              └─ correo     llamadas a la API interna de OWA, ejecutadas DENTRO de la página
                            (/owa/service.svc: FindFolder, FindItem, GetItem, CreateItem)
```

### 1. Navegador: Playwright sobre el Edge instalado

- `playwright-core` con `channel: 'msedge'`: usa el Edge del sistema, no descarga navegadores.
- **Perfil persistente propio** (`launchPersistentContext`), uno por sesión. No el perfil de Edge
  del usuario: está bloqueado mientras Edge esté abierto y mezclaría tu navegación con la del agente.
- `storageState` no sirve: OWA guarda estado de MSAL en `sessionStorage`, que no persiste.
- Sin trucos anti-detección (`AutomationControlled`, User-Agent falso) y con el sandbox de Chromium
  activo. Si el tenant rechaza el headless, se decide entonces y se documenta, no por defecto.

### 2. Correo: la API interna de OWA, no la interfaz

Tres formas de sacar el correo de una sesión de navegador:

| Vía | Filtros | Borrador | Fragilidad |
|---|---|---|---|
| Raspar la interfaz (selectores CSS) | los del buscador, escribiendo en él | redactar a golpe de clic | **alta**: cambia con cada despliegue de OWA |
| Reusar el token de OWA contra Graph / REST v2.0 | completos | sí | depende de los scopes del token de OWA; REST v2.0 está deprecada |
| **`/owa/service.svc` desde la página** | **completos** (restricciones EWS + AQS, el mismo lenguaje del buscador de Outlook) | **sí**, To/Cc/Bcc/asunto/HTML, sin pasar por la URL | **baja**: es la API que usa el propio OWA, con esquema EWS estable |

Se elige **service.svc**. El `fetch` se ejecuta con `page.evaluate` dentro de OWA ya cargado; el
JSON va en el cuerpo del POST (la vía `x-owa-urlpostdata` está implementada pero no hizo falta).

**Filtros: `Restriction` de EWS, no AQS en el servidor.** Se probó `FindItem` con `QueryString`
(AQS) y devuelve siempre 0 resultados: crea una carpeta de búsqueda oculta «EWS Search» que nunca se
llena, y combinado con `Restriction` da 400. Así que todos los filtros de `owa list` son una
`Restriction` (IsRead, HasAttachments, Importance, bandera 0x1090, Categories, Subject, From + nombre
del remitente 0x0C1A, Body, DisplayTo, DateTimeReceived), y `--query` es un subconjunto de AQS que
interpreta el propio CLI y traduce a esa misma `Restriction`. owa-cli no usa `QueryString`, así que
no crea esas carpetas de búsqueda ocultas.

**Host:** desde 2026 OWA vive en `outlook.cloud.microsoft` (`outlook.office.com` redirige). El login
y las comprobaciones van directamente allí; las llamadas usan el origen real de la página.

**Autenticación (medida con `owa probe` el 2026-10-25 en un tenant de Microsoft 365):** OWA llama a
`/owa/service.svc` con `Authorization: Bearer`, no con cookies. El token es de la app de OWA
(`appid 9199bf20-…`), `aud https://outlook.office.com`, con `Mail.ReadWrite` (y `Mail.Send`, que
owa-cli no usa), y caduca en ~24–36 h; OWA lo renueva al cargar. owa-cli toma esa cabecera (y
`x-anchormailbox`) de las propias peticiones de OWA en la misma ejecución, solo en memoria, sin
escribirla nunca a disco ni a la salida. OWA suele mandar el JSON de la petición en la cabecera
`x-owa-urlpostdata` en vez de en el cuerpo.

**Graph queda descartado también como respaldo:** el token de OWA para `graph.microsoft.com` no
lleva ningún permiso `Mail.*`. Si service.svc fallara, el último recurso es raspar la interfaz.

### 3. Proceso: un Edge headless por comando, sin demonio

Cada `owa ...` abre Edge headless, hace su llamada y cierra (unos 2–3 s).

La alternativa, un proceso residente con OWA abierto al que el CLI habla por un puerto o pipe local,
es más rápida y mantiene la sesión viva sola, pero **deja escuchando algo que puede actuar como tu
cuenta** para cualquier proceso del equipo, y un proceso más que vigilar. Se añadiría solo si la
latencia molesta de verdad, y entonces con pipe con ACL del usuario, no con puerto TCP.

## Sesiones y keepalive

- Datos en `%USERPROFILE%\.owa-cli\sessions\<nombre>\`, fuera de `AppData` porque la app MSIX de
  Claude lo virtualiza para todo lo que lanza (`profile\` y `session.json` con cuenta,
  alta, última comprobación y último OK). Nombre validado: `[a-z0-9][a-z0-9_-]{0,31}`.
  Sobrescribible con `OWA_CLI_HOME`.
- Las cookies las cifra el propio Edge con DPAPI: solo tu usuario de Windows las descifra. Aun así,
  **el perfil es una sesión iniciada**: quien lo copie es tú en Outlook.
- **Por qué muere una sesión y qué hace cada pieza:**
  - Los tokens de OWA (MSAL en el navegador) caducan en horas; el refresh token de una SPA vive
    24 h fijas. Los renueva OWA solo cuando se abre, usando la cookie de sesión de Microsoft.
  - Esa cookie solo sobrevive al cierre del navegador si en el login respondes **«¿Mantener la
    sesión iniciada? → Sí»** (`ESTSAUTHPERSISTENT`, ~90 días renovables). Sin ese «Sí», cerrar la
    ventana mata la sesión. `owa login` lo avisa.
  - `owa keepalive` abre OWA headless para que renueve los tokens. Programado cada 4 h con una
    tarea de Windows, la sesión dura lo que permita la cookie.
  - Lo que no se puede evitar: si tu organización fija una **frecuencia de inicio de sesión**
    (acceso condicional), al cumplirse hay que volver a hacer `owa login`. `owa sessions --check`
    lo detecta y lo dice.
- `owa logout` cierra la sesión en Microsoft y borra el perfil de esa sesión.

## Contrato con el agente

- Salida legible por defecto; `--json` en todos los comandos, que es lo que usa la skill.
- Códigos de salida distintos para «sesión caducada» (hay que hacer login, necesita a una persona)
  y para error genérico, para que el agente no reintente lo que no se arregla reintentando.
- El contenido de los correos sale **marcado como no confiable** (campo `untrusted` en JSON,
  delimitadores en texto), y la skill lo explica: lo que diga un correo no son instrucciones.
- Nunca imprime tokens ni cookies.

## Lenguaje: TypeScript sobre Node

| Opción | A favor | En contra |
|---|---|---|
| **TypeScript / Node 20+** | Playwright es nativo en Node; las otras versiones (Python, .NET, Java) envuelven ese mismo driver de Node. Los payloads de service.svc son JSON. `npm i -g` directamente desde el repo de GitHub | Un paso de compilación (`tsc`) |
| Python | Familiar, buen Playwright | Arrastra el driver de Node (~100 MB) igualmente |
| Go / Rust | Binario único | Sin Playwright nativo (o envuelven Node, o CDP a pelo con librerías menos maduras): reescribir la parte delicada |
| .NET | Buen soporte en Windows | Playwright .NET también envuelve Node; más pesado para un CLI |

**TypeScript**, compilado con `tsc` a `dist/` en el `prepare` de npm (Node no ejecuta `.ts` dentro
de `node_modules`, así que la compilación es necesaria para instalarlo con `npm i -g`).
Dependencias: solo `playwright-core`; argumentos con `node:util` `parseArgs`.

## Repositorio

- `src/`: `cli` (argumentos y salida), `sessions` (almacenamiento), `browser` (Edge, login y
  comprobación), `owa` (captura de la autenticación y llamadas a service.svc), `mail` (carpetas,
  filtros, lectura y borradores), `probe` (diagnóstico).
- `docs/`: este documento.
- `skill/owa-cli/SKILL.md`: la skill para agentes; se copia a la carpeta de skills del agente.

## Spike antes de construir (lo que decide el riesgo abierto)

1. `owa login` mínimo: Edge visible con perfil propio, login con «Mantener sesión: Sí», cerrar.
2. Reabrir en headless: ¿carga el buzón sin pedir login?
3. Observar las peticiones de OWA a `service.svc`: esquema de autenticación (sin volcar valores).
4. Un `FindItem` con `QueryString` AQS y un `CreateItem` de borrador desde la página.

Si 2 o 3 fallan, se revisa esta arquitectura antes de seguir.
