# Huerta Sustentable — App web (Netlify + Excel SharePoint)

App de pedidos para Huerta Sustentable. La **base de datos es el Excel de SharePoint**
(`huerta_base_semanal.xlsx`), accedido vía Microsoft Graph desde funciones de Netlify.
Reusa la misma app de Azure AD del proyecto `RPA0045_VariablesDestileria`.

## Qué incluye
- **`/pedidos.html`** — formulario público del cliente (sin login). Lee el catálogo y precios
  del Excel y guarda cada pedido en la hoja **Pedidos**. Ofrece confirmar por WhatsApp.
- **`/admin.html`** — panel de gestión (con login). Pedidos en vivo, cambiar estado / precio
  final, **Detalle por verdura** y **Cosecha** calculados, edición de **precios y stock**, y
  **exportar el PDV** (.xls) para el ERP por rango de fechas.

```
netlify/functions/   _graph.js · catalogo · pedido · admin · estado · config · export-pdv
public/              index.html · pedidos.html · admin.html · app.js · admin.js
netlify.toml · package.json · .env (no se sube) · .env.example
```

## Variables de entorno
Se cargan en **Netlify → Site settings → Environment variables** (y en `.env` para local).
Son las mismas credenciales Graph de RPA0045; los IDs apuntan al Excel de Huerta:

| Variable | Valor |
|---|---|
| `GRAPH_TENANT_ID` | (igual que RPA0045) |
| `GRAPH_CLIENT_ID` | (igual que RPA0045) |
| `GRAPH_CLIENT_SECRET` | (igual que RPA0045) |
| `HUERTA_DRIVE_ID` | `b!S9qumy5pKEOo2DvM6Fus1W_IwICamuBIr0IkK-JDOgSrHZ7qDlu1SJea8KXVLQ06` |
| `HUERTA_ITEM_ID` | `01KH7MR26LGUQPL6VRBFA35J2537GOMEYL` |

> El archivo de Huerta está en el **mismo drive** que el Excel de destilería, así que la app
> de Azure ya tiene acceso (no hace falta registrar nada nuevo). Si algún día se mueve el Excel,
> reobtené los IDs con: `GET /shares/{u!<url-base64url>}/driveItem`.

El archivo `.env` local **no se sube a git** (está en `.gitignore`). Tiene el secreto real.

## Estructura del Excel (ya preparada)
El script de preparación ya agregó al Excel de SharePoint, de forma **aditiva** (sin tocar datos):
- **Detalle por Verdura** → columna **Stock** (TRUE/FALSE). Es el catálogo maestro (verduras +
  filas de Extras y Bolsón, que la app usa solo como precios).
- **Pedidos** → encabezados **Q=Estado, R=FechaPedido, S=IDPedido**. La app escribe A–O; la
  **P (Total) mantiene su fórmula**; Q/R/S las maneja la app.
- **Config** (clave/valor) → `PRECIO_BOLSON, PRECIO_ENVIO, PRECIO_SOPERA, PRECIO_ENSALADA,
  PRECIO_ESCABECHE, SS_SOPERA, SS_ENSALADA, SS_ESCABECHE, WHATSAPP, ALIAS, BANCO`.
- **Usuarios** (login admin) → `Mail | Nombre | Pss`.
- **Mapeo PDV** → mapeo de productos al ERP + valores fijos de cabecera.

### ⚠️ Cosas a completar
1. **Usuarios** — viene un usuario de ejemplo **`admin@huerta` / `huerta2026`**. Cambialo por los
   reales (una fila por persona). La contraseña se guarda en texto plano en el Excel; usá una
   distinta de las importantes.
2. **Mapeo PDV** — completá los códigos del ERP:
   - Tabla de productos (debajo de `Nombre app`): poné el **PRODUCTO** (código ERP) de cada
     verdura / extra / bolsón. `DESCRIPCIONITEM` y `TIPOITEM` son opcionales.
   - Bloque de campos fijos (debajo de `CAMPO`): completá los valores que van iguales en todas las
     filas: `CLIENTE` (código de cliente, ej. consumidor final), `COMPROBANTE`, `VENDEDOR`,
     `SUCURSAL`, `DEPOSITOORIGEN`, `DEPOSITODESTINO`, `DIMENSION*`, etc. `CONDICIONPAGO_TRANSFERENCIA`
     y `CONDICIONPAGO_EFECTIVO` permiten un valor distinto según el pago. `NUMERO_INICIAL` (opcional)
     es el primer número de pedido.
   - Mientras falte un código, el export pone `[FALTA CODIGO: <verdura>]` y el panel avisa cuántos
     ítems quedaron sin código.

## Probar localmente
```bash
npm install
npx netlify dev        # sirve public/ + funciones con las vars del .env
```
Abrí `http://localhost:8891/pedidos.html` y `/admin.html`.

## Deploy en Netlify
1. Subí el repo (o arrastrá la carpeta) a un sitio nuevo de Netlify.
2. Build: command vacío, **publish = `public`**, **functions = `netlify/functions`** (ya en `netlify.toml`).
3. Cargá las 5 variables de entorno de la tabla de arriba.
4. Deploy. El formulario queda en la raíz del sitio; el panel en `/admin.html`.

## Flujo semanal (sin cambios respecto al Excel)
- Los pedidos entran por la web y se cargan solos en la hoja **Pedidos**.
- El jueves, en el panel, mirás **Detalle por verdura** / **Cosecha** para saber cuánto cosechar.
- Para el ERP, en **Exportar PDV** elegís el rango de fechas y descargás el `.xls`.
- Fin de semana: archivás el Excel como hasta ahora (las hojas nuevas viajan con él).

## Notas
- La fórmula de Total (columna P) suma envío fijo **$4.000**; si cambia el costo de envío en Config,
  actualizá también esa constante en la fórmula del Excel si querés que la P del Excel coincida
  (el panel siempre usa el precio final que confirmás por pedido).
- Las contraseñas de **Usuarios** están en texto plano en el Excel (igual que en RPA0045). No uses
  contraseñas sensibles.
- Las páginas estáticas originales (`pedidos-huerta-sustentable.html`, `ventas-huerta-sustentable.html`)
  quedan en la raíz como referencia; **no** se publican (solo se publica `public/`).
