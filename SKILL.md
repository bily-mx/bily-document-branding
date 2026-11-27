---
name: bily-brand
description: Genera documentos de Word (.docx), PDF y presentaciones (.pptx) con la identidad visual de Bily, la plataforma de control fiscal de CFDI de ProFact. Úsala siempre que pidan un documento, reporte, propuesta, cotización, carta, manual, one-pager, PDF, presentación, deck o láminas "de Bily", "con la imagen de Bily", "con nuestra marca" o para un cliente de Bily, aunque no mencionen la marca de forma explícita si el contenido es de Bily o ProFact.
---

# Bily: documentos con una sola imagen

Esta skill trae la marca ya resuelta en código: colores, tipografía, logo, portadas, tablas y láminas. Tu trabajo es el contenido. **No armes el estilo a mano** ni uses otra skill de docx, pptx o pdf para estos entregables: usa las librerías de `lib/`, que son las que garantizan que todos los documentos del equipo se vean igual.

## Flujo

1. **Elige el formato.** Si no lo dijeron: texto que se va a editar → Word; algo para enviar o imprimir → PDF; algo para exponer → presentación.
2. **Prepara la skill** (solo la primera vez): si no existe `node_modules/` en la carpeta de esta skill, corre `npm install` ahí. Requiere Node 18 o más.
3. **Escribe un script `.mjs`** en la carpeta de trabajo del usuario que importe la librería por su ruta absoluta (en Windows, como URL: `file:///C:/Users/.../bily-brand/lib/docx.mjs`). Copia la forma de `ejemplos/`.
4. **Córrelo con `node`** y guarda el resultado donde el usuario lo pidió (por defecto, su carpeta actual).
5. **Míralo antes de entregar.** Un PDF se lee directo. Un `.docx` o `.pptx` se convierte en Windows con `powershell -ExecutionPolicy Bypass -File scripts/office-a-pdf.ps1 <archivo>` y se revisa el PDF. Busca texto cortado o encimado, títulos de más de dos renglones y tablas que se salen.
6. Borra el script si el usuario no lo quiere conservar.

## Word: `lib/docx.mjs`

```js
import { crearDocumento } from 'file:///RUTA/bily-brand/lib/docx.mjs';

const doc = crearDocumento({
  tipo: 'Propuesta comercial',          // rótulo de la portada
  titulo: 'Control fiscal continuo',    // primera parte, en tinta
  tenue: 'para tu operación',           // segunda parte, en el acento
  lead: 'Una frase que resume el documento.',
  meta: [['Fecha', '27 de noviembre de 2026'], ['Preparado para', 'Cliente']],  // hasta 3
  acento: 'verde',                      // verde | morado | naranja | rojo
});
doc.rotulo('resumen').h1('Tus cifras coinciden', 'con las del SAT').lead('…').p('Texto con **palabra clave** y `DATO123`.');
await doc.guardar('C:/ruta/salida.docx');
```

Bloques (todos encadenables): `rotulo(texto)`, `h1(titulo, tenue, { nuevaPagina })`, `h2`, `h3`, `lead`, `p`, `nota`, `lista([...])`, `numerada([...])`, `tabla({ columnas, filas, total })`, `cifras([{ valor, etiqueta, destacar }])`, `aviso(texto)`, `estado(tono, texto, detalle)`, `imagen(ruta, { anchoCm, pie })`, `salto()`, `acento(nombre)` para cambiar el acento de lo que sigue, y `crudo(...)` para elementos de `docx` hechos a mano.

Las fuentes van incrustadas en el `.docx`: se ve igual aunque quien lo abra no tenga Geist.

## Presentación: `lib/pptx.mjs`

```js
import { crearPresentacion } from 'file:///RUTA/bily-brand/lib/pptx.mjs';

const p = crearPresentacion({ titulo: 'Control fiscal continuo' });
p.portada({ rotulo: 'Presentación comercial', titulo: 'Control fiscal continuo', tenue: 'para tus CFDI', pie: 'Cliente · Noviembre 2026' });
p.contenido({ rotulo: 'el problema', acento: 'rojo', titulo: 'Lo que no revisas', tenue: 'se convierte en riesgo', lead: '…', puntos: ['**Idea.** Detalle.'] });
p.cierre({ titulo: 'Bily revisa cada CFDI', tenue: 'para que nada se te escape', accion: 'Solicitar demostración', contacto: ['bily.mx'] });
await p.guardar('C:/ruta/salida.pptx');
```

Láminas: `portada`, `seccion` (banda oscura), `contenido` (hasta 5 puntos), `dosColumnas` (`puntos` + `imagen` o `derecha`), `cifras` (hasta 4), `tabla` (hasta 8 filas), `lienzo` (encabezado de marca y área libre: devuelve `{ lamina, y, alto }` para dibujar con pptxgenjs) y `cierre`. Todas las de contenido aceptan `rotulo`, `titulo`, `tenue`, `lead`, `acento` y `nota`.

**PowerPoint no incrusta fuentes.** Para que el `.pptx` se vea bien hace falta Geist instalada: `node scripts/instalar-fuentes.mjs` (una vez por equipo; pide permiso al usuario antes de instalar). Si la presentación sale de la empresa, entrega también el PDF.

## PDF: `lib/pdf.mjs`

```js
import { crearPdf } from 'file:///RUTA/bily-brand/lib/pdf.mjs';

await crearPdf({ tipo, titulo, tenue, lead, meta, acento, cuerpo: `<section class="acc-rojo">…</section>` }, 'C:/ruta/salida.pdf');
```

`cuerpo` es HTML con las clases de `plantillas/documento.css` (léela si necesitas algo más que lo de `ejemplos/pdf.mjs`): `.rotulo`, `h1` con `<em>` para la segunda parte, `.lead`, `.nota`, `.cifras`, tablas con `.num`, `.pill.ok|wa|dup|bad`, `.aviso`, `.sello`, `.dato`, `.salto`, y `.acc-verde|morado|naranja|rojo` en cada `<section>`. Usa el Chrome o Edge instalado. No agregues `<style>` propio salvo para una pieza que la hoja no cubra, y entonces con los colores de `tokens.json`.

## Marcas dentro del texto

En `p`, `lead`, `nota`, listas, puntos y celdas: `**palabra clave**` (seminegrita en tinta, de 1 a 3 por párrafo) y `` `dato` `` (mono: RFC, folios, UUID, importes).

Tablas: `columnas: [{ titulo, ancho, tipo }]` con `tipo` `'texto'`, `'num'` (mono, a la derecha), `'dato'` (mono) o `'estado'`; una celda de estado es `{ estado: 'ok' | 'wa' | 'dup' | 'bad' | 'mute', texto }`.

## Reglas de marca

Las librerías ya aplican tamaños, colores y márgenes. Estas reglas son las decisiones que te tocan a ti:

- **Títulos a dos tonos.** Todo título grande son dos partes: `titulo` en tinta y `tenue` en el acento. Que se lean como una sola frase («Tus cifras coinciden» + «con las del SAT»). Máximo dos renglones.
- **Una idea por sección o lámina.** Título, una frase y lo visual (cifra, tabla, lista). Si una lámina necesita más de cinco puntos, son dos láminas.
- **Diversidad de acentos.** Cada sección usa un acento distinto al de sus vecinas; un documento nunca va completo en un solo color. Portada y cierre siempre en verde. El acento no significa nada: es ritmo.
- **Los cuatro colores como estado significan algo.** Verde = resuelto, naranja = pendiente, morado = por revisar, rojo = con riesgo, gris = vigente. En pastillas y sellos solo se usan con ese significado, nunca como decoración ni para categorías.
- **La mono es para datos.** RFC, folios, UUID, importes, fechas de corte, cifras. Nunca títulos ni párrafos (excepción ya resuelta: rótulos y avisos).
- **Fondo blanco.** Nada de degradados, rellenos de color en renglones o tarjetas, ni sombras decorativas. El color va en el texto, el punto o la pastilla. La banda oscura (`seccion`) se usa una o dos veces por presentación.
- **Sin etiquetas en mayúsculas** sobre los títulos; el rótulo va en minúsculas.
- **Sin íconos decorativos, emojis ni imágenes de stock** (personas, edificios, escudos, palomitas gigantes). Lo visual es el dato o una vista real del producto.
- **El logo no se toca.** No se deforma, no se recolorea y no se le agregan los cuatro puntos sueltos como adorno. Sobre fondo oscuro, la versión blanca. Es un PNG de 684 px: no lo uses a más de 5 cm de ancho.
- **Verde de marca (#16C79A) nunca como texto** sobre blanco: para texto verde usa el acento (ya lo hacen las librerías).

## Contenido

- Español de México, lenguaje fiscal preciso. «IVA no acreditable», no «no deducible»; «Control Fiscal Integrado»; lista 69-B y 49 Bis.
- **No inventes** clientes, testimonios, certificaciones, precios ni cifras. Trayectoria confirmada: «18+ años» de experiencia de la empresa (no de Bily) y «100M+ CFDI». Si falta un dato, pregunta o deja el hueco señalado.
- Los datos de ejemplo son ficticios y se rotulan como ilustrativos (en `meta`, en una `nota` o en un `aviso`).
- La sincronización con el SAT es a día vencido: no digas «en tiempo real».
- Sin raya larga (—) en los textos: usa coma, dos puntos o punto.

## Si algo no está cubierto

Primero intenta componerlo con los bloques que hay. Si de verdad falta una pieza (una gráfica, un diagrama), dibújala con `lienzo` o `crudo` usando solo los colores y fuentes de `tokens.json`: líneas finas grises, color únicamente en lo que importa. Si la pieza se va a repetir, propón al usuario agregarla a la librería en vez de resolverla en cada documento.
