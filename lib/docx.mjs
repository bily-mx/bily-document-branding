// Documentos de Word con la identidad de Bily.
// Uso: crearDocumento({...}) → agregar bloques → await guardar(ruta).
// Las fuentes Geist van incrustadas en el archivo: se ve igual aunque quien lo abra no las tenga.
import fs from 'node:fs';
import {
  AlignmentType, BorderStyle, CharacterSet, Document, Footer, Header, HeadingLevel, ImageRun,
  LevelFormat, Packer, PageBreak, PageNumber, Paragraph, ShadingType, TabStopType, Table,
  TableCell, TableLayoutType, TableRow, TextRun, VerticalAlign, WidthType,
} from 'docx';
import {
  T, LOGO, acento as acc, asegurarCarpeta, estado as est, fechaLarga, medidasPng, rutaFuente, tramos,
} from './base.mjs';

const C = T.color;
const F = T.fuente;
const PT = T.documento.pt;
const cm = (n) => Math.round((n / 2.54) * 1440); // centímetros → twips
const px = (n) => Math.round((n / 2.54) * 96); // centímetros → pixeles de imagen
const M = T.documento.margenCm;
const ANCHO = 12240 - 2 * cm(M.lados); // ancho útil de la hoja carta, en twips

const SIN = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const raya = (color = C.linea, size = 6) => ({ style: BorderStyle.SINGLE, size, color });
const SIN_BORDES = { top: SIN, bottom: SIN, left: SIN, right: SIN };

/** Un tramo de texto. `pt` en puntos; `fuente` es el nombre de familia de Office. */
function tramo(texto, { fuente = F.texto, pt = PT.cuerpo, color = C.tinta2, espaciado, salto } = {}) {
  return new TextRun({ text: texto, font: fuente, size: Math.round(pt * 2), color, characterSpacing: espaciado, break: salto });
}

/** Texto con marcas: **palabra clave** (seminegrita en tinta) y `dato` (mono). */
function enLinea(texto, base = {}) {
  const pt = base.pt ?? PT.cuerpo;
  return tramos(texto).map((t) => {
    if (t.tipo === 'clave') return tramo(t.texto, { ...base, fuente: F.titulo, color: base.colorClave ?? C.tinta });
    if (t.tipo === 'dato') return tramo(t.texto, { ...base, fuente: F.datoMedio, pt: pt - 1, color: base.colorClave ?? C.tinta });
    return tramo(t.texto, base);
  });
}

/** Título a dos tonos: la primera parte en tinta y la segunda en el acento suave. Solo a 18 pt o más. */
function dosTonos(titulo, tenue, { pt, acento, espaciado }) {
  const base = { fuente: F.titulo, pt, espaciado };
  const partes = [tramo(titulo, { ...base, color: C.tinta })];
  if (tenue) partes.push(tramo(' ' + tenue, { ...base, color: acento ? acc(acento).suave : C.tinta5 }));
  return partes;
}

function imagen(ruta, anchoCm) {
  const m = medidasPng(ruta);
  const w = px(anchoCm);
  return new ImageRun({ type: 'png', data: fs.readFileSync(ruta), transformation: { width: w, height: Math.round((w * m.alto) / m.ancho) } });
}

function celda(hijos, { ancho, bordes = SIN_BORDES, margen = { top: 90, bottom: 90, left: 0, right: 160 }, vertical = VerticalAlign.CENTER } = {}) {
  return new TableCell({
    children: Array.isArray(hijos) ? hijos : [hijos],
    width: { size: ancho, type: WidthType.DXA },
    borders: bordes,
    margins: margen,
    verticalAlign: vertical,
  });
}

function pastilla(tono, texto) {
  const e = est(tono);
  return new TextRun({
    text: ` • ${texto ?? e.nombre}  `,
    font: F.titulo, size: Math.round(PT.rotulo * 2), color: e.texto,
    shading: { type: ShadingType.CLEAR, fill: e.fondo, color: 'auto' },
  });
}

const parrafo = (hijos, opciones = {}) => new Paragraph({ children: hijos, ...opciones });

export function crearDocumento({
  titulo,
  tenue = '',
  tipo = '',
  lead = '',
  meta = [],
  acento = 'verde',
  autor = 'Bily',
  portada = true,
  encabezado,
} = {}) {
  if (!titulo) throw new Error('crearDocumento: falta `titulo`.');
  const bloques = [];
  let acentoActual = acento;

  const api = {
    /** Cambia el acento de las secciones que siguen (verde | morado | naranja | rojo). */
    acento(nombre) { acc(nombre); acentoActual = nombre; return api; },

    /** Rótulo de sección: punto de acento + texto mono en minúsculas. Va sobre un h1. */
    rotulo(texto, { acento: a = acentoActual } = {}) {
      bloques.push(parrafo([
        tramo('•  ', { fuente: F.titulo, pt: PT.rotulo + 3, color: acc(a).punto }),
        tramo(texto.toLowerCase(), { fuente: F.datoMedio, pt: PT.rotulo, color: acc(a).texto }),
      ], { spacing: { before: 480, after: 60 }, keepNext: true }));
      return api;
    },

    /** Título de capítulo, a dos tonos. `nuevaPagina` lo manda a hoja nueva. */
    h1(texto, tenue2 = '', { acento: a = acentoActual, nuevaPagina = false } = {}) {
      bloques.push(parrafo(dosTonos(texto, tenue2, { pt: PT.h1, acento: a, espaciado: -12 }), {
        heading: HeadingLevel.HEADING_1, pageBreakBefore: nuevaPagina,
      }));
      return api;
    },
    h2(texto) { bloques.push(parrafo([tramo(texto, { fuente: F.titulo, pt: PT.h2, color: C.tinta, espaciado: -6 })], { heading: HeadingLevel.HEADING_2 })); return api; },
    h3(texto) { bloques.push(parrafo([tramo(texto, { fuente: F.titulo, pt: PT.h3, color: C.tinta })], { heading: HeadingLevel.HEADING_3 })); return api; },

    /** Frase de entrada bajo un título: una sola idea. */
    lead(texto) {
      bloques.push(parrafo(enLinea(texto, { pt: PT.lead, color: C.tinta2 }), { spacing: { after: 240, line: 336 } }));
      return api;
    },
    p(texto) { bloques.push(parrafo(enLinea(texto))); return api; },

    /** Nota de valor: texto chico en gris que abre con la idea en negritas. */
    nota(texto) {
      bloques.push(parrafo(enLinea(texto, { pt: PT.chico, color: C.tinta3 }), { spacing: { after: 140, line: 312 } }));
      return api;
    },

    lista(items) {
      for (const it of items) bloques.push(parrafo(enLinea(it), { numbering: { reference: 'bily-vineta', level: 0 }, spacing: { after: 80 } }));
      bloques.push(parrafo([], { spacing: { after: 60 } }));
      return api;
    },
    numerada(items) {
      const instancia = (api._n = (api._n ?? 0) + 1);
      for (const it of items) bloques.push(parrafo(enLinea(it), { numbering: { reference: 'bily-numero', level: 0, instance: instancia }, spacing: { after: 80 } }));
      bloques.push(parrafo([], { spacing: { after: 60 } }));
      return api;
    },

    /**
     * Tabla de líneas finas, sin rellenos.
     * columnas: [{ titulo, ancho (proporción), tipo: 'texto' | 'num' | 'dato' | 'estado' }]
     * filas: arreglos de celdas. Celda de estado: { estado: 'ok' | 'wa' | 'dup' | 'bad' | 'mute', texto }.
     */
    tabla({ columnas, filas, total }) {
      const suma = columnas.reduce((s, c) => s + (c.ancho ?? 1), 0);
      const anchos = columnas.map((c) => Math.floor((ANCHO * (c.ancho ?? 1)) / suma));
      const derecha = (c) => c.tipo === 'num';
      const fila = (celdas, { cabeza = false, ultima = false, esTotal = false } = {}) => new TableRow({
        tableHeader: cabeza, cantSplit: true,
        children: celdas.map((v, i) => {
          const col = columnas[i];
          let hijos;
          if (cabeza) hijos = [tramo(String(v), { fuente: F.medio, pt: PT.rotulo, color: C.tinta3 })];
          else if (v && typeof v === 'object' && v.estado) hijos = [pastilla(v.estado, v.texto)];
          else if (col.tipo === 'num' || col.tipo === 'dato') hijos = [tramo(String(v ?? ''), { fuente: esTotal ? F.datoFuerte : F.datoMedio, pt: PT.dato, color: C.tinta })];
          else hijos = enLinea(String(v ?? ''), { pt: PT.cuerpo - 0.5, color: esTotal || i === 0 ? C.tinta : C.tinta2, ...(esTotal ? { fuente: F.titulo } : {}) });
          return celda(parrafo(hijos, { alignment: derecha(col) ? AlignmentType.RIGHT : AlignmentType.LEFT, spacing: { after: 0, line: 276 } }), {
            ancho: anchos[i],
            margen: { top: cabeza ? 60 : 100, bottom: cabeza ? 80 : 100, left: i === 0 ? 0 : 80, right: i === columnas.length - 1 ? 0 : 80 },
            bordes: { ...SIN_BORDES, top: esTotal ? raya(C.tinta, 8) : SIN, bottom: cabeza ? raya(C.lineaFuerte, 8) : ultima || esTotal ? SIN : raya(C.linea, 6) },
          });
        }),
      });
      const filasTabla = [fila(columnas.map((c) => c.titulo), { cabeza: true })];
      filas.forEach((f, i) => filasTabla.push(fila(f, { ultima: i === filas.length - 1 && !total })));
      if (total) filasTabla.push(fila(total, { esTotal: true }));
      bloques.push(new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: anchos, layout: TableLayoutType.FIXED, rows: filasTabla, borders: { ...SIN_BORDES, insideHorizontal: SIN, insideVertical: SIN } }));
      bloques.push(parrafo([], { spacing: { after: 160 } }));
      return api;
    },

    /** Cifras protagonistas: hasta cuatro, en mono, con su etiqueta. `destacar` pinta el valor en el acento. */
    cifras(items, { acento: a = acentoActual } = {}) {
      const w = Math.floor(ANCHO / items.length);
      // La cifra nunca se parte en dos renglones: si no cabe, baja el tamaño de toda la fila.
      const largo = Math.max(...items.map((it) => String(it.valor).length));
      const ptCifra = Math.max(PT.h3, Math.min(PT.cifra, Math.floor(((w - 200) / 20 / (0.6 * largo)) * 2) / 2));
      bloques.push(new Table({
        width: { size: ANCHO, type: WidthType.DXA }, columnWidths: items.map(() => w), layout: TableLayoutType.FIXED,
        borders: { ...SIN_BORDES, insideHorizontal: SIN, insideVertical: SIN },
        rows: [new TableRow({ cantSplit: true, children: items.map((it) => celda([
          parrafo([tramo(String(it.valor), { fuente: F.datoMedio, pt: ptCifra, color: it.destacar ? acc(a).texto : C.tinta, espaciado: -10 })], { spacing: { after: 40, line: 240 } }),
          parrafo([tramo(it.etiqueta, { pt: PT.chico, color: C.tinta3 })], { spacing: { after: 0, line: 288 } }),
        ], { ancho: w, vertical: VerticalAlign.TOP, margen: { top: 160, bottom: 120, left: 0, right: 200 }, bordes: { ...SIN_BORDES, top: raya(C.lineaFuerte, 8) } })) })],
      }));
      bloques.push(parrafo([], { spacing: { after: 160 } }));
      return api;
    },

    /** Aviso: papel, borde fino y raya de acento a la izquierda; texto en mono. A todo el ancho. */
    aviso(texto, { acento: a = acentoActual } = {}) {
      bloques.push(new Table({
        width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [ANCHO], layout: TableLayoutType.FIXED,
        rows: [new TableRow({ cantSplit: true, children: [celda(
          parrafo(enLinea(texto, { fuente: F.datoMedio, pt: PT.chico, color: C.tinta2 }), { spacing: { after: 0, line: 312 } }),
          { ancho: ANCHO, margen: { top: 160, bottom: 160, left: 240, right: 240 }, bordes: { top: raya(), bottom: raya(), right: raya(), left: raya(acc(a).punto, 18) } },
        )] })],
      }));
      bloques.push(parrafo([], { spacing: { after: 160 } }));
      return api;
    },

    /** Renglón con un estado del semáforo fiscal y un texto al lado. */
    estado(tono, texto, detalle = '') {
      const hijos = [pastilla(tono, texto)];
      if (detalle) hijos.push(tramo('   '), ...enLinea(detalle));
      bloques.push(parrafo(hijos));
      return api;
    },

    /** Imagen (PNG) al ancho indicado en centímetros, con pie opcional. */
    imagen(ruta, { anchoCm = 16, pie = '' } = {}) {
      bloques.push(parrafo([imagen(ruta, Math.min(anchoCm, (ANCHO / 1440) * 2.54))], { spacing: { before: 120, after: pie ? 80 : 240 } }));
      if (pie) bloques.push(parrafo([tramo(pie, { pt: PT.rotulo, color: C.tinta3 })], { spacing: { after: 240 } }));
      return api;
    },

    salto() { bloques.push(parrafo([new PageBreak()])); return api; },

    /** Para casos especiales: agrega párrafos o tablas de `docx` hechos a mano. */
    crudo(...elementos) { bloques.push(...elementos); return api; },

    construir() {
      const hijos = [];
      if (portada) {
        hijos.push(parrafo([imagen(LOGO.color, 3.4)], { spacing: { after: 0 } }));
        hijos.push(parrafo([], { spacing: { before: cm(6.2), after: 0 } }));
        if (tipo) hijos.push(parrafo([
          tramo('•  ', { fuente: F.titulo, pt: PT.rotulo + 3, color: acc(acento).punto }),
          tramo(tipo.toLowerCase(), { fuente: F.datoMedio, pt: PT.rotulo + 0.5, color: acc(acento).texto }),
        ], { spacing: { after: 160 } }));
        hijos.push(parrafo(dosTonos(titulo, tenue, { pt: PT.portada, acento, espaciado: -20 }), { style: 'BilyPortada' }));
        if (lead) hijos.push(parrafo(enLinea(lead, { pt: PT.lead + 1, color: C.tinta2 }), { spacing: { before: 120, after: 0, line: 336 }, indent: { right: cm(3) } }));
        const datos = meta.length ? meta : [['Fecha', fechaLarga()]];
        const w = Math.floor(ANCHO / Math.max(datos.length, 3));
        hijos.push(parrafo([], { spacing: { before: cm(2.2), after: 0 } }));
        hijos.push(new Table({
          width: { size: w * datos.length, type: WidthType.DXA }, columnWidths: datos.map(() => w), layout: TableLayoutType.FIXED,
          borders: { ...SIN_BORDES, insideHorizontal: SIN, insideVertical: SIN },
          rows: [new TableRow({ children: datos.map(([k, v]) => celda([
            parrafo([tramo(k, { fuente: F.medio, pt: PT.rotulo, color: C.tinta3 })], { spacing: { after: 40 } }),
            parrafo([tramo(v, { fuente: F.datoMedio, pt: PT.dato, color: C.tinta })], { spacing: { after: 0 } }),
          ], { ancho: w, vertical: VerticalAlign.TOP, margen: { top: 140, bottom: 0, left: 0, right: 200 }, bordes: { ...SIN_BORDES, top: raya(C.lineaFuerte, 8) } })) })],
        }));
        hijos.push(parrafo([new PageBreak()]));
      }
      hijos.push(...bloques);

      const pie = new Footer({ children: [parrafo([
        tramo(T.marca.sitio, { fuente: F.datoMedio, pt: PT.minimo, color: C.tinta3 }),
        new TextRun({ children: ['\t', PageNumber.CURRENT], font: F.datoMedio, size: PT.minimo * 2, color: C.tinta3 }),
      ], { tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }], border: { top: { ...raya(), space: 8 } }, spacing: { after: 0 } })] });
      const cabeza = new Header({ children: [parrafo([
        imagen(LOGO.icono, 0.5),
        tramo(`\t${encabezado ?? [titulo, tenue].filter(Boolean).join(' ')}`, { fuente: F.medio, pt: PT.minimo, color: C.tinta3 }),
      ], { tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }], spacing: { after: 0 } })] });
      const vacio = () => parrafo([], { spacing: { after: 0 } });

      const incrustar = (nombre, extra = {}) => ({ name: nombre, data: fs.readFileSync(rutaFuente(nombre)), characterSet: CharacterSet.ANSI, ...extra });

      return new Document({
        creator: autor, title: [titulo, tenue].filter(Boolean).join(' '), description: tipo || undefined,
        fonts: [
          incrustar(F.texto), incrustar(F.medio), incrustar(F.titulo), incrustar(F.dato), incrustar(F.datoMedio), incrustar(F.datoFuerte),
        ],
        styles: {
          default: {
            document: { run: { font: F.texto, size: PT.cuerpo * 2, color: C.tinta2 }, paragraph: { spacing: { after: 160, line: 348 } } },
            heading1: { run: { font: F.titulo, size: PT.h1 * 2, color: C.tinta, characterSpacing: -12 }, paragraph: { spacing: { before: 120, after: 200, line: 259 }, keepNext: true } },
            heading2: { run: { font: F.titulo, size: PT.h2 * 2, color: C.tinta, characterSpacing: -6 }, paragraph: { spacing: { before: 400, after: 140, line: 276 }, keepNext: true } },
            heading3: { run: { font: F.titulo, size: PT.h3 * 2, color: C.tinta }, paragraph: { spacing: { before: 280, after: 100, line: 288 }, keepNext: true } },
          },
          paragraphStyles: [{
            id: 'BilyPortada', name: 'Bily Portada', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: F.titulo, size: PT.portada * 2, color: C.tinta, characterSpacing: -20 },
            paragraph: { spacing: { after: 200, line: 250 }, indent: { right: cm(1.5) } },
          }],
        },
        numbering: { config: [
          { reference: 'bily-vineta', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { run: { font: F.texto, color: C.tinta4 }, paragraph: { indent: { left: 340, hanging: 240 } } } }] },
          { reference: 'bily-numero', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1', alignment: AlignmentType.LEFT, style: { run: { font: F.datoMedio, color: C.tinta3, size: PT.dato * 2 }, paragraph: { indent: { left: 400, hanging: 400 } } } }] },
        ] },
        sections: [{
          properties: {
            titlePage: portada,
            page: { size: { width: 12240, height: 15840 }, margin: { top: cm(M.arriba), bottom: cm(M.abajo), left: cm(M.lados), right: cm(M.lados), header: cm(1.1), footer: cm(1) } },
          },
          headers: { default: cabeza, first: new Header({ children: [vacio()] }) },
          footers: { default: pie, first: new Footer({ children: [vacio()] }) },
          children: hijos,
        }],
      });
    },

    async guardar(ruta) {
      asegurarCarpeta(ruta);
      fs.writeFileSync(ruta, await Packer.toBuffer(api.construir()));
      return ruta;
    },
  };
  return api;
}
