// Presentaciones de PowerPoint con la identidad de Bily (16:9).
// Uso: crearPresentacion({...}) → portada / seccion / contenido / cifras / tabla / dosColumnas / cierre → await guardar(ruta).
// PowerPoint no incrusta fuentes: quien abra el .pptx necesita Geist instalada (scripts/instalar-fuentes.mjs)
// o recibe el PDF. Sin Geist, PowerPoint la sustituye y el diseño cambia.
import modulo from 'pptxgenjs';
import { T, LOGO, acento as acc, asegurarCarpeta, estado as est, medidasPng, tramos } from './base.mjs';

const PptxGenJS = modulo.default ?? modulo;
const C = T.color;
const F = T.fuente;
const PT = T.presentacion.pt;
const { ancho: W, alto: H, margen: M } = T.presentacion.pulgadas;
const UTIL = W - 2 * M;

const altoLogo = (ruta, w) => { const m = medidasPng(ruta); return (w * m.alto) / m.ancho; };

/** Texto con marcas → tramos de pptxgenjs. **clave** en seminegrita y tinta; `dato` en mono. */
function enLinea(texto, { color = C.tinta2, colorClave = C.tinta, pt } = {}) {
  return tramos(texto).map((t) => {
    if (t.tipo === 'clave') return { text: t.texto, options: { fontFace: F.titulo, color: colorClave } };
    if (t.tipo === 'dato') return { text: t.texto, options: { fontFace: F.datoMedio, color: colorClave, fontSize: pt ? pt - 1 : undefined } };
    return { text: t.texto, options: { fontFace: F.texto, color } };
  });
}

function dosTonos(titulo, tenue, colorTenue, colorBase = C.tinta) {
  const partes = [{ text: titulo, options: { color: colorBase } }];
  if (tenue) partes.push({ text: ' ' + tenue, options: { color: colorTenue } });
  return partes;
}

export function crearPresentacion({ titulo = 'Presentación', autor = 'Bily', acento = 'verde' } = {}) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.title = titulo;
  pptx.author = autor;
  pptx.company = 'Bily';
  pptx.theme = { headFontFace: F.titulo, bodyFontFace: F.texto };

  const numero = { x: W - M - 1, y: H - 0.62, w: 1, h: 0.3, align: 'right', fontFace: F.datoMedio, fontSize: 9, color: C.tinta3 };
  const hIcono = 0.26;
  const wIcono = hIcono * (medidasPng(LOGO.icono).ancho / medidasPng(LOGO.icono).alto);
  pptx.defineSlideMaster({
    title: 'BILY_PAPEL', background: { color: C.papel },
    objects: [
      { image: { path: LOGO.icono, x: M, y: H - 0.6, w: wIcono, h: hIcono } },
      { text: { text: T.marca.sitio, options: { x: M + wIcono + 0.12, y: H - 0.62, w: 2, h: 0.3, fontFace: F.datoMedio, fontSize: 9, color: C.tinta3, margin: 0, valign: 'middle' } } },
    ],
    slideNumber: numero,
  });
  pptx.defineSlideMaster({ title: 'BILY_LIMPIA', background: { color: C.papel } });
  pptx.defineSlideMaster({ title: 'BILY_NOCHE', background: { color: C.noche }, slideNumber: { ...numero, color: C.sobreNoche2 } });

  let acentoActual = acento;

  /** Encabezado común: rótulo, título a dos tonos a la izquierda y frase a la derecha, alineados abajo. */
  function encabezado(s, { rotulo, titulo: t, tenue, lead, acento: a = acentoActual }) {
    const A = acc(a);
    if (rotulo) {
      s.addShape(pptx.ShapeType.ellipse, { x: M, y: 0.74, w: 0.11, h: 0.11, fill: { color: A.punto }, line: { color: A.punto, width: 0 } });
      s.addText(rotulo.toLowerCase(), { x: M + 0.22, y: 0.64, w: 6, h: 0.3, fontFace: F.datoMedio, fontSize: PT.rotulo, color: A.texto, margin: 0, valign: 'middle' });
    }
    const wT = lead ? UTIL * 0.54 : UTIL * 0.8;
    s.addText(dosTonos(t, tenue, A.suave), { x: M, y: 1.0, w: wT, h: 1.55, fontFace: F.titulo, fontSize: PT.titulo, charSpacing: -0.6, lineSpacingMultiple: 0.95, margin: 0, valign: 'bottom' });
    if (lead) s.addText(enLinea(lead, { pt: PT.lead }), { x: M + UTIL * 0.58, y: 1.0, w: UTIL * 0.42, h: 1.55, fontSize: PT.lead, lineSpacingMultiple: 1.25, margin: 0, valign: 'bottom' });
    return { y: 2.95, alto: H - 2.95 - 0.95, A };
  }

  const raya = (s, x, y, w, color = C.linea, grosor = 0.75) => s.addShape(pptx.ShapeType.line, { x, y, w, h: 0, line: { color, width: grosor } });

  function notaAlPie(s, nota) {
    if (nota) s.addText(enLinea(nota, { color: C.tinta3, pt: PT.chico }), { x: M, y: H - 1.45, w: UTIL * 0.7, h: 0.6, fontSize: PT.chico, lineSpacingMultiple: 1.2, margin: 0, valign: 'bottom' });
  }

  const api = {
    pptx,
    medidas: { W, H, M, UTIL },
    acento(nombre) { acc(nombre); acentoActual = nombre; return api; },

    /** Portada: logo arriba, titular a dos tonos abajo a la izquierda. `pie`: fecha, autor o cliente. */
    portada({ titulo: t, tenue = '', rotulo = '', pie = '', acento: a = acentoActual }) {
      const s = pptx.addSlide({ masterName: 'BILY_LIMPIA' });
      const A = acc(a);
      const wL = 1.55;
      s.addImage({ path: LOGO.color, x: M, y: 0.6, w: wL, h: altoLogo(LOGO.color, wL) });
      if (rotulo) {
        s.addShape(pptx.ShapeType.ellipse, { x: M, y: 3.42, w: 0.12, h: 0.12, fill: { color: A.punto }, line: { color: A.punto, width: 0 } });
        s.addText(rotulo.toLowerCase(), { x: M + 0.24, y: 3.32, w: 8, h: 0.32, fontFace: F.datoMedio, fontSize: PT.rotulo + 1, color: A.texto, margin: 0, valign: 'middle' });
      }
      s.addText(dosTonos(t, tenue, A.suave), { x: M, y: 3.8, w: UTIL * 0.82, h: 2.35, fontFace: F.titulo, fontSize: PT.portada, charSpacing: -1.2, lineSpacingMultiple: 0.95, margin: 0, valign: 'top' });
      if (pie) {
        raya(s, M, H - 0.95, UTIL, C.lineaFuerte);
        s.addText(pie, { x: M, y: H - 0.8, w: UTIL, h: 0.3, fontFace: F.datoMedio, fontSize: PT.rotulo, color: C.tinta3, margin: 0, valign: 'middle' });
      }
      return s;
    },

    /** Cambio de tema en banda oscura. Una o dos por presentación. */
    seccion({ titulo: t, tenue = '', rotulo = '' }) {
      const s = pptx.addSlide({ masterName: 'BILY_NOCHE' });
      const wL = 1.2;
      s.addImage({ path: LOGO.blanco, x: M, y: 0.6, w: wL, h: altoLogo(LOGO.blanco, wL) });
      if (rotulo) s.addText(rotulo.toLowerCase(), { x: M, y: 3.5, w: 8, h: 0.32, fontFace: F.datoMedio, fontSize: PT.rotulo + 1, color: C.sobreNoche2, margin: 0, valign: 'middle' });
      s.addText(dosTonos(t, tenue, C.sobreNoche2, C.sobreNoche), { x: M, y: 3.95, w: UTIL * 0.8, h: 2.3, fontFace: F.titulo, fontSize: PT.seccion, charSpacing: -1, lineSpacingMultiple: 0.95, margin: 0, valign: 'top' });
      return s;
    },

    /** Una idea con hasta cinco puntos, separados por líneas finas. Cada punto abre con su **idea**. */
    contenido({ puntos = [], nota = '', ...cabeza }) {
      const s = pptx.addSlide({ masterName: 'BILY_PAPEL' });
      const { y, alto } = encabezado(s, cabeza);
      if (puntos.length > 5) throw new Error('contenido: máximo cinco puntos por lámina (una idea por lámina). Divide en dos láminas.');
      const libre = alto - (nota ? 0.55 : 0);
      const h = Math.min(0.78, libre / Math.max(puntos.length, 1));
      puntos.forEach((p, i) => {
        raya(s, M, y + i * h, UTIL);
        s.addText(enLinea(p, { pt: PT.cuerpo }), { x: M, y: y + i * h, w: UTIL, h, fontSize: PT.cuerpo, margin: 0, valign: 'middle' });
      });
      notaAlPie(s, nota);
      return s;
    },

    /** Dividida: texto y puntos a la izquierda, imagen (o segunda columna de puntos) a la derecha. */
    dosColumnas({ puntos = [], imagen, derecha = [], nota = '', ...cabeza }) {
      const s = pptx.addSlide({ masterName: 'BILY_PAPEL' });
      const { y, alto } = encabezado(s, cabeza);
      const wCol = UTIL * 0.46;
      const xDer = M + UTIL * 0.54;
      const columna = (items, x) => {
        const h = Math.min(0.78, alto / Math.max(items.length, 1));
        items.forEach((p, i) => {
          raya(s, x, y + i * h, wCol);
          s.addText(enLinea(p, { pt: PT.cuerpo - 1 }), { x, y: y + i * h, w: wCol, h, fontSize: PT.cuerpo - 1, margin: 0, valign: 'middle' });
        });
      };
      columna(puntos, M);
      if (imagen) {
        const m = medidasPng(imagen);
        let w = wCol; let h = (w * m.alto) / m.ancho;
        if (h > alto) { h = alto; w = (h * m.ancho) / m.alto; }
        s.addImage({ path: imagen, x: xDer, y, w, h });
      } else columna(derecha, xDer);
      notaAlPie(s, nota);
      return s;
    },

    /** Cifras protagonistas en mono (hasta cuatro). `destacar` pinta la cifra en el acento de la lámina. */
    cifras({ cifras: items = [], nota = '', ...cabeza }) {
      const s = pptx.addSlide({ masterName: 'BILY_PAPEL' });
      const { y, A } = encabezado(s, cabeza);
      if (items.length > 4) throw new Error('cifras: máximo cuatro cifras por lámina.');
      const sep = 0.4;
      const w = (UTIL - sep * (items.length - 1)) / items.length;
      const largo = Math.max(...items.map((it) => String(it.valor).length));
      const pt = Math.max(18, Math.min(PT.cifra, Math.floor((w * 72) / (0.62 * largo))));
      items.forEach((it, i) => {
        const x = M + i * (w + sep);
        raya(s, x, y + 0.1, w, C.lineaFuerte);
        s.addText(String(it.valor), { x, y: y + 0.35, w, h: 0.95, fontFace: F.datoMedio, fontSize: pt, color: it.destacar ? A.texto : C.tinta, charSpacing: -0.8, margin: 0, valign: 'middle' });
        s.addText(it.etiqueta, { x, y: y + 1.35, w: w - 0.2, h: 0.8, fontFace: F.texto, fontSize: PT.chico + 1, color: C.tinta2, lineSpacingMultiple: 1.2, margin: 0, valign: 'top' });
      });
      notaAlPie(s, nota);
      return s;
    },

    /**
     * Tabla de líneas finas, sin rellenos. Máximo ocho filas por lámina.
     * columnas: [{ titulo, ancho (proporción), tipo: 'texto' | 'num' | 'dato' | 'estado' }]
     * Celda de estado: { estado: 'ok' | 'wa' | 'dup' | 'bad' | 'mute', texto }.
     */
    tabla({ columnas, filas, nota = '', ...cabeza }) {
      const s = pptx.addSlide({ masterName: 'BILY_PAPEL' });
      const { y } = encabezado(s, cabeza);
      if (filas.length > 8) throw new Error('tabla: máximo ocho filas por lámina. Divide la tabla o lleva el detalle a un documento.');
      const suma = columnas.reduce((t, c) => t + (c.ancho ?? 1), 0);
      const colW = columnas.map((c) => (UTIL * (c.ancho ?? 1)) / suma);
      const borde = (color) => [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.75, color }, { type: 'none' }];
      const al = (c) => (c.tipo === 'num' ? 'right' : 'left');
      const cabezaFila = columnas.map((c) => ({ text: c.titulo, options: { fontFace: F.medio, fontSize: PT.rotulo, color: C.tinta3, align: al(c), border: borde(C.lineaFuerte), valign: 'bottom' } }));
      const cuerpo = filas.map((f, n) => f.map((v, i) => {
        const col = columnas[i];
        const base = { align: al(col), border: n === filas.length - 1 ? borde(C.papel) : borde(C.linea), valign: 'middle' };
        if (v && typeof v === 'object' && v.estado) {
          const e = est(v.estado);
          return { text: [{ text: '●  ', options: { color: e.punto, fontSize: 8 } }, { text: v.texto ?? e.nombre, options: { color: e.texto, fontFace: F.titulo } }], options: { ...base, fontSize: PT.chico } };
        }
        if (col.tipo === 'num' || col.tipo === 'dato') return { text: String(v ?? ''), options: { ...base, fontFace: F.datoMedio, fontSize: PT.chico + 1, color: C.tinta } };
        return { text: String(v ?? ''), options: { ...base, fontFace: F.texto, fontSize: PT.chico + 2, color: i === 0 ? C.tinta : C.tinta2 } };
      }));
      s.addTable([cabezaFila, ...cuerpo], { x: M, y, w: UTIL, colW, rowH: 0.42, margin: [0.04, 0.08, 0.04, 0] });
      notaAlPie(s, nota);
      return s;
    },

    /** Lámina en blanco con el encabezado de la marca, para dibujar algo propio. Devuelve { lamina, y, alto }. */
    lienzo(cabeza) {
      const s = pptx.addSlide({ masterName: 'BILY_PAPEL' });
      const { y, alto } = encabezado(s, cabeza);
      return { lamina: s, y, alto };
    },

    /** Cierre: siempre en verde. El botón lleva el verde; los datos de contacto van en tinta. */
    cierre({ titulo: t, tenue = '', accion = '', contacto = [] }) {
      const s = pptx.addSlide({ masterName: 'BILY_LIMPIA' });
      const A = acc('verde');
      const wL = 1.2;
      s.addImage({ path: LOGO.color, x: M, y: 0.6, w: wL, h: altoLogo(LOGO.color, wL) });
      s.addText(dosTonos(t, tenue, A.suave), { x: M, y: 2.6, w: UTIL * 0.8, h: 2.1, fontFace: F.titulo, fontSize: PT.seccion, charSpacing: -1, lineSpacingMultiple: 0.95, margin: 0, valign: 'bottom' });
      let x = M;
      if (accion) {
        const w = Math.max(2.2, accion.length * 0.125 + 0.7);
        s.addText(accion, { x, y: 5.1, w, h: 0.56, shape: pptx.ShapeType.roundRect, rectRadius: 0.28, fill: { color: C.verde }, line: { color: C.verde, width: 0 }, fontFace: F.titulo, fontSize: 14, color: C.tinta, align: 'center', valign: 'middle', margin: 0 });
        x += w + 0.4;
      }
      if (contacto.length) s.addText(contacto.join('   ·   '), { x, y: 5.1, w: W - M - x, h: 0.56, fontFace: F.datoMedio, fontSize: PT.chico, color: C.tinta2, margin: 0, valign: 'middle' });
      return s;
    },

    async guardar(ruta) {
      asegurarCarpeta(ruta);
      await pptx.writeFile({ fileName: ruta });
      return ruta;
    },
  };
  return api;
}
