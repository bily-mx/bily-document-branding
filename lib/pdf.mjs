// PDF con la identidad de Bily, desde HTML + la hoja de estilo de plantillas/documento.css.
// Usa el Chrome o Edge que ya esté instalado (no descarga navegador). Las fuentes quedan dentro del PDF.
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import puppeteer from 'puppeteer-core';
import { RAIZ, LOGO, acento as acc, asegurarCarpeta, fechaLarga } from './base.mjs';

const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function navegador() {
  const candidatos = [
    process.env.BILY_CHROME,
    'C:/Program Files/Google/Chrome/Application/chrome.exe',
    'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
    'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
    'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    '/usr/bin/google-chrome', '/usr/bin/chromium', '/usr/bin/chromium-browser', '/usr/bin/microsoft-edge',
  ].filter(Boolean);
  const ruta = candidatos.find((c) => fs.existsSync(c));
  if (!ruta) throw new Error('No encontré Chrome ni Edge. Instala uno o define la variable BILY_CHROME con la ruta del ejecutable.');
  return ruta;
}

function fuentesCss() {
  const cara = (familia, archivo, peso) =>
    `@font-face { font-family: '${familia}'; src: url('${pathToFileURL(path.join(RAIZ, 'assets/fonts', archivo)).href}') format('truetype'); font-weight: ${peso}; font-style: normal; }`;
  return [
    cara('Geist', 'Geist-Regular.ttf', 400), cara('Geist', 'Geist-Medium.ttf', 500), cara('Geist', 'Geist-SemiBold.ttf', 600), cara('Geist', 'Geist-Bold.ttf', 700),
    cara('Geist Mono', 'GeistMono-Regular.ttf', 400), cara('Geist Mono', 'GeistMono-Medium.ttf', 500), cara('Geist Mono', 'GeistMono-SemiBold.ttf', 600),
  ].join('\n');
}

/** HTML completo del documento. `cuerpo` es HTML con las clases de plantillas/documento.css. */
export function htmlDocumento({ titulo, tenue = '', tipo = '', lead = '', meta = [], acento = 'verde', portada = true, cuerpo = '' }) {
  if (!titulo) throw new Error('htmlDocumento: falta `titulo`.');
  acc(acento);
  const css = fs.readFileSync(path.join(RAIZ, 'plantillas/documento.css'), 'utf8').replace('__FUENTES__', fuentesCss());
  const datos = meta.length ? meta : [['Fecha', fechaLarga()]];
  const tapa = !portada ? '' : `
<section class="portada acc-${acento}">
  <img class="logo" src="${pathToFileURL(LOGO.color).href}" alt="Bily">
  <div class="centro">
    ${tipo ? `<p class="rotulo">${esc(tipo)}</p>` : ''}
    <h1>${esc(titulo)}${tenue ? ` <em>${esc(tenue)}</em>` : ''}</h1>
    ${lead ? `<p class="lead">${esc(lead)}</p>` : ''}
    <dl class="meta">${datos.map(([k, v]) => `<div><dt>${esc(k)}</dt><dd>${esc(v)}</dd></div>`).join('')}</dl>
  </div>
</section>`;
  return `<!doctype html>
<html lang="es-MX"><head><meta charset="utf-8"><title>${esc([titulo, tenue].filter(Boolean).join(' '))}</title>
<style>${css}</style></head>
<body class="acc-${acento}">${tapa}
<main>${cuerpo}</main>
</body></html>`;
}

/** Imprime un HTML (cadena) a PDF en hoja carta. */
export async function htmlAPdf(html, salida) {
  asegurarCarpeta(salida);
  const temporal = path.resolve(path.dirname(salida), `.${path.basename(salida, '.pdf')}.html`);
  fs.writeFileSync(temporal, html, 'utf8');
  const nav = await puppeteer.launch({ executablePath: navegador(), headless: true, args: ['--allow-file-access-from-files'] });
  try {
    const pagina = await nav.newPage();
    await pagina.goto(pathToFileURL(temporal).href, { waitUntil: 'load' });
    await pagina.evaluate(() => document.fonts.ready);
    await pagina.pdf({ path: salida, format: 'Letter', printBackground: true, preferCSSPageSize: true });
  } finally {
    await nav.close();
    fs.rmSync(temporal, { force: true });
  }
  return salida;
}

/** Atajo: arma el documento y lo guarda como PDF. Mismas opciones que htmlDocumento. */
export async function crearPdf(opciones, salida) {
  return htmlAPdf(htmlDocumento(opciones), salida);
}
