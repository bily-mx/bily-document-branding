// Base común: tokens de marca, rutas de logos y fuentes, y utilidades compartidas.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const RAIZ = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const T = JSON.parse(fs.readFileSync(path.join(RAIZ, 'tokens.json'), 'utf8'));

export const LOGO = {
  color: path.join(RAIZ, 'assets/logo/bily-logo.png'),
  blanco: path.join(RAIZ, 'assets/logo/bily-logo-blanco.png'),
  icono: path.join(RAIZ, 'assets/logo/bily-icono.png'),
};

// Archivo de cada familia tal como la nombra Office.
export const FUENTES = {
  'Geist': 'Geist-Regular.ttf',
  'Geist Medium': 'Geist-Medium.ttf',
  'Geist SemiBold': 'Geist-SemiBold.ttf',
  'Geist Mono': 'GeistMono-Regular.ttf',
  'Geist Mono Medium': 'GeistMono-Medium.ttf',
  'Geist Mono SemiBold': 'GeistMono-SemiBold.ttf',
};
export const rutaFuente = (familia) => path.join(RAIZ, 'assets/fonts', FUENTES[familia]);

/** Ancho y alto en pixeles de un PNG (lee la cabecera IHDR). */
export function medidasPng(ruta) {
  const b = fs.readFileSync(ruta);
  return { ancho: b.readUInt32BE(16), alto: b.readUInt32BE(20) };
}

/** Acento de sección: 'verde' | 'morado' | 'naranja' | 'rojo'. */
export function acento(nombre = 'verde') {
  const a = T.acento[nombre];
  if (!a) throw new Error(`Acento desconocido: "${nombre}". Usa verde, morado, naranja o rojo.`);
  return a;
}

/** Estado del semáforo fiscal: 'ok' | 'wa' | 'dup' | 'bad' | 'mute'. */
export function estado(tono = 'mute') {
  const e = T.estado[tono];
  if (!e) throw new Error(`Estado desconocido: "${tono}". Usa ok, wa, dup, bad o mute.`);
  return e;
}

/**
 * Parte un texto con marcas ligeras en tramos:
 *   **palabra clave** → { tipo: 'clave' }   `dato` → { tipo: 'dato' }   resto → { tipo: 'texto' }
 */
export function tramos(texto) {
  return String(texto ?? '')
    .split(/(\*\*[^*]+\*\*|`[^`]+`)/)
    .filter(Boolean)
    .map((t) => {
      if (t.startsWith('**')) return { tipo: 'clave', texto: t.slice(2, -2) };
      if (t.startsWith('`')) return { tipo: 'dato', texto: t.slice(1, -1) };
      return { tipo: 'texto', texto: t };
    });
}

export function asegurarCarpeta(archivo) {
  fs.mkdirSync(path.dirname(path.resolve(archivo)), { recursive: true });
}

/** Fecha larga en español de México: «27 de noviembre de 2026». */
export function fechaLarga(d = new Date()) {
  return new Intl.DateTimeFormat('es-MX', { day: 'numeric', month: 'long', year: 'numeric' }).format(d);
}
