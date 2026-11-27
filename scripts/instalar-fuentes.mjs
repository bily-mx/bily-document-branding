// Instala Geist y Geist Mono para el usuario actual (sin permisos de administrador).
// Hace falta para editar o presentar los .pptx; los .docx y los PDF ya llevan la fuente dentro.
// Uso: node scripts/instalar-fuentes.mjs
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const origen = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../assets/fonts');
const archivos = fs.readdirSync(origen).filter((f) => f.endsWith('.ttf'));

const destinos = {
  win32: path.join(process.env.LOCALAPPDATA ?? '', 'Microsoft/Windows/Fonts'),
  darwin: path.join(os.homedir(), 'Library/Fonts'),
  linux: path.join(os.homedir(), '.local/share/fonts'),
};
const destino = destinos[process.platform];
if (!destino) throw new Error(`Sistema no soportado: ${process.platform}. Instala a mano los .ttf de ${origen}.`);
fs.mkdirSync(destino, { recursive: true });

for (const f of archivos) {
  const final = path.join(destino, f);
  fs.copyFileSync(path.join(origen, f), final);
  if (process.platform === 'win32') {
    // Windows solo muestra la fuente si queda registrada para el usuario.
    const nombre = `${f.replace('.ttf', '').replace('GeistMono', 'Geist Mono').replace('-', ' ')} (TrueType)`;
    execFileSync('reg', ['add', 'HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts', '/v', nombre, '/t', 'REG_SZ', '/d', final, '/f'], { stdio: 'ignore' });
  }
  console.log('Instalada:', f);
}
if (process.platform === 'win32') {
  // El registro solo surte efecto al volver a iniciar sesión: se cargan también en la sesión actual.
  const lista = archivos.map((f) => `'${path.join(destino, f).replace(/'/g, "''")}'`).join(',');
  const ps = `Add-Type -Namespace Bily -Name Gdi -MemberDefinition '[DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int AddFontResource(string ruta); [DllImport("user32.dll")] public static extern int SendNotifyMessage(IntPtr h, uint m, IntPtr w, IntPtr l);'; foreach ($f in @(${lista})) { [void][Bily.Gdi]::AddFontResource($f) }; [void][Bily.Gdi]::SendNotifyMessage([IntPtr]0xffff, 0x001D, [IntPtr]::Zero, [IntPtr]::Zero)`;
  try { execFileSync('powershell', ['-NoProfile', '-Command', ps], { stdio: 'ignore' }); } catch { console.log('No pude cargarlas en esta sesión: cierra sesión en Windows y vuelve a entrar.'); }
}
if (process.platform === 'linux') { try { execFileSync('fc-cache', ['-f'], { stdio: 'ignore' }); } catch { /* sin fontconfig */ } }
console.log(`\nListo: ${archivos.length} fuentes en ${destino}\nCierra y vuelve a abrir Word y PowerPoint para que las vean.`);
