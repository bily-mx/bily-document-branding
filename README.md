# bily-brand

Skill de Claude para generar documentos de Word, PDF y presentaciones de PowerPoint con la identidad de Bily. Quien la instala le pide a Claude «hazme una propuesta de Bily para tal cliente» y el resultado sale con el mismo logo, colores, tipografía y composición que el de cualquier otra persona del equipo.

## Instalación

Requisitos: [Claude Code](https://claude.com/claude-code) (terminal, app de escritorio o IDE), Node.js 18 o más, y Chrome o Edge (para los PDF).

Windows (PowerShell):

```powershell
git clone https://github.com/bily-mx/bily-document-branding "$env:USERPROFILE\.claude\skills\bily-brand"
cd "$env:USERPROFILE\.claude\skills\bily-brand"
npm install
node scripts/instalar-fuentes.mjs
```

macOS o Linux:

```bash
git clone https://github.com/bily-mx/bily-document-branding ~/.claude/skills/bily-brand
cd ~/.claude/skills/bily-brand && npm install && node scripts/instalar-fuentes.mjs
```

Abre una sesión nueva de Claude y la skill ya está disponible en cualquier carpeta. Para actualizarla: `git pull` dentro de esa carpeta (y `npm install` si cambió `package.json`).

`instalar-fuentes.mjs` instala Geist y Geist Mono para tu usuario, sin permisos de administrador. Hace falta para abrir y editar las presentaciones; los Word y los PDF llevan la fuente dentro.

## Uso

En cualquier sesión de Claude Code:

- «Hazme una propuesta comercial de Bily en Word para Empresa X con estos puntos…»
- «Convierte estas notas en una presentación de Bily de 8 láminas.»
- «Genera el reporte mensual en PDF con la imagen de Bily.»

También se puede invocar directo con `/bily-brand`.

## Qué hay aquí

| Ruta | Contenido |
| --- | --- |
| `SKILL.md` | Instrucciones que lee Claude: flujo, API y reglas de marca. |
| `tokens.json` | Colores, fuentes y tamaños. Fuente única de la marca. |
| `lib/docx.mjs` | Documentos de Word (portada, títulos, tablas, cifras, avisos, estados). |
| `lib/pptx.mjs` | Presentaciones 16:9 (portada, sección, contenido, cifras, tabla, cierre). |
| `lib/pdf.mjs` + `plantillas/documento.css` | PDF desde HTML, en hoja carta. |
| `assets/` | Logos y fuentes (Geist, licencia OFL incluida). |
| `ejemplos/` | Un ejemplo de cada formato. `npm run ejemplos` los genera en `salida/`. |
| `scripts/` | Instalación de fuentes y conversión de Office a PDF (Windows). |

## Cambiar la marca

- **Un color o un tamaño:** `tokens.json`, y el mismo valor en `plantillas/documento.css` (la hoja de impresión lleva los colores escritos).
- **Una regla de uso** (qué se puede y qué no): `SKILL.md`.
- **Una pieza nueva** (otro tipo de lámina, un bloque de documento): `lib/`, con su ejemplo en `ejemplos/`.

Después de cambiar algo, corre `npm run ejemplos` y revisa los tres archivos de `salida/` antes de subirlo.

## Pendientes

- El logo está en PNG de 684 px (tomado del sitio). Sirve hasta unos 5 cm de ancho; para portadas grandes o impresión hace falta la versión vectorial del paquete de marca.
- Las reglas vienen del sistema de diseño del sitio (`DESIGN.md` de bilysite, noviembre de 2026). Si el sitio cambia de color o de tipografía, hay que traer el cambio aquí.
