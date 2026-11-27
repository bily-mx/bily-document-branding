// Ejemplo de PDF desde HTML. Datos 100% ficticios, rotulados como ilustrativos.
import { fileURLToPath } from 'node:url';
import { crearPdf } from '../lib/pdf.mjs';

const cuerpo = `
<section class="acc-verde">
  <p class="rotulo">resumen</p>
  <h1>Tus cifras coinciden <em>con las del SAT</em></h1>
  <p class="lead">Bily tiene el mismo universo de CFDI que el SAT y aplica los mismos criterios de periodo.</p>
  <div class="cifras">
    <div class="destacar"><b>1,284</b><span>CFDI revisados en el mes</span></div>
    <div><b>17</b><span>Hallazgos que requieren atención</span></div>
    <div><b>$2.4 M</b><span>Ingresos netos del periodo</span></div>
  </div>
  <p>Cada CFDI se consulta <strong>todos los días</strong> ante el SAT, y cada hora cuando hay una cancelación en proceso. El detalle vive en el visor, por ejemplo el folio <span class="dato">A-10482</span> del RFC <span class="dato">EEJ010101AB1</span>.</p>
  <p class="nota"><strong>No es una foto del día que descargaste.</strong> Bily mantiene el estatus al día y guarda el historial de cada cambio.</p>
</section>

<section class="acc-rojo">
  <p class="rotulo">hallazgos</p>
  <h1>Lo que requiere <em>tu atención</em></h1>
  <table>
    <thead><tr><th>Hallazgo</th><th>Estado</th><th class="num">CFDI</th><th class="num">Importe</th></tr></thead>
    <tbody>
      <tr><td>Proveedor en lista 69-B</td><td><span class="pill bad">Con riesgo</span></td><td class="num">2</td><td class="num">$84,120.00</td></tr>
      <tr><td>Posible duplicado</td><td><span class="pill dup">Por revisar</span></td><td class="num">5</td><td class="num">$31,400.50</td></tr>
      <tr><td>Pago pendiente (PPD sin REP)</td><td><span class="pill wa">Pendiente</span></td><td class="num">9</td><td class="num">$212,870.00</td></tr>
      <tr><td>Cancelación aceptada</td><td><span class="pill ok">Resuelto</span></td><td class="num">1</td><td class="num">$4,640.00</td></tr>
    </tbody>
    <tfoot><tr><td>Total</td><td></td><td class="num">17</td><td class="num">$333,030.50</td></tr></tfoot>
  </table>
  <p class="aviso">Los importes son ilustrativos. El reporte real toma las cifras de tu cuenta al cierre del mes.</p>
</section>

<section class="acc-morado">
  <p class="rotulo">siguientes pasos</p>
  <h1>Qué hacer <em>con cada hallazgo</em></h1>
  <ol>
    <li><strong>Revisa a los dos proveedores en lista 69-B</strong> antes de deducir sus comprobantes.</li>
    <li><strong>Confirma los cinco posibles duplicados</strong> y descarta los que sean equivalencias válidas.</li>
    <li><strong>Solicita los complementos de pago</strong> de las nueve facturas PPD pendientes.</li>
  </ol>
  <h2>Cómo se calcula</h2>
  <p>Los conteos consideran los CFDI con fecha de emisión dentro del periodo, con los mismos filtros que aplica el SAT.</p>
  <ul>
    <li>Emitidos y recibidos, vigentes y cancelados.</li>
    <li>Notas de crédito y anticipos aplicados a su factura.</li>
  </ul>
</section>`;

const ruta = await crearPdf({
  tipo: 'Reporte mensual',
  titulo: 'Control fiscal de octubre,',
  tenue: 'sin pendientes ocultos',
  lead: 'Lo que Bily revisó en tus CFDI emitidos y recibidos, lo que encontró y lo que requiere tu atención.',
  meta: [['Periodo', 'Octubre 2026'], ['Preparado para', 'Empresa Ejemplo, S.A. de C.V.'], ['Datos', 'Ilustrativos']],
  cuerpo,
}, fileURLToPath(new URL('../salida/ejemplo-pdf.pdf', import.meta.url)));
console.log('Listo:', ruta);
