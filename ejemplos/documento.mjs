// Ejemplo de documento de Word. Datos 100% ficticios, rotulados como ilustrativos.
import { crearDocumento } from '../lib/docx.mjs';

const doc = crearDocumento({
  tipo: 'Reporte mensual',
  titulo: 'Control fiscal de octubre,',
  tenue: 'sin pendientes ocultos',
  lead: 'Lo que Bily revisó en tus CFDI emitidos y recibidos, lo que encontró y lo que requiere tu atención.',
  meta: [['Periodo', 'Octubre 2026'], ['Preparado para', 'Empresa Ejemplo, S.A. de C.V.'], ['Datos', 'Ilustrativos']],
});

doc
  .rotulo('resumen')
  .h1('Tus cifras coinciden', 'con las del SAT')
  .lead('Bily tiene el mismo universo de CFDI que el SAT y aplica los mismos criterios de periodo.')
  .cifras([
    { valor: '1,284', etiqueta: 'CFDI revisados en el mes', destacar: true },
    { valor: '17', etiqueta: 'Hallazgos que requieren atención' },
    { valor: '$2,418,930.00', etiqueta: 'Ingresos netos del periodo' },
  ])
  .p('Cada CFDI se consulta **todos los días** ante el SAT, y cada hora cuando hay una cancelación en proceso. El detalle vive en el visor, ya filtrado, por ejemplo el folio `A-10482` del RFC `EEJ010101AB1`.')
  .nota('**No es una foto del día que descargaste.** Bily mantiene el estatus al día y guarda el historial de cada cambio.')

  .acento('rojo')
  .rotulo('hallazgos')
  .h1('Lo que requiere', 'tu atención')
  .tabla({
    columnas: [
      { titulo: 'Hallazgo', ancho: 3 },
      { titulo: 'Estado', ancho: 2, tipo: 'estado' },
      { titulo: 'CFDI', ancho: 1, tipo: 'num' },
      { titulo: 'Importe', ancho: 2, tipo: 'num' },
    ],
    filas: [
      ['Proveedor en lista 69-B', { estado: 'bad', texto: 'Con riesgo' }, '2', '$84,120.00'],
      ['Posible duplicado', { estado: 'dup', texto: 'Por revisar' }, '5', '$31,400.50'],
      ['Pago pendiente (PPD sin REP)', { estado: 'wa', texto: 'Pendiente' }, '9', '$212,870.00'],
      ['Cancelación aceptada', { estado: 'ok', texto: 'Resuelto' }, '1', '$4,640.00'],
    ],
    total: ['Total', '', '17', '$333,030.50'],
  })
  .aviso('Los importes son ilustrativos. El reporte real toma las cifras de tu cuenta al cierre del mes.')

  .acento('morado')
  .rotulo('siguientes pasos')
  .h1('Qué hacer', 'con cada hallazgo')
  .numerada([
    '**Revisa a los dos proveedores en lista 69-B** antes de deducir sus comprobantes.',
    '**Confirma los cinco posibles duplicados** y descarta los que sean equivalencias válidas.',
    '**Solicita los complementos de pago** de las nueve facturas PPD pendientes.',
  ])
  .h2('Cómo se calcula')
  .p('Los conteos consideran los CFDI con fecha de emisión dentro del periodo, con los mismos filtros que aplica el SAT.')
  .lista(['Emitidos y recibidos, vigentes y cancelados.', 'Notas de crédito y anticipos aplicados a su factura.', 'Cancelaciones posteriores que modifican un periodo anterior.']);

const ruta = await doc.guardar(new URL('../salida/ejemplo-documento.docx', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'));
console.log('Listo:', ruta);
