// Ejemplo de presentación. Datos 100% ficticios, rotulados como ilustrativos.
import { fileURLToPath } from 'node:url';
import { crearPresentacion } from '../lib/pptx.mjs';

const p = crearPresentacion({ titulo: 'Control fiscal continuo' });

p.portada({
  rotulo: 'Presentación comercial',
  titulo: 'Control fiscal continuo',
  tenue: 'para tus CFDI',
  pie: 'Empresa Ejemplo, S.A. de C.V.  ·  Noviembre 2026  ·  Datos ilustrativos',
});

p.contenido({
  rotulo: 'el problema',
  acento: 'rojo',
  titulo: 'Lo que no revisas',
  tenue: 'se convierte en riesgo',
  lead: 'Un CFDI puede cambiar después de recibirlo. Nadie lo vigila todos los días.',
  puntos: [
    '**Cancelaciones sin aviso.** Un comprobante vigente pasa a cancelado.',
    '**Proveedores en lista 69-B.** El riesgo aparece meses después de la operación.',
    '**Duplicados.** El mismo gasto entra dos veces a la contabilidad.',
  ],
});

p.seccion({ rotulo: 'cómo lo resuelve bily', titulo: 'Una capa de control', tenue: 'que no descansa' });

p.cifras({
  rotulo: 'resultado',
  acento: 'verde',
  titulo: 'Tus cifras coinciden',
  tenue: 'con las del SAT',
  lead: 'El mismo universo de CFDI y los mismos criterios de periodo.',
  cifras: [
    { valor: '1,284', etiqueta: 'CFDI revisados en el mes', destacar: true },
    { valor: '17', etiqueta: 'Hallazgos que requieren atención' },
    { valor: '24 h', etiqueta: 'Entre una consulta y la siguiente al SAT' },
  ],
  nota: '**Cifras ilustrativas.** El reporte real toma los datos de tu cuenta.',
});

p.tabla({
  rotulo: 'hallazgos',
  acento: 'morado',
  titulo: 'Bily te dice',
  tenue: 'qué requiere atención',
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
});

p.dosColumnas({
  rotulo: 'conexión',
  acento: 'naranja',
  titulo: 'Tus CFDI llegan a Bily',
  tenue: 'de tres formas',
  puntos: ['**Sincronización con el SAT.** La recomendada: cubre todo el universo.', '**Timbrado y autofacturación.**', '**Carga directa de XML.**'],
  derecha: ['**Se combinan.** Puedes timbrar, cargar y sincronizar a la vez.', '**Sin migración.** No depende de tu sistema de facturación.'],
});

p.cierre({
  titulo: 'Bily revisa cada CFDI',
  tenue: 'para que nada se te escape',
  accion: 'Solicitar demostración',
  contacto: ['bily.mx'],
});

const ruta = await p.guardar(fileURLToPath(new URL('../salida/ejemplo-presentacion.pptx', import.meta.url)));
console.log('Listo:', ruta);
