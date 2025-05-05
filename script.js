
document.getElementById('generateBtn').addEventListener('click', async function () {
  const input = document.getElementById('fileInput');
  if (!input.files.length) return alert('Subí un archivo Excel primero.');

  const { jsPDF } = window.jspdf;
  const file = input.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  const convertirFecha = (val) => {
    if (!val) return '';
    const num = Number(val);
    if (!isNaN(num)) {
      const fecha = new Date(Date.UTC(1899, 11, 30) + num * 86400000);
      return `${String(fecha.getDate()).padStart(2, '0')}/${String(fecha.getMonth() + 1).padStart(2, '0')}/${fecha.getFullYear()}`;
    }
    if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(val)) {
      const [y, m, d] = val.split('-');
      return `${d}/${m}/${y}`;
    }
    return val;
  };

  const doc = new jsPDF();
  const remitoBase = 24291;
  const internoBase = 1910075353;
  const prefijo = '0283-';

  json.forEach((originalRow, index) => {
    if (!originalRow || Object.keys(originalRow).length === 0) return;
    if (index > 0) doc.addPage();

    const row = {};
    for (const [key, val] of Object.entries(originalRow)) {
      const cleanKey = key.replace(/[:\s]+$/g, '').replace(/\s{2,}/g, ' ').replace(/[:]/g, '').trim();
      row[cleanKey] = val != null ? val.toString() : '';
    }

    let y = 15;
    const usados = new Set();
    const remitoNro = prefijo + String(remitoBase + index).padStart(8, '0');
    const internoNro = internoBase + index;

    // === ENCABEZADO GRIS ===
    doc.setFillColor(224, 224, 224);
    doc.rect(10, y, 190, 20, 'F');
    doc.setFontSize(16);
    doc.text(`Remito N° ${remitoNro}`, 105, y + 7, { align: 'center' });
    doc.setFontSize(12);
    doc.text(`Número Interno: ${internoNro}`, 105, y + 14, { align: 'center' });
    y += 25;

    // === FECHA DE EMISIÓN ===
    doc.setFontSize(10);
    if (row['Fecha de Emisión']) {
      doc.text(`Fecha de Emisión: ${convertirFecha(row['Fecha de Emisión'])}`, 15, y);
      usados.add('Fecha de Emisión');
      y += 6;
    }

    // === DATOS FIJOS SYNGENTA ===
    const fijos = {
      'C.U.I.T.': '30-64632845-0',
      'Ingresos Brutos (CM)': '901-962580-1',
      'Inicio de actividades': '31/12/1991',
      'I.V.A.': 'Responsable Inscripto'
    };

    const camposSyngenta = [
      'Transporte', 'Nro. Transporte',
      'C.U.I.T.', 'Ingresos Brutos (CM)', 'Inicio de actividades',
      'I.V.A.', 'Fecha de Vencimiento del C.A.I.', 'C.A.I. Nº'
    ];

    camposSyngenta.forEach(campo => {
      let valor = fijos[campo] || row[campo] || '';
      if (campo.toLowerCase().includes('fecha') || campo.toLowerCase().includes('inicio')) {
        valor = convertirFecha(valor);
      }
      doc.text(`${campo}: ${valor}`, 15, y);
      usados.add(campo);
      y += 6;
    });

    // === CUADRO GRIS CLIENTE ===
    doc.setFillColor(245, 245, 245);
    doc.rect(10, y, 190, 25, 'F');
    doc.text(`Cliente: ${row['Cliente Receptor'] || row['Cliente'] || ''}`, 15, y + 6);
    doc.text(`Dirección: ${row['Dirección receptor'] || ''}`, 15, y + 12);
    doc.text(`CUIT: ${row['C.U.I.T. Receptor'] || ''}`, 15, y + 18);
    usados.add('Cliente Receptor'); usados.add('Cliente'); usados.add('Dirección receptor'); usados.add('C.U.I.T. Receptor');
    y += 30;

    // === CAMPOS GENERALES ===
    const generales = [
      'Deposito Origen', 'Deposito Destino', 'Teléfono Recptor',
      'Código de Cliente', 'Pedido'
    ];
    generales.forEach(campo => {
      if (row[campo]) {
        doc.text(`${campo}: ${row[campo]}`, 15, y);
        usados.add(campo);
        y += 6;
      }
    });

    // === PRODUCTOS ===
    doc.setFontSize(11);
    doc.text('Productos:', 15, y); y += 6;
    doc.setFontSize(10);
    const productos = ['Código', 'Descripción', 'Cantidad', 'Lotes', 'Peso estimado Total'];
    productos.forEach(p => {
      if (row[p]) {
        doc.text(`${p}: ${row[p]}`, 15, y);
        usados.add(p);
        y += 6;
      }
    });

    // === OTROS CAMPOS ===
    doc.setFontSize(10);
    doc.text('Otros campos:', 15, y); y += 6;
    for (const key in row) {
      if (usados.has(key)) continue;
      let valor = row[key];
      if (key.toLowerCase().includes('fecha') || key.toLowerCase().includes('inicio')) {
        valor = convertirFecha(valor);
      }
      doc.text(`${key}: ${valor}`, 15, y);
      y += 6;
      if (y > 270) {
        doc.addPage();
        y = 20;
      }
    }

    // === FIRMA Y PIE ===
    y += 6;
    doc.setFontSize(12);
    doc.text('Recibí Conforme: ___________________________', 15, y); y += 10;

    doc.setFontSize(8);
    doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 15, 280);
    doc.text('Seguro de mercadería por cuenta de Syngenta. Jurisdicción Rosario - Santa Fe.', 15, 285);
  });

  doc.save('Remitos_Syngenta.pdf');
});
