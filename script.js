document.getElementById('generateBtn').addEventListener('click', async function () {
  const input = document.getElementById('fileInput');
  if (!input.files.length) {
    alert('Por favor, sube un archivo Excel.');
    return;
  }

  const { jsPDF } = window.jspdf;
  const file = input.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  if (json.length === 0) {
    alert('No se encontraron datos en el archivo.');
    return;
  }

  const convertirFecha = (valor) => {
    if (!valor) return '';
    const num = Number(valor);
    if (!isNaN(num)) {
      const fecha = new Date(Date.UTC(1899, 11, 30) + num * 86400000);
      return `${String(fecha.getDate()).padStart(2, '0')}/${String(fecha.getMonth() + 1).padStart(2, '0')}/${fecha.getFullYear()}`;
    }
    if (typeof valor === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(valor)) {
      const [y, m, d] = valor.split('-');
      return `${d}/${m}/${y}`;
    }
    return valor;
  };

  const doc = new jsPDF();
  const remitoBase = 24291;
  const prefijo = '0283-';

  json.forEach((originalRow, index) => {
    if (!originalRow || Object.keys(originalRow).length === 0) return;
    if (index > 0) doc.addPage();

    const row = {};
    for (const [key, value] of Object.entries(originalRow)) {
      const cleanKey = key.replace(/[:\s]+$/g, '').replace(/\s{2,}/g, ' ').replace(/[:]/g, '').trim();
      row[cleanKey] = value != null ? value.toString() : '';
    }

    const usados = new Set();
    let y = 20;

    const remitoNro = prefijo + String(remitoBase + index).padStart(8, '0');

    // === ENCABEZADO ===
    doc.setFillColor(77, 77, 77);
    doc.rect(10, y, 190, 20, 'F');
    doc.setTextColor(255);
    doc.setFontSize(16);
    doc.text(`Remito N° ${remitoNro}`, 105, y + 8, { align: 'center' });

    doc.setFontSize(12);
    doc.text(`Número Interno: ${row['Número Interno'] || ''}`, 105, y + 14, { align: 'center' });

    const fechaEmision = convertirFecha(row['Fecha de Emisión']);
    doc.text(`Fecha de Emisión: ${fechaEmision}`, 105, y + 20, { align: 'center' });

    usados.add('Número Interno');
    usados.add('Fecha de Emisión');
    y += 25;
    doc.setTextColor(0);

    // === SYNGENTA ===
    doc.setFillColor(191, 191, 191);
    doc.rect(10, y, 190, 30, 'F');
    const valoresFijos = {
      'C.U.I.T.': '30-64632845-0',
      'Ingresos Brutos (CM)': '901-962580-1',
      'Inicio de actividades': '31/12/1991',
      'I.V.A.': 'Responsable Inscripto'
    };
    const camposSyngenta = ['Nro. Transporte', 'Transporte', 'C.U.I.T.', 'Ingresos Brutos (CM)', 'Inicio de actividades', 'I.V.A.', 'Fecha de Vencimiento del C.A.I.', 'C.A.I. Nº'];
    camposSyngenta.forEach(campo => {
      let valor = valoresFijos[campo] || row[campo] || '';
      if (campo.toLowerCase().includes('fecha') || campo.toLowerCase().includes('inicio')) {
        valor = convertirFecha(valor);
      }
      doc.setFontSize(10);
      doc.text(`${campo}: ${valor}`, 20, y + 6);
      usados.add(campo);
      y += 6;
    });

    // === EMISOR ===
    doc.setFillColor(217, 217, 217);
    doc.rect(10, y, 190, 20, 'F');
    const camposEmisor = ['Cliente', 'Deposito Origen', 'Dirección receptor', 'Teléfono Recptor', 'Pedido'];
    camposEmisor.forEach(campo => {
      if (row[campo]) {
        doc.text(`${campo}: ${row[campo]}`, 20, y + 6);
        usados.add(campo);
        y += 6;
      }
    });

    // === RECEPTOR ===
    doc.setFillColor(230, 230, 230);
    doc.rect(10, y, 190, 20, 'F');
    const camposReceptor = ['Cliente Receptor', 'Deposito Destino', 'Dirección receptor', 'Código de Cliente', 'C.U.I.T. Receptor'];
    camposReceptor.forEach(campo => {
      if (row[campo]) {
        doc.text(`${campo}: ${row[campo]}`, 20, y + 6);
        usados.add(campo);
        y += 6;
      }
    });

    // === PRODUCTOS ===
    doc.setFontSize(12);
    doc.setTextColor(0);
    doc.text('Productos:', 20, y + 6);
    y += 8;
    doc.setFontSize(10);
    const camposProducto = ['Código', 'Descripción', 'Cantidad', 'Lotes', 'Peso estimado Total'];
    camposProducto.forEach(campo => {
      if (row[campo]) {
        doc.text(`${campo}: ${row[campo]}`, 20, y);
        usados.add(campo);
        y += 6;
      }
    });

    // === OTROS CAMPOS ===
    doc.setFontSize(10);
    doc.text('Otros campos:', 20, y); y += 6;
    for (const key in row) {
      if (usados.has(key)) continue;
      let valor = row[key];
      if (key.toLowerCase().includes('fecha') || key.toLowerCase().includes('inicio')) {
        valor = convertirFecha(valor);
      }
      doc.text(`${key}: ${valor}`, 20, y);
      y += 6;
      if (y > 270) {
        doc.addPage();
        y = 20;
      }
    }

    y += 6;
    doc.setFontSize(12);
    doc.text('Recibí Conforme: ___________________________', 20, y); y += 10;
    doc.setFontSize(8);
    doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 20, 280);
    doc.text('Seguro de mercadería por cuenta de Syngenta.', 20, 285);
  });

  doc.save('Remitos_Syngenta.pdf');
});

