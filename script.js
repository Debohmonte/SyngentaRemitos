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

    // === ENCABEZADO ===
    doc.setFillColor(100); // gris oscuro
    doc.rect(10, y, 190, 20, 'F');
    doc.setFontSize(16);
    doc.setTextColor(255);
    const remitoNro = prefijo + String(remitoBase + index).padStart(8, '0');
    doc.text(`Remito N° ${remitoNro}`, 105, y + 8, { align: 'center' });

    doc.setFontSize(12);
    doc.text(`Número Interno: ${row['Número Interno'] || ''}`, 105, y + 14, { align: 'center' });

    const fechaEmision = convertirFecha(row['Fecha de Emisión']);
    doc.text(`Fecha de Emisión: ${fechaEmision}`, 105, y + 20, { align: 'center' });
    y += 28;
    doc.setTextColor(0);

    // === SYNGENTA ===
    doc.setFillColor(200);
    const syngentaBoxStart = y;
    const camposSyngenta = [
      ['Nro. Transporte', row['Nro. Transporte']],
      ['Transporte', row['Transporte']],
      ['C.U.I.T.', '30-64632845-0'],
      ['Ingresos Brutos (CM)', '901-962580-1'],
      ['Inicio de actividades', '31/12/1991'],
      ['I.V.A.', 'Responsable Inscripto'],
      ['Fecha de Vencimiento del C.A.I.', convertirFecha(row['Fecha de Vencimiento del C.A.I.'])],
      ['C.A.I. Nº', row['C.A.I. Nº']]
    ];
    doc.rect(10, y, 190, camposSyngenta.length * 6, 'F');
    camposSyngenta.forEach(([label, value]) => {
      doc.text(`${label}: ${value || ''}`, 20, y + 5);
      usados.add(label);
      y += 6;
    });

    // === EMISOR ===
    doc.setFillColor(230);
    const camposEmisor = [
      ['Cliente', row['Cliente']],
      ['Deposito Origen', row['Deposito Origen']],
      ['Dirección receptor', row['Dirección receptor']],
      ['Teléfono Recptor', row['Teléfono Recptor']]
    ];
    doc.rect(10, y, 190, camposEmisor.length * 6, 'F');
    camposEmisor.forEach(([label, value]) => {
      doc.text(`${label}: ${value || ''}`, 20, y + 5);
      usados.add(label);
      y += 6;
    });

    // === RECEPTOR ===
    doc.setFillColor(240);
    const camposReceptor = [
      ['Código de Cliente', row['Código de Cliente']],
      ['Cliente Receptor', row['Cliente Receptor']],
      ['Deposito Destino', row['Deposito Destino']],
      ['Dirección receptor', row['Dirección receptor']],
      ['C.U.I.T. Receptor', row['C.U.I.T. Receptor']],
      ['Pedido', row['Pedido']]
    ];
    doc.rect(10, y, 190, camposReceptor.length * 6, 'F');
    camposReceptor.forEach(([label, value]) => {
      doc.text(`${label}: ${value || ''}`, 20, y + 5);
      usados.add(label);
      y += 6;
    });

    // === PRODUCTOS (blanco) ===
    doc.setFontSize(12);
    doc.text('Productos:', 20, y); y += 8;
    doc.setFontSize(10);
    const camposProducto = [
      ['Código', row['Código']],
      ['Descripción', row['Descripción']],
      ['Cantidad', row['Cantidad']],
      ['Lotes', row['Lotes']],
      ['Peso estimado Total', row['Peso estimado Total']]
    ];
    camposProducto.forEach(([label, value]) => {
      doc.text(`${label}: ${value || ''}`, 20, y);
      usados.add(label);
      y += 6;
    });

    // === OTROS ===
    doc.setFontSize(10);
    doc.setFillColor(245);
    doc.rect(10, y, 190, 10, 'F');
    doc.text('Otros campos:', 20, y + 6);
    y += 12;

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

    // === FIRMA Y PIE ===
    y += 6;
    doc.setFontSize(12);
    doc.text('Recibí Conforme: ___________________________', 20, y); y += 10;
    doc.setFontSize(8);
    doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 20, 280);
    doc.text('Seguro de mercadería por cuenta de Syngenta.', 20, 285);
  });

  doc.save('Remitos_Syngenta_Estetica.pdf');
});
