json.forEach((originalRow, index) => {
  if (!originalRow || Object.keys(originalRow).length === 0) return;
  if (index !== 0) doc.addPage();

  // 🔧 Normalizar claves y forzar valores a string (incluso si son números)
  const row = {};
  for (const [key, value] of Object.entries(originalRow)) {
    const cleanKey = key
      .replace(/[:\s]+$/g, '')       // elimina : y espacios finales
      .replace(/[:\s]+/g, ' ')       // reemplaza múltiples espacios o :
      .trim();

    const stringValue = typeof value === 'number' ? String(value) : (value || '').toString();
    row[cleanKey] = stringValue;
  }

  const usados = new Set();

  // === ENCABEZADO ===
  doc.setFontSize(16);
  doc.text(`Remito N° ${row['Remito N°'] || '(sin número)'}`, 105, 15, { align: 'center' });

  doc.setFontSize(12);
  doc.text(`Número Interno: ${row['Número Interno'] || ''}`, 105, 22, { align: 'center' });

  const fechaEmision = convertirFecha(row['Fecha de Emisión']);
  doc.text(`Fecha de Emisión: ${fechaEmision}`, 105, 29, { align: 'center' });

  usados.add('Remito N°');
  usados.add('Número Interno');
  usados.add('Fecha de Emisión');

  doc.setFontSize(10);
  let y = 40;

  // === SYNGENTA ===
  const camposSyngenta = [
    'C.U.I.T.',
    'Ingresos Brutos (CM)',
    'Inicio de actividades',
    'I.V.A.',
    'Fecha de Vencimiento del C.A.I.',
    'C.A.I. Nº'
  ];
  camposSyngenta.forEach(campo => {
    let valor = row[campo] || '';
    if (campo.toLowerCase().includes('fecha')) valor = convertirFecha(valor);
    doc.text(`${campo}: ${valor}`, 20, y);
    usados.add(campo);
    y += 6;
  });

  // === EMISOR ===
  const camposEmisor = [
    'Cliente Recptor',
    'Deposito Origen',
    'Dirección receptor',
    'Teléfono Recptor',
    'Pedido',
    'Transporte',
    'Nro. Transporte'
  ];
  camposEmisor.forEach(campo => {
    doc.text(`${campo}: ${row[campo] || ''}`, 20, y);
    usados.add(campo);
    y += 6;
  });

  // === RECEPTOR ===
  const camposReceptor = [
    'Deposito Destino',
    'Código de Cliente',
    'Cliente Receptor',
    'Dirección receptor',
    'C.U.I.T. Receptor',
    'Pedido'
  ];
  camposReceptor.forEach(campo => {
    doc.text(`${campo}: ${row[campo] || ''}`, 20, y);
    usados.add(campo);
    y += 6;
  });

  // === PRODUCTOS ===
  doc.setFontSize(12);
  doc.text('Productos:', 20, y); y += 8;
  doc.setFontSize(10);
  const camposProducto = [
    'Código',
    'Descripción',
    'Cantidad',
    'Lotes',
    'PESO ESTIMADO TOTAL'
  ];
  camposProducto.forEach(campo => {
    doc.text(`${campo}: ${row[campo] || ''}`, 20, y);
    usados.add(campo);
    y += 6;
  });

  // === OTROS CAMPOS ===
  doc.setFontSize(10);
  doc.text('Otros campos:', 20, y); y += 6;

  for (const key in row) {
    if (usados.has(key)) continue;

    let valor = row[key];
    if (key.toLowerCase().includes('fecha')) valor = convertirFecha(valor);

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

