document.getElementById('generateBtn').addEventListener('click', async function() {
  const input = document.getElementById('fileInput');
  if (!input.files.length) {
    alert('Por favor, sube un archivo Excel.');
    return;
  }

  const file = input.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);

  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  if (json.length === 0) {
    alert('No se encontraron datos en la hoja de Excel.');
    return;
  }

  for (let index = 0; index < json.length; index++) {
    const row = json[index];
    if (!row || Object.keys(row).length === 0) continue;

    const dataRow = {};
    for (const key in row) {
      const normalizedKey = key.trim();
      dataRow[normalizedKey] = row[key];
    }

    if (!dataRow['Cliente Recptor:']) continue;

    // Crear un nuevo PDF
    const doc = new jspdf.jsPDF();

    doc.setFontSize(16);
    doc.text(`Remito N° ${dataRow['Número Interno:'] || '(sin número)'}`, 20, 20);

    doc.setFontSize(12);
    let y = 30;
    doc.text(`Fecha de Emisión: ${dataRow['Fecha de Emisión:'] || ''}`, 20, y); y += 10;
    doc.text(`Cliente: ${dataRow['Cliente Recptor:'] || ''}`, 20, y); y += 10;
    doc.text(`Dirección: ${dataRow['Dirección receptor:'] || ''}`, 20, y); y += 10;
    doc.text(`CUIT: ${dataRow['C.U.I.T. RECPTOR:'] || ''}`, 20, y); y += 10;
    doc.text(`Pedido: ${dataRow['Pedido:'] || ''}`, 20, y); y += 10;

    doc.setFontSize(14);
    doc.text('Productos:', 20, y); y += 10;
    doc.setFontSize(12);
    doc.text(`Código: ${dataRow['Código:'] || ''}`, 20, y); y += 10;
    doc.text(`Descripción: ${dataRow['Descripción:'] || ''}`, 20, y); y += 10;
    doc.text(`Cantidad: ${dataRow['Cantidad:'] || ''}`, 20, y); y += 10;
    doc.text(`Peso Estimado Total: ${dataRow['PESO ESTIMADO TOTAL:'] || ''}`, 20, y); y += 10;
    doc.text(`Lotes: ${dataRow['Lotes:'] || ''}`, 20, y); y += 10;

    doc.setFontSize(14);
    doc.text('Transporte:', 20, y); y += 10;
    doc.setFontSize(12);
    doc.text(`Número: ${dataRow['Nro. Transporte:'] || ''}`, 20, y); y += 10;
    doc.text(`Nombre: ${dataRow['Transporte:'] || ''}`, 20, y); y += 10;

    // Guardar el PDF
    doc.save(`remito_${dataRow['Número Interno:'] || index + 1}.pdf`);
  }
});

