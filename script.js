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

    // Crear PDF usando jsPDF
    const pdf = new jspdf.jsPDF();

    pdf.setFontSize(16);
    pdf.text(`Remito N° ${dataRow['Número Interno:'] || '(sin número)'}`, 20, 20);

    pdf.setFontSize(12);
    pdf.text(`Fecha de Emisión: ${dataRow['Fecha de Emisión:'] || ''}`, 20, 40);
    pdf.text(`Cliente: ${dataRow['Cliente Recptor:'] || ''}`, 20, 50);
    pdf.text(`Dirección: ${dataRow['Dirección receptor:'] || ''}`, 20, 60);
    pdf.text(`CUIT: ${dataRow['C.U.I.T. RECPTOR:'] || ''}`, 20, 70);
    pdf.text(`Pedido: ${dataRow['Pedido:'] || ''}`, 20, 80);

    pdf.setFontSize(14);
    pdf.text('Productos:', 20, 100);

    pdf.setFontSize(12);
    pdf.text(`Código: ${dataRow['Código:'] || ''}`, 20, 110);
    pdf.text(`Descripción: ${dataRow['Descripción:'] || ''}`, 20, 120);
    pdf.text(`Cantidad: ${dataRow['Cantidad:'] || ''}`, 20, 130);
    pdf.text(`Peso Estimado Total: ${dataRow['PESO ESTIMADO TOTAL:'] || ''}`, 20, 140);
    pdf.text(`Lotes: ${dataRow['Lotes:'] || ''}`, 20, 150);

    pdf.setFontSize(14);
    pdf.text('Transporte:', 20, 170);

    pdf.setFontSize(12);
    pdf.text(`Nro. Transporte: ${dataRow['Nro. Transporte:'] || ''}`, 20, 180);
    pdf.text(`Nombre Transporte: ${dataRow['Transporte:'] || ''}`, 20, 190);

    // Guardar el PDF
    pdf.save(`remito_${dataRow['Número Interno:'] || index + 1}.pdf`);
  }
});

