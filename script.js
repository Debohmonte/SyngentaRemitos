document.getElementById('generateBtn').addEventListener('click', async function() {
  const input = document.getElementById('fileInput');
  if (!input.files.length) {
    alert('Por favor, sube un archivo Excel.');
    return;
  }

  const file = input.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);

  const sheetName = workbook.SheetNames[1]; // Hoja 2
  const worksheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  for (let index = 0; index < json.length; index++) {
    const row = json[index];
    if (!row['Cliente:.1']) continue; // Saltar filas vacías

    // Crear un DIV temporal invisible para generar el PDF
    const div = document.createElement('div');
    div.style.display = 'none';
    div.innerHTML = `
      <div class="remito">
        <h1>Remito N° ${row['Número Interno:'] || '(sin número)'}</h1>
        <p><strong>Fecha de Emisión:</strong> ${row['Fecha de Emisión:']}</p>
        <p><strong>Cliente:</strong> ${row['Cliente:.1']}</p>
        <p><strong>Dirección:</strong> ${row['Dirección:']}</p>
        <p><strong>CUIT:</strong> ${row['C.U.I.T.:.1']}</p>
        <p><strong>Pedido:</strong> ${row['Pedido:']}</p>
        <h3>Productos</h3>
        <p><strong>Código:</strong> ${row['Código:']} - ${row['Descripción:']}</p>
        <p><strong>Cantidad:</strong> ${row['Cantidad:']}</p>
        <p><strong>Peso Estimado Total:</strong> ${row['PESO ESTIMADO TOTAL: ']}</p>
        <p><strong>Lotes:</strong> ${row['Lotes:']}</p>
        <h3>Transporte</h3>
        <p><strong>Número:</strong> ${row['Nro. Transporte:']} - <strong>Nombre:</strong> ${row['Transporte:']}</p>
      </div>
    `;
    document.body.appendChild(div);

    // Usar html2pdf para generar y descargar automáticamente
    await html2pdf().from(div).set({
      filename: `remito_${row['Número Interno:'] || index + 1}.pdf`,
      margin: 10,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2 },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    }).save();

    document.body.removeChild(div);
  }
});


