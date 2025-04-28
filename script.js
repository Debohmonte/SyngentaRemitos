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

    if (!row || Object.keys(row).length === 0) continue; // Saltar filas vacías

    // Normalizar claves
    const dataRow = {};
    for (const key in row) {
      const normalizedKey = key.trim(); // quitar espacios al principio y final
      dataRow[normalizedKey] = row[key];
    }

    if (!dataRow['Cliente Recptor:']) continue; // ahora seguro

    // Crear un div visible para capturar
    const div = document.createElement('div');
    div.style.position = 'fixed';
    div.style.top = '0';
    div.style.left = '0';
    div.style.background = 'white';
    div.style.zIndex = '9999';
    div.style.width = '800px';
    div.style.padding = '20px';
    div.innerHTML = `
      <div class="remito" style="font-family: Arial, sans-serif;">
        <h1>Remito N° ${dataRow['Número Interno:'] || '(sin número)'}</h1>
        <p><strong>Fecha de Emisión:</strong> ${dataRow['Fecha de Emisión:']}</p>
        <p><strong>Cliente:</strong> ${dataRow['Cliente Recptor:']}</p>
        <p><strong>Dirección:</strong> ${dataRow['Dirección receptor:']}</p>
        <p><strong>CUIT:</strong> ${dataRow['C.U.I.T. RECPTOR:']}</p>
        <p><strong>Pedido:</strong> ${dataRow['Pedido:']}</p>
        <h3>Productos</h3>
        <p><strong>Código:</strong> ${dataRow['Código:']} - ${dataRow['Descripción:']}</p>
        <p><strong>Cantidad:</strong> ${dataRow['Cantidad:']}</p>
        <p><strong>Peso Estimado Total:</strong> ${dataRow['PESO ESTIMADO TOTAL:']}</p>
        <p><strong>Lotes:</strong> ${dataRow['Lotes:']}</p>
        <h3>Transporte</h3>
        <p><strong>Número:</strong> ${dataRow['Nro. Transporte:']} - <strong>Nombre:</strong> ${dataRow['Transporte:']}</p>
      </div>
    `;
    document.body.appendChild(div);

    await new Promise(resolve => setTimeout(resolve, 300));

    await html2pdf().from(div).set({
      filename: `remito_${dataRow['Número Interno:'] || index + 1}.pdf`,
      margin: 10,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2 },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    }).save();

    document.body.removeChild(div);
  }
});

