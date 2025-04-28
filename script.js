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
    if (!row['Cliente Recptor:']) continue; // Saltar filas vacías

    // Crear un DIV visible para capturar
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
        <h1>Remito N° ${row['Número Interno: '] || '(sin número)'}</h1>
        <p><strong>Fecha de Emisión:</strong> ${row['Fecha de Emisión:']}</p>
        <p><strong>Cliente:</strong> ${row['Cliente Recptor:']}</p>
        <p><strong>Dirección:</strong> ${row['Dirección receptor: ']}</p>
        <p><strong>CUIT:</strong> ${row['C.U.I.T. RECPTOR:']}</p>
        <p><strong>Pedido:</strong> ${row['Pedido:']}</p>
        <h3>Productos</h3>
        <p><strong>Código:</strong> ${row['Código: '] || ''} - ${row['Descripción:'] || ''}</p>
        <p><strong>Cantidad:</strong> ${row['Cantidad:'] || ''}</p>
        <p><strong>Peso Estimado Total:</strong> ${row['PESO ESTIMADO TOTAL: '] || ''}</p>
        <p><strong>Lotes:</strong> ${row['Lotes:'] || ''}</p>
        <h3>Transporte</h3>
        <p><strong>Número:</strong> ${row['Nro. Transporte:'] || ''} - <strong>Nombre:</strong> ${row['Transporte:'] || ''}</p>
      </div>
    `;
    document.body.appendChild(div);

    // 👇 Esperamos a que el navegador pinte el contenido
    await new Promise(resolve => setTimeout(resolve, 300)); 

    // Ahora sí capturamos y generamos el PDF
    await html2pdf().from(div).set({
      filename: `remito_${row['Número Interno: '] || index + 1}.pdf`,
      margin: 10,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2 },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    }).save();

    // Limpiar el div
    document.body.removeChild(div);
  }
});
