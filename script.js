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

  const remitosDiv = document.getElementById('remitos');
  remitosDiv.innerHTML = '';

  json.forEach((row, index) => {
    if (!row['Cliente:.1']) return; // Saltar filas vacías

    // Crear contenedor de remito
    const remito = document.createElement('div');
    remito.className = 'remito';
    remito.innerHTML = `
      <h2>Remito N° ${row['Número Interno:'] || '(sin número)'}</h2>
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
      <button onclick="descargarPDF(this)">Descargar PDF</button>
      <hr>
    `;
    remitosDiv.appendChild(remito);
  });
});

function descargarPDF(button) {
  const remitoDiv = button.parentElement;
  html2pdf()
    .from(remitoDiv)
    .save(`remito.pdf`);
}

  
