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
    if (!row['Cliente:.1']) continue; // Saltear filas vacías

    // 🛑 Primero abrir la pestaña
    const nuevaPestana = window.open('', '_blank');
    if (!nuevaPestana) {
      alert('Por favor, habilita las ventanas emergentes en tu navegador.');
      return;
    }

    // 🛠 Después construir el contenido
    const contenido = `
      <html>
      <head>
        <title>Remito ${row['Número Interno:'] || '(sin número)'}</title>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
          h1 { text-align: center; }
          .remito { border: 2px solid #333; padding: 20px; border-radius: 8px; background-color: #f9f9f9; }
          button { margin-top: 20px; padding: 10px 20px; font-size: 16px; cursor: pointer; }
        </style>
      </head>
      <body>
        <div class="remito" id="remito">
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

        <button onclick="descargarPDF()">Descargar como PDF</button>

        <script>
          function descargarPDF() {
            const remito = document.getElementById('remito');
            html2pdf().from(remito).save('remito_${row['Número Interno:'] || 'sin_numero'}.pdf');
          }
        </script>
      </body>
      </html>
    `;

    // Cargar el contenido en la nueva pestaña
    nuevaPestana.document.open();
    nuevaPestana.document.write(contenido);
    nuevaPestana.document.close();
  }
});

