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
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    if (json.length === 0) {
        alert('No se encontraron datos en la hoja.');
        return;
    }

    const doc = new jsPDF();

    for (let index = 0; index < json.length; index++) {
        const row = json[index];
        if (!row || Object.keys(row).length === 0) continue;

        if (index !== 0) {
            doc.addPage();
        }

        // Evitar duplicar campos
        const usados = new Set();

        // Encabezado
        doc.setFontSize(16);
        doc.text(`Remito N° ${row['Número Interno:'] || '(sin número)'}`, 105, 15, { align: 'center' });

        doc.setFontSize(10);
        let y = 25;

        // --- Syngenta (orden personalizado)
        const camposFijos = [
            'Nro. Transporte:',
            'C.U.I.T.:',
            'Ingresos Brutos (CM):',
            'Inicio de actividades:',
            'I.V.A.:',
            'Fecha de Vencimiento del C.A.I.:',
            'C.A.I. Nº:',
            'Fecha de Emisión:'
        ];

        camposFijos.forEach(campo => {
            doc.text(`${campo} ${row[campo] || ''}`, 20, y);
            usados.add(campo);
            y += 6;
        });

        // --- Emisor
        const camposEmisor = [
            'Cliente Recptor:',
            'Deposito Origen',
            'Dirección receptor:',
            'Teléfono Recptor:',
            'Pedido:',
            'Transporte:',
            'Nro. Transporte:'
        ];
        camposEmisor.forEach(campo => {
            doc.text(`${campo} ${row[campo] || ''}`, 20, y);
            usados.add(campo);
            y += 6;
        });

        // --- Receptor
        const camposReceptor = [
            'Deposito Destino',
            'Código de Cliente:',
            'Cliente Receptor:',
            'Dirección receptor:',
            'C.U.I.T. Receptor:',
            'Pedido:'
        ];
        camposReceptor.forEach(campo => {
            doc.text(`${campo} ${row[campo] || ''}`, 20, y);
            usados.add(campo);
            y += 6;
        });

        // --- Productos
        doc.setFontSize(12);
        doc.text('Productos:', 20, y); y += 8;
        doc.setFontSize(10);
        const camposProducto = [
            'Código:',
            'Descripción:',
            'Cantidad:',
            'Lotes:',
            'PESO ESTIMADO TOTAL:'
        ];
        camposProducto.forEach(campo => {
            doc.text(`${campo} ${row[campo] || ''}`, 20, y);
            usados.add(campo);
            y += 6;
        });

        // --- Otros campos (dinámicos)
        doc.setFontSize(10);
        doc.text('Otros campos:', 20, y); y += 6;

        for (const key in row) {
            if (usados.has(key)) continue;
            const value = String(row[key]).trim();
            doc.text(`${key}: ${value}`, 20, y);
            y += 6;

            if (y > 270) {
                doc.addPage();
                y = 20;
            }
        }

        // Firma
        y += 6;
        doc.setFontSize(12);
        doc.text('Recibí Conforme: ___________________________', 20, y); y += 10;

        // Pie de página
        doc.setFontSize(8);
        doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 20, 280);
        doc.text('Seguro de mercadería por cuenta de Syngenta.', 20, 285);
    }

    doc.save('Remitos_Syngenta.pdf');
});

