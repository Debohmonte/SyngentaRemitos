document.getElementById('generateBtn').addEventListener('click', async function() {
    const input = document.getElementById('fileInput');
    if (!input.files.length) {
        alert('Por favor, sube un archivo Excel.');
        return;
    }

    const { jsPDF } = window.jspdf; // ✅ IMPORTANTE
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

    for (let index = 0; index < json.length; index++) {
        const row = json[index];
        if (!row || Object.keys(row).length === 0) continue;

        const dataRow = {};
        for (const key in row) {
            const normalizedKey = key.trim();
            dataRow[normalizedKey] = row[key];
        }

        if (!dataRow['Cliente Recptor:']) continue;

        const doc = new jsPDF(); // ✅ Ahora sí funciona

        // Títulos
        doc.setFontSize(16);
        doc.text(`Remito N° ${dataRow['Número Interno:'] || '(sin número)'}`, 105, 15, { align: 'center' });

        doc.setFontSize(10);

        let y = 25;

        // Datos generales
        doc.text(`Fecha de Emisión: ${dataRow['Fecha de Emisión:'] || ''}`, 20, y); y += 6;
        doc.text(`Cliente: ${dataRow['Cliente Recptor:'] || ''}`, 20, y); y += 6;
        doc.text(`Dirección: ${dataRow['Dirección receptor:'] || ''}`, 20, y); y += 6;
        doc.text(`CUIT: ${dataRow['C.U.I.T. RECPTOR:'] || ''}`, 20, y); y += 6;
        doc.text(`Pedido: ${dataRow['Pedido:'] || ''}`, 20, y); y += 6;

        // Transporte
        doc.text(`Transporte: ${dataRow['Transporte:'] || ''}`, 20, y); y += 6;
        doc.text(`Número Transporte: ${dataRow['Nro. Transporte:'] || ''}`, 20, y); y += 10;

        // Productos
        doc.setFontSize(12);
        doc.text('Productos:', 20, y); y += 8;
        doc.setFontSize(10);

        doc.text(`Código: ${dataRow['Código:'] || ''}`, 20, y); y += 5;
        doc.text(`Descripción: ${dataRow['Descripción:'] || ''}`, 20, y); y += 5;
        doc.text(`Cantidad: ${dataRow['Cantidad:'] || ''}`, 20, y); y += 5;
        doc.text(`Peso Estimado Total: ${dataRow['PESO ESTIMADO TOTAL:'] || ''}`, 20, y); y += 5;
        doc.text(`Lotes: ${dataRow['Lotes:'] || ''}`, 20, y); y += 10;

        // Firma
        doc.setFontSize(12);
        doc.text('Recibí Conforme: ___________________________', 20, y); y += 10;

        // Pie
        doc.setFontSize(8);
        doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 20, 280);
        doc.text('Jurisdicción Rosario - Santa Fe. No válido como factura.', 20, 285);

        // Guardar
        doc.save(`remito_${dataRow['Número Interno:'] || (index + 1)}.pdf`);
    }
});
