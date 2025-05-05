document.getElementById('generateBtn').addEventListener('click', async function() {
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
            doc.addPage(); // 👉 Nueva página para cada remito
        }

        doc.setFontSize(16);
        const nro = row['Número Interno:'] || `(sin número)`;
        doc.text(`Remito N° ${nro}`, 105, 15, { align: 'center' });

        doc.setFontSize(10);
        let y = 25;

        // 🔁 Mostrar TODAS las columnas (de la A a la Z, cualquier nombre)
        for (const key in row) {
            const cleanKey = key.trim();
            const value = String(row[key]).trim();
            doc.text(`${cleanKey}: ${value}`, 20, y);
            y += 6;

            // ⚠️ Si se acerca al final de la hoja, agregar página nueva
            if (y > 270) {
                doc.addPage();
                y = 20;
            }
        }

        // Firma
        y += 6;
        doc.setFontSize(12);
        doc.text('Recibí Conforme: ___________________________', 20, y);

        // Pie
        doc.setFontSize(8);
        doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 20, 280);
        doc.text('Seguro de mercaderia por cuenta de Syngenta', 20, 285);
    }

    doc.save('Remitos_Completos.pdf');
});
