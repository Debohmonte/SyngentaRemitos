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

      // 🔁 Conversor de fechas
      const convertirFecha = (valor) => {
        if (!valor) return '';
        // número de Excel
        if (!isNaN(valor) && typeof valor === 'number') {
          const epoch = new Date(Date.UTC(1899, 11, 30));
          const fecha = new Date(epoch.getTime() + valor * 86400000);
          return `${String(fecha.getDate()).padStart(2, '0')}/${String(fecha.getMonth() + 1).padStart(2, '0')}/${fecha.getFullYear()}`;
        }
        // texto tipo "2024-03-01"
        if (typeof valor === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(valor)) {
          const [y, m, d] = valor.split('-');
          return `${d}/${m}/${y}`;
        }
        return valor;
      };

      const doc = new jsPDF();

      json.forEach((row, index) => {
        if (!row || Object.keys(row).length === 0) return;
        if (index !== 0) doc.addPage();

        const usados = new Set();

        // === ENCABEZADO ===
        doc.setFontSize(16);
        doc.text(`Remito N° ${row['Remito N°:'] || '(sin número)'}`, 105, 15, { align: 'center' });

        doc.setFontSize(12);
        doc.text(`Número Interno: ${row['Número Interno:'] || ''}`, 105, 22, { align: 'center' });

        doc.setFontSize(10);
        let y = 30;

        // === Transporte + Fecha Emisión
        const transporte = row['Transporte:'] || '';
        const fechaEmision = convertirFecha(row['Fecha de Emisión:']);
        doc.text(`Transporte: ${transporte}`, 20, y); y += 6;
        doc.text(`Fecha de Emisión: ${fechaEmision}`, 20, y); y += 6;
        usados.add('Transporte:');
        usados.add('Fecha de Emisión:');

        // === Syngenta ===
        const camposFijos = [
          'C.U.I.T.:',
          'Ingresos Brutos (CM):',
          'Inicio de actividades:',
          'I.V.A.:',
          'Fecha de Vencimiento del C.A.I.:',
          'C.A.I. Nº:'
        ];
        camposFijos.forEach(campo => {
          let valor = row[campo] || '';
          if (campo.toLowerCase().includes('fecha')) valor = convertirFecha(valor);
          doc.text(`${campo} ${valor}`, 20, y);
          usados.add(campo);
          y += 6;
        });

        // === Emisor ===
        const camposEmisor = [
          'Cliente Recptor:',
          'Deposito Origen',
          'Dirección receptor:',
          'Teléfono Recptor:',
          'Pedido:',
          'Nro. Transporte:'
        ];
        camposEmisor.forEach(campo => {
          doc.text(`${campo} ${row[campo] || ''}`, 20, y);
          usados.add(campo);
          y += 6;
        });

        // === Receptor ===
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

        // === Productos ===
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

        // === Otros campos dinámicos ===
        doc.setFontSize(10);
        doc.text('Otros campos:', 20, y); y += 6;

        for (const key in row) {
          if (usados.has(key)) continue;

          let valor = row[key];
          if (key.toLowerCase().includes('fecha')) valor = convertirFecha(valor);

          doc.text(`${key}: ${valor}`, 20, y);
          y += 6;

          if (y > 270) {
            doc.addPage();
            y = 20;
          }
        }

        // === Firma y pie ===
        y += 6;
        doc.setFontSize(12);
        doc.text('Recibí Conforme: ___________________________', 20, y); y += 10;

        doc.setFontSize(8);
        doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 20, 280);
        doc.text('Seguro de mercadería por cuenta de Syngenta.', 20, 285);
      });

      doc.save('Remitos_Syngenta.pdf');
    });

