
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

      // 🔁 Función para convertir fechas
      const convertirFecha = (valor) => {
        if (!valor) return '';
        const numero = Number(valor);
        if (!isNaN(numero)) {
          const epoch = new Date(Date.UTC(1899, 11, 30));
          const fecha = new Date(epoch.getTime() + numero * 86400000);
          return `${String(fecha.getDate()).padStart(2, '0')}/${String(fecha.getMonth() + 1).padStart(2, '0')}/${fecha.getFullYear()}`;
        }
        if (typeof valor === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(valor)) {
          const [y, m, d] = valor.split('-');
          return `${d}/${m}/${y}`;
        }
        return valor;
      };

      // 🔢 Número base para remitos autoincrementales
      let ultimoRemito = 24291;
      const prefijoRemito = '0283-';

      const doc = new jsPDF();

      json.forEach((originalRow, index) => {
        if (!originalRow || Object.keys(originalRow).length === 0) return;
        if (index !== 0) doc.addPage();

        // 🔧 Normalizar claves y forzar valores a string
        const row = {};
        for (const [key, value] of Object.entries(originalRow)) {
          const cleanKey = key
            .replace(/[:]+$/g, '')
            .replace(/\s{2,}/g, ' ')
            .replace(/[:]/g, '')
            .trim();
          row[cleanKey] = value != null ? value.toString() : '';
        }

        const usados = new Set();

        // === ENCABEZADO ===
        const numeroRemito = ultimoRemito + index;
        const remitoFormateado = `${prefijoRemito}${String(numeroRemito).padStart(8, '0')}`;

        doc.setFontSize(16);
        doc.text(`Remito N° ${remitoFormateado}`, 105, 15, { align: 'center' });

        doc.setFontSize(12);
        doc.text(`Número Interno: ${row['Número Interno'] || ''}`, 105, 22, { align: 'center' });

        const fechaEmision = convertirFecha(row['Fecha de Emisión']);
        doc.text(`Fecha de Emisión: ${fechaEmision}`, 105, 29, { align: 'center' });

        usados.add('Número Interno');
        usados.add('Fecha de Emisión');

        doc.setFontSize(10);
        let y = 40;

        // === SYNGENTA (con valores fijos) ===
        const valoresFijos = {
          'C.U.I.T.': '30-64632845-0',
          'Ingresos Brutos (CM)': '901-962580-1',
          'Inicio de actividades': '31/12/1991',
          'I.V.A.': 'Responsable Inscripto'
        };

        const camposSyngenta = [
          'Nro. Transporte',
          'Transporte',
          'C.U.I.T.',
          'Ingresos Brutos (CM)',
          'Inicio de actividades',
          'I.V.A.',
          'Fecha de Vencimiento del C.A.I.',
          'C.A.I. Nº'
        ];
        camposSyngenta.forEach(campo => {
          let valor = valoresFijos[campo] || row[campo] || '';
          if (
            campo.toLowerCase().includes('fecha') ||
            campo.toLowerCase().includes('inicio de actividades')
          ) {
            valor = convertirFecha(valor);
          }
          doc.text(`${campo}: ${valor}`, 20, y);
          usados.add(campo);
          y += 6;
        });

        // === EMISOR ===
        const camposEmisor = [
          'Cliente',
          'Deposito Origen',
          'Dirección receptor',
          'Teléfono Recptor',
          'Pedido'
        ];
        camposEmisor.forEach(campo => {
          doc.text(`${campo}: ${row[campo] || ''}`, 20, y);
          usados.add(campo);
          y += 6;
        });

        // === RECEPTOR ===
        const camposReceptor = [
          'Deposito Destino',
          'Código de Cliente',
          'Cliente Receptor',
          'Dirección receptor',
          'C.U.I.T. Receptor',
          'Pedido'
        ];
        camposReceptor.forEach(campo => {
          doc.text(`${campo}: ${row[campo] || ''}`, 20, y);
          usados.add(campo);
          y += 6;
        });

        // === PRODUCTOS ===
        doc.setFontSize(12);
        doc.text('Productos:', 20, y); y += 8;
        doc.setFontSize(10);
        const camposProducto = [
          'Código',
          'Descripción',
          'Cantidad',
          'Lotes',
          'Peso estimado Total'
        ];
        camposProducto.forEach(campo => {
          doc.text(`${campo}: ${row[campo] || ''}`, 20, y);
          usados.add(campo);
          y += 6;
        });

        // === OTROS CAMPOS ===
        doc.setFontSize(10);
        doc.text('Otros campos:', 20, y); y += 6;

        for (const key in row) {
          if (usados.has(key)) continue;
          let valor = row[key];
          if (
            key.toLowerCase().includes('fecha') ||
            key.toLowerCase().includes('inicio de actividades')
          ) {
            valor = convertirFecha(valor);
          }
          doc.text(`${key}: ${valor}`, 20, y);
          y += 6;

          if (y > 270) {
            doc.addPage();
            y = 20;
          }
        }

        // === FIRMA Y PIE ===
        y += 6;
        doc.setFontSize(12);
        doc.text('Recibí Conforme: ___________________________', 20, y); y += 10;

        doc.setFontSize(8);
        doc.text('La mercadería será transportada bajo exclusiva responsabilidad del transportista.', 20, 280);
        doc.text('Seguro de mercadería por cuenta de Syngenta.', 20, 285);
      });

      doc.save('Remitos_Syngenta.pdf');
    });


