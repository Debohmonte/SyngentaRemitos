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
  
    const zip = new JSZip();
  
    json.forEach((row, index) => {
      if (!row['Cliente:.1']) return; // Saltear filas vacías
  
      const xmlContent = `
  <remito>
    <fecha_emision>${row['Fecha de Emisión:']}</fecha_emision>
    <numero_interno>${row['Número Interno:']}</numero_interno>
    <cliente>${row['Cliente:.1']}</cliente>
    <direccion>${row['Dirección:']}</direccion>
    <telefono>${row['Teléfono:']}</telefono>
    <producto>
      <codigo>${row['Código:']}</codigo>
      <descripcion>${row['Descripción:']}</descripcion>
      <cantidad>${row['Cantidad:']}</cantidad>
      <peso>${row['PESO ESTIMADO TOTAL: ']}</peso>
      <lotes>${row['Lotes:']}</lotes>
    </producto>
    <transporte>
      <numero>${row['Nro. Transporte:']}</numero>
      <nombre>${row['Transporte:']}</nombre>
    </transporte>
  </remito>
  `.trim();
  
      const fileName = `remito_${index + 1}.xml`;
      zip.file(fileName, xmlContent);
    });
  
    // Descargar ZIP
    const content = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(content);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'remitos.zip';
    a.click();
  });
  