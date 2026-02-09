const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');

module.exports = async (req, res) => {
  // Solo POST
  if (req.method === 'GET') {
    return res.status(200).json({ status: 'ok', message: 'SF Automation Processor Test' });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { baseName, files, fileContents } = req.body;

    if (!fileContents || typeof fileContents !== 'object') {
      return res.status(400).json({
        success: false,
        error: 'No se recibieron archivos. Asegurate de activar "Enviar contenido de archivos" en la configuracion del procesador.'
      });
    }

    // Buscar internal_all en los archivos recibidos
    const internalFileName = Object.keys(fileContents).find(name =>
      name.toLowerCase().includes('internal_all')
    );

    if (!internalFileName) {
      return res.status(400).json({
        success: false,
        error: 'No se encontro internal_all.xlsx en los archivos enviados. Archivos recibidos: ' + Object.keys(fileContents).join(', ')
      });
    }

    // Decodificar Excel de base64
    const excelBuffer = Buffer.from(fileContents[internalFileName], 'base64');

    // Leer con ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelBuffer);

    const sheet = workbook.worksheets[0];
    if (!sheet) {
      return res.status(400).json({ success: false, error: 'El archivo Excel no tiene hojas' });
    }

    // Contar URLs (filas con datos, excluyendo header)
    const totalRows = sheet.rowCount;
    const urlCount = Math.max(0, totalRows - 1); // Restar header

    // Recopilar algunos datos adicionales para el informe
    const sheetName = sheet.name;
    const columnCount = sheet.columnCount;

    // === Generar Excel modificado con hoja extra ===
    const newSheet = workbook.addWorksheet('Conteo URLs');

    // Estilos
    newSheet.columns = [
      { header: 'Metrica', key: 'metric', width: 35 },
      { header: 'Valor', key: 'value', width: 20 }
    ];

    // Header con estilo
    const headerRow = newSheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B82F6' } };

    // Datos
    newSheet.addRow({ metric: 'Total de URLs', value: urlCount });
    newSheet.addRow({ metric: 'Hoja de origen', value: sheetName });
    newSheet.addRow({ metric: 'Columnas en el archivo', value: columnCount });
    newSheet.addRow({ metric: 'Archivo procesado', value: internalFileName });
    newSheet.addRow({ metric: 'Fecha de procesado', value: new Date().toISOString() });
    newSheet.addRow({ metric: 'Procesado por', value: 'Vercel Processor Test' });

    // Estilo para filas de datos
    for (let i = 2; i <= 7; i++) {
      const row = newSheet.getRow(i);
      row.getCell(1).font = { bold: true };
      if (i % 2 === 0) {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
      }
    }

    // Guardar Excel a buffer
    const excelOutputBuffer = await workbook.xlsx.writeBuffer();
    const excelBase64 = Buffer.from(excelOutputBuffer).toString('base64');

    // === Generar PDF ===
    const pdfBuffer = await generatePDF(urlCount, baseName, internalFileName);
    const pdfBase64 = pdfBuffer.toString('base64');

    // Nombres de archivos de salida
    const excelOutputName = `${baseName || 'resultado'}_conteo_urls.xlsx`;
    const pdfOutputName = `${baseName || 'resultado'}_conteo_urls.pdf`;

    return res.status(200).json({
      success: true,
      message: `Procesado completado: ${urlCount} URLs encontradas`,
      urlCount,
      generatedFiles: {
        [excelOutputName]: excelBase64,
        [pdfOutputName]: pdfBase64
      }
    });

  } catch (error) {
    console.error('Error procesando:', error);
    return res.status(500).json({
      success: false,
      error: `Error interno: ${error.message}`
    });
  }
};

/**
 * Genera un PDF sencillo con el conteo de URLs
 */
function generatePDF(urlCount, baseName, fileName) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ margin: 50 });
      const chunks = [];

      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      // Titulo
      doc.fontSize(22).fillColor('#1e293b').text('Conteo de URLs', { align: 'center' });
      doc.moveDown(0.5);
      doc.fontSize(10).fillColor('#64748b').text('Generado por SF Automation - Procesador de prueba Vercel', { align: 'center' });
      doc.moveDown(2);

      // Linea separadora
      doc.moveTo(50, doc.y).lineTo(545, doc.y).strokeColor('#e2e8f0').stroke();
      doc.moveDown(1.5);

      // Numero grande
      doc.fontSize(72).fillColor('#3b82f6').text(String(urlCount), { align: 'center' });
      doc.fontSize(16).fillColor('#475569').text('URLs encontradas', { align: 'center' });
      doc.moveDown(2);

      // Linea separadora
      doc.moveTo(50, doc.y).lineTo(545, doc.y).strokeColor('#e2e8f0').stroke();
      doc.moveDown(1.5);

      // Detalles
      doc.fontSize(11).fillColor('#334155');
      doc.text(`Rastreo: ${baseName || 'N/A'}`);
      doc.moveDown(0.3);
      doc.text(`Archivo: ${fileName}`);
      doc.moveDown(0.3);
      doc.text(`Fecha: ${new Date().toLocaleString('es-ES')}`);

      doc.end();
    } catch (err) {
      reject(err);
    }
  });
}
