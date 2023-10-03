import fs from 'fs';
import xml2js from 'xml2js';
import xlsx from 'xlsx';

// Function to convert XML data to Excel
export function convertXmlToExcel(xmlFilePath, excelFilePath, cb) {
  // Read XML file
  fs.readFile(xmlFilePath, 'utf-8', (err, xmlData) => {
    if (err) {
      cb(err);
      return;
    }
    // Parse XML to JS object
    const parser = new xml2js.Parser({
      explicitArray: false,
      mergeAttrs: true,
    });
    parser.parseString(xmlData, (err, jsObject) => {
      if (err) {
        cb(err);
        return;
      }

      // Function to add formatting
      function formatRange(sheet, range, style) {
        for (let row = range.s.r; row <= range.e.r; row++) {
          for (let col = range.s.c; col <= range.e.c; col++) {
            const reference = xlsx.utils.encode_cell({ r: row, c: col });
            if (!sheet[reference]) continue;
            sheet[reference].s = style;
          }
        }
      }

      try {
        //Define movies and series data from the parsed JavaScript object
        const moviesData = jsObject.binge.movies.movie;
        const seriesData = jsObject.binge.series.serie;

        // Create worksheets for movies and series
        const moviesWs = xlsx.utils.json_to_sheet(moviesData);
        const seriesWs = xlsx.utils.json_to_sheet(seriesData);

        // Create new Workbook
        const wb = xlsx.utils.book_new();

        // Add worksheets to workbook
        xlsx.utils.book_append_sheet(wb, moviesWs, 'Movies');
        xlsx.utils.book_append_sheet(wb, seriesWs, 'Series');

        // Define cell style
        const bold = { font: { bold: true } };

        // Apply formatting
        const headerRange = xlsx.utils.decode_range(moviesWs['!ref']);
        formatRange(moviesWs, headerRange, bold);

        // Save the workbook
        xlsx.writeFile(wb, excelFilePath);
        cb(null, 'Success');
      } catch (error) {
        cb(error);
      }
    });
  });
}
