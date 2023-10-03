import fs from 'fs';
import xml2js from 'xml2js';
import xlsx from 'xlsx';

// Function to convert XML data to Excel
export const convertXmlToExcel = async (xmlFilePath, excelFilePath) => {
  try {
    // Read XML file
    const xmlData = await fs.promises.readFile(xmlFilePath, 'utf-8');
    // Parse XML to JS object
    const parser = new xml2js.Parser({
      explicitArray: false,
      mergeAttrs: true,
    });
    parser.parseString(xmlData, (err, jsObject) => {
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
      // formatRange(moviesWs, headerRange, bold);

      // Save the workbook
      xlsx.writeFile(wb, excelFilePath);
      console.log('Successfully converted XML to Excel');
    });
  } catch (error) {
    console.log("Error in Excel conversion: ", error);
  }
};
