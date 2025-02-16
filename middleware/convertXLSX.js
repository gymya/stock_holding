const convertXLSX = (jsonData) => {
  const XLSX = require('xlsx');
  const workbook = XLSX.utils.book_new();

  const ws = XLSX.utils.json_to_sheet(jsonData);
  XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');
  const excelBuffer = XLSX.write(workbook, {
    type: 'buffer',
    bookType: 'xlsx',
  });
  return excelBuffer;
};

module.exports = convertXLSX;
