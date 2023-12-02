const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");
const { styleColumn } = require("./js/utils");
const {
  style_data,
  style_header,
  mergeDestinations,
  combinations,
} = require("./js/data");
// let serialNumber = 1;
// Read the Excel file
const workbook = XLSX.readFile("./test.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Group data based on 'الجهة'
const groupedData = {};

function normalizeDestination(destination) {
  return combinations[destination] || destination;
}

function processRow(row) {
  let destination = normalizeDestination(row["الجهة"]);

  // Check if the destination is in the mergeDestinations list or if it's undefined
  if (!destination || mergeDestinations.includes(destination)) {
    destination = "ديوان المديرية";
  }

  if (!groupedData[destination]) {
    groupedData[destination] = [];
  }

  groupedData[destination].push(row);
}

// Iterate through each row and process it
const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: true });
jsonData.forEach(processRow);

// Create a new workbook for each group
for (const destination in groupedData) {
  const newWorkbook = XlsxPopulate.fromBlankAsync();

  newWorkbook
    .then((workbook) => {
      const newWorksheet = workbook.sheet(0);
      const headerMappings = [
        "الاسم",
        "الرقم القومى",
        "الجهة",
        "المبلغ",
        "التوقيع",
      ];

      // Set column widths
      const columnWidths = [35, 30, 15, 10, 35];
      const columns = ["B", "C", "D", "E", "F"];
      columns.forEach((col, index) =>
        newWorksheet.column(col).width(columnWidths[index])
      );

      // Add headers
      headerMappings.forEach((header, index) =>
        newWorksheet
          .cell(`${columns[index]}1`)
          .value(header)
          .style(style_header)
      );

      let serialNumber = 1;
      groupedData[destination].forEach((row, rowIndex) => {
        // Increment the serial number for each row
        const currentSerialNumber = serialNumber++;
        // Add the serial number to the first column
        newWorksheet
          .cell("A" + (rowIndex + 2))
          .value(currentSerialNumber)
          .style(style_data);
        // Shift the rest of the columns to the right
        columns.forEach((col, index) => {
          const value = row[headerMappings[index]];
          const formattedValue = value ? value.toLocaleString("ar-EG") : "";
          newWorksheet
            .cell(`${col}${rowIndex + 2}`)
            .value(formattedValue)
            .style({ numberFormat: "0.00" });
          styleColumn(newWorksheet, col, rowIndex + 1, style_data);
        });
      });
      // Sum the values in the 'المبلغ' column
      const totalAmount = groupedData[destination].reduce(
        (total, row) => total + parseFloat(row["المبلغ"] || 0),
        0
      );

      const lastRowIndex = groupedData[destination].length + 2;

      // Add total row with total in the fourth column (index 3)
      const totalRowHeaders = [
        "الاجمالى",
        "",
        "",
        totalAmount.toLocaleString("ar-EG"),
        "",
      ];
      const totalRowColumns = ["B", "C", "D", "E", "F"];

      totalRowColumns.forEach((col, index) => {
        newWorksheet
          .cell(`${col}${lastRowIndex}`)
          .value(totalRowHeaders[index])
          .style(style_header);
      });
      newWorksheet.rightToLeft(true);

      // Save the new workbook
      return workbook.toFileAsync(`${destination}.xlsx`);
    })
    .then(() => {
      console.log(`Styles and data added to ${destination}.xlsx`);
    })
    .catch((error) => {
      console.error(
        `Error adding styles and data to ${destination}.xlsx:`,
        error.message
      );
    });
}
