const XLSX = require("xlsx");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const { styleColumn } = require("./js/utils");

let date = new Date().toLocaleDateString("ar-eg", {
  month: "long",
  year: "numeric",
});

const {
  style_data,
  style_header,
  mergeDestinations,
  combinations,
} = require("./js/data");

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

// Create an array to store information about each created sheet
const createdSheetsInfo = [];

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
      const columnWidths = [35, 30, 15, 10, 40];
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

      newWorksheet.cell("A1").value("م").style(style_header);
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
      // Save the new workbook in th` "${}مرتبات شهر  directory
      const filePath = `./مرتبات شهر ${date}/${destination}.xlsx`;

      // Ensure th` "${}مرتبات شهر " directory exists
      if (!fs.existsSync(`./مرتبات شهر ${date}`)) {
        fs.mkdirSync(`./مرتبات شهر ${date}`);
      }

      // Add information about the created sheet to the array
      createdSheetsInfo.push({
        destination,
        filePath,
        totalRows: groupedData[destination].length,
        totalAmount,
      });

      // Save the new workbook
      return workbook.toFileAsync(filePath);
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

// Create a summary sheet
const summaryWorkbook = XlsxPopulate.fromBlankAsync();
summaryWorkbook
  .then((workbook) => {
    const summarySheet = workbook.sheet(0);
    const summaryHeader = ["م", "الجهة", "العدد", "اجمالى المبلغ"];
    const summaryColumns = ["A", "B", "C", "D"];
    const columnWidths = [5, 35, 10, 20];
    summaryColumns.forEach((col, index) => {
      if (col !== "A") {
        summarySheet.column(col).width(columnWidths[index]);
      }
    });
    summarySheet.rightToLeft(true);
    // Add headers to the summary sheet
    // Merge cells for the title
    summarySheet.range("A1:D1").merged(true).style(style_header);

    // Add title to the merged cells
    summarySheet
      .cell("A1")
      .value(`اجمالى مرتبات شهر ${date}`)
      .style(style_header);

    summaryHeader.forEach((header, index) => {
      summarySheet
        .cell(`${summaryColumns[index]}2`)
        .value(header)
        .style(style_header);
    });

    // Add information about each created sheet to the summary sheet
    createdSheetsInfo.forEach((info, index) => {
      const rowIndex = index + 3;
      summarySheet
        .cell(`${summaryColumns[0]}${rowIndex}`)
        .value(index+1)
        .style(style_data);
      summarySheet
        .cell(`${summaryColumns[1]}${rowIndex}`)
        .value(info.destination)
        .style(style_data);
      // summarySheet.cell(`${summaryColumns[1]}${rowIndex}`).value(info.filePath).style(style_data);
      summarySheet
        .cell(`${summaryColumns[2]}${rowIndex}`)
        .value(info.totalRows.toLocaleString("ar-EG"))
        .style(style_data);
      summarySheet
        .cell(`${summaryColumns[3]}${rowIndex}`)
        .value(info.totalAmount.toLocaleString("ar-EG"))
        .style(style_data);
    });

    // Save the summary sheet
    const summaryFilePath = `./مرتبات شهر ${date}/Summary.xlsx`;
    return workbook.toFileAsync(summaryFilePath);
  })
  .then(() => {
    console.log(`Summary sheet created and saved: Summary.xlsx`);
  })
  .catch((error) => {
    console.error(
      `Error creating and saving the summary sheet:`,
      error.message
    );
  });
