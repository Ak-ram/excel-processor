const XLSX = require('xlsx');
const XlsxPopulate = require('xlsx-populate');

// Read the Excel file
const workbook = XLSX.readFile('./test.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// List of destinations to merge into "ديوان المديرية"
const mergeDestinations = [
  "وحدة الخدمات", "مفتش الدخلية", "ادارة شئون الخدمة", "م.المدير للوحدات",
  "م.المدير للشؤن المالية", "م.المدير للافراد والتدريب", "1قسم العلاقات",
  "1قسم الانضباط", "1قسم الاسلحة", "1قسم المعلومات والتوثيق", "1قسم الانشاءات",
  "م.المدير للامن العام", "1قسم التخطيط والمتابعة", "1قسم التحقيقات", "1قسم الرخص",
  "نائب م.الامن", "1قسم الرقابة الجنائية", "1قسم حقوق الانسان",
];

// Group data based on 'الجهة'
const groupedData = {};

function normalizeDestination(destination) {
  // Combine 'بوفيه' and 'مستشفى' into a single category
  if (destination === 'بوفيه' || destination === 'مستشفى') {
    return 'بوفيه_مستشفى';
  }

  // Combine 'مباحث الادارة' and 'السياسين' into a single category
  if (destination === 'مباحث الادارة' || destination === 'السياسين') {
    return 'مباحث_سياسين';
  }

  return destination;
}

function processRow(row) {
  let destination = normalizeDestination(row['الجهة']);

  // Check if the destination is in the mergeDestinations list or if it's undefined
  if (!destination || mergeDestinations.includes(destination)) {
    destination = 'ديوان المديرية';
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

        newWorksheet.column("A").width(35);
        newWorksheet.column("B").width(30);
        newWorksheet.column("C").width(15);
        newWorksheet.column("D").width(10);
    

      // Add styles to the worksheet (fill color for illustration)
      newWorksheet.cell('A1').value('الاسم');
      newWorksheet.cell('B1').value('الرقم القومى');
      newWorksheet.cell('C1').value('الجهة');
      newWorksheet.cell('D1').value('المبلغ');

      // Add data to the new worksheet
      groupedData[destination].forEach((row, rowIndex) => {
        newWorksheet.cell(`A${rowIndex + 2}`).value(row['الاسم']);
        newWorksheet.cell(`B${rowIndex + 2}`).value(row['الرقم القومى']);
        newWorksheet.cell(`C${rowIndex + 2}`).value(row['الجهة']);
        newWorksheet.cell(`D${rowIndex + 2}`).value(row['المبلغ']);
      });

      // Sum the values in the 'المبلغ' column
      const totalAmount = groupedData[destination].reduce(
        (total, row) => total + parseFloat(row['المبلغ'] || 0),
        0
      );

      // Add total row with total in the fourth column (index 3)
      newWorksheet.cell(`A${groupedData[destination].length + 2}`).value('الاجمالى');
      newWorksheet.cell(`B${groupedData[destination].length + 2}`).value(totalAmount.toFixed(2));

      // Save the new workbook
      return workbook.toFileAsync(`${destination}.xlsx`);
    })
    .then(() => {
      console.log(`Styles added to ${destination}.xlsx`);
    })
    .catch((error) => {
      console.error(`Error adding styles to ${destination}.xlsx:`, error.message);
    });
}
