const XLSX = require('xlsx');

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
XLSX.utils.sheet_to_json(worksheet, { raw: true }).forEach(processRow);

// Create a new workbook for each group
for (const destination in groupedData) {
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(groupedData[destination]);

  // Sum the values in the 'المبلغ' column
  const totalAmount = groupedData[destination].reduce((total, row) => total + parseFloat(row['المبلغ'] || 0), 0);

  // Add total row with total in the fourth column (index 3)
  const totalRow = {
    الاسم: 'الاجمالى',
    المبلغ: '', // Empty for formatting purposes
    الرقم_القومي: '', // Add columns based on your data structure
    الجهة: totalAmount.toFixed(2), // Round to 2 decimal places
    التوقيع: '', // Add columns based on your data structure
  };
  XLSX.utils.sheet_add_json(newWorksheet, [totalRow], { skipHeader: true, origin: -1 });

  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

  // Save the new workbook
  XLSX.writeFile(newWorkbook, `${destination}.xlsx`);
}
