exports.styleColumn = (newWorksheet, column,rowIndex, style) => {
  newWorksheet.cell(`${column}${rowIndex + 2}`).width;
  newWorksheet.cell(`${column}${rowIndex + 2}`).style(style);
};
