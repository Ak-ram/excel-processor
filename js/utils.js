exports.styleColumn = (newWorksheet, column,rowIndex, style) => {
  newWorksheet.cell(`${column}${rowIndex + 1}`).width;
  newWorksheet.cell(`${column}${rowIndex + 1}`).style(style);
};
