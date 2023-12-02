# Excel Processor 📊

This script processes an Excel file, groups data based on a specific column, and creates new Excel files for each group. It also generates a summary sheet.

## Getting Started 🚀

### Prerequisites

- Node.js installed on your machine 🌐
- npm package manager 📦

### Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/your-repo.git
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

### Usage

Run the script with the following command:

```bash
node excelProcessor.js /path/to/your/excel/file.xlsx
```

Replace `/path/to/your/excel/file.xlsx` with the actual path to your Excel file.

## Code Overview 🧐

The script uses Node.js and several npm packages, including XLSX and XlsxPopulate, to process Excel files. It performs the following steps:

1. Reads the Excel file.
2. Groups data based on a specified column (in this case, 'الجهة').
3. Creates a new workbook for each group, applying styling and adding a summary row.
4. Generates a summary sheet with information about each created sheet.

### Folder Structure 📁

- `js/`: Contains utility functions and data used by the main script.
- `excelProcessor.js`: The main script for processing Excel files.

### Customization ⚙️

- Modify the script to suit your specific Excel file structure.
- Adjust the path to the Excel file, column names, and other parameters as needed.

## License 📜

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.


