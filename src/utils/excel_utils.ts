import { Workbook } from 'exceljs';
import * as fs from 'fs';

export class ExcelAccessor {
  private workbook: Workbook;
  private filePath: string;

  constructor(filePath: string) {
    this.filePath = filePath;
    this.workbook = new Workbook();
    this.loadWorkbook();
  }

  private async loadWorkbook(): Promise<void> {
    await this.workbook.xlsx.readFile(this.filePath);
  }

  /**
   * Get the value at a specific cell.
   *
   * @param row The row index (0-based).
   * @param col The column index (0-based).
   */
  public getValue(row: number, col: number): any {
    return this.workbook.getWorksheet('Sheet1')?.getCell(`A${row}`, `B${col}`).value;
  }

  /**
   * Set the value at a specific cell.
   *
   * @param row The row index (0-based).
   * @param col The column index (0-based).
   * @param value The new value for the cell.
   */
  public setValue(row: number, col: number, value: any): void {
    const worksheet = this.workbook.getWorksheet('Sheet1');
    if(worksheet !== undefined){
      worksheet.getCell(`A${row}`, `B${col}`).value = value;
    }
  }

  /**
   * Add a new row to the end of the Excel file.
   *
   * @param values An array of values for the new row.
   */
  public addRow(values: any[]): void {
    const worksheet = this.workbook.getWorksheet('Sheet1');
    worksheet?.addRow(values);
  }

  /**
   * Save the changes made to the workbook.
   */
  public async saveChanges(): Promise<void> {
    await this.workbook.xlsx.writeFile(this.filePath);
  }
}

const accessor = new ExcelAccessor('path/to/excel/file.xlsx');

// Get the value at cell A1:B1
console.log(accessor.getValue(0, 0));

// Set the value at cell A1:B1
accessor.setValue(0, 0, 'New Value');

// Add a new row to the end of the Excel file
const newRow = ['Value1', 'Value2'];
accessor.addRow(newRow);

// Remove the first row from the Excel file

// Save the changes made to the workbook
accessor.saveChanges().then(() => {
  console.log('Changes saved successfully');
}).catch((err) => {
  console.error(err);
});