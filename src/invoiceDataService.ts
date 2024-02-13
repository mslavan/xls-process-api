import * as xlsx from 'xlsx';
import { InvoiceDataUtils, FieldColumnMapping, CurrencyRates } from './utils/invoiceDataUtils';

export class InvoiceDataService {
  private mandatoryFields: string[];
  private fieldColumnMapping: FieldColumnMapping;
  private currencyRates: CurrencyRates;

  constructor(mandatoryFields: string[]) {
    this.mandatoryFields = mandatoryFields.map(field => field.toLowerCase());
    this.fieldColumnMapping = {};
    this.currencyRates = {};
  }

  private isValidDateFormat(dateString: string): boolean {
    const regex = /^\d{4}-\d{2}$/; // YYYY-MM format
    return regex.test(dateString);
  }

  private parseDateString(dateString: string): Date {
    return new Date(dateString);
  }

  setFieldColumnMapping(fieldColumnMapping: FieldColumnMapping): void {
    this.fieldColumnMapping = fieldColumnMapping;
  }

  setCurrencyRates(currencyRates: { [p: string]: number }): void {
    this.currencyRates = currencyRates;
  }

  getInvoicingMonth(sheet: xlsx.WorkSheet): string | undefined {
    return InvoiceDataUtils.getInvoicingMonth(sheet);
  }

  validateFile(file: Express.Multer.File, mandatoryFields: string[], invoicingMonthParam?: string): void {
    // Ensure a file is uploaded
    if (!file) {
      throw new Error('No file uploaded');
    }

    // Ensure invoicingMonthParam is provided
    if (!invoicingMonthParam) {
      throw new Error('Invoicing date parameter required');
    }

    // Read the Excel workbook
    const workbook = xlsx.read(file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Find the header row in the sheet
    const headerRow = this.findInvoicesDataHeaderRow(sheet);

    // Ensure the header row is found
    if (!headerRow.startRow || !headerRow.columns) {
      throw new Error('Invalid file structure. Unable to find the required header row.');
    }

    // Get the invoicing date from the sheet
    const invoicingDateFile = this.getInvoicingMonth(sheet);

    // Check if the date format is valid and matches the expected format
    if (!invoicingDateFile ||
      this.isValidDateFormat(invoicingDateFile) ||
      this.parseDateString(invoicingDateFile).toISOString().slice(0, 7) !== invoicingMonthParam) {
      throw new Error(`Invalid or mismatched invoicing date format. Expected format: YYYY-MM`);
    }
  }

  findInvoicesDataHeaderRow(sheet: xlsx.WorkSheet): InvoiceDataUtils.InvoicesDataHeaderInfo {
    return InvoiceDataUtils.findInvoicesDataHeaderRow(sheet, this.mandatoryFields);
  }
  extractCurrencyRates(sheet: xlsx.WorkSheet): CurrencyRates {
    return InvoiceDataUtils.extractCurrencyRates(sheet);
  }

  processInvoicesData(params: InvoiceDataUtils.BasicProcessParams): any[] {
    const { sheet, startRow, columns } = params;
    return InvoiceDataUtils.processInvoicesData({
      sheet,
      startRow,
      columns,
      fieldColumnMapping: this.fieldColumnMapping,
      mandatoryFields: this.mandatoryFields,
      currencyRates: this.currencyRates
    });
  }
}
