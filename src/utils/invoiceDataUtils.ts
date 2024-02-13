import * as xlsx from 'xlsx';

export interface FieldColumnMapping {
  [field: string]: string;
}
export interface CurrencyRates {
  [p: string]: number;
}

export module InvoiceDataUtils {
  export function validateFile(file: Express.Multer.File): void {
    if (!file) {
      throw new Error('No file uploaded');
    }
  }

  export interface InvoicesDataHeaderInfo {
    startRow: number | undefined;
    columns: string[] | undefined;
    fieldColumnMapping: { [field: string]: string };
  }

  export interface ProcessInvoicesDataParams {
    sheet: xlsx.WorkSheet;
    startRow: number;
    columns: string[];
    fieldColumnMapping: FieldColumnMapping;
    mandatoryFields: string[];
    currencyRates: CurrencyRates;
  }

  export interface BasicProcessParams {
    sheet: xlsx.WorkSheet;
    startRow: number;
    columns: string[];
  }

  function isRelevantLine(record: any): boolean {
    return (record['status']?.toLowerCase() === 'ready') || record['invoice #'];
  }

  export function findInvoicesDataHeaderRow(sheet: xlsx.WorkSheet, mandatoryFields: string[]): { startRow: number | undefined, columns: string[] | undefined, fieldColumnMapping: { [field: string]: string } } {
    const lastRow = xlsx.utils.decode_range(sheet['!ref'] || 'A1').e.r + 1;

    for (let row = 1; row <= lastRow; row++) {
      const foundFields: string[] = [];
      const fieldColumnMapping: { [field: string]: string } = {};

      // Iterate over each column in the row
      let col = 0;
      while (true) {
        const cellAddress = xlsx.utils.encode_cell({ r: row, c: col });
        const cellValue = sheet[cellAddress] ? xlsx.utils.format_cell(sheet[cellAddress]) : '';

        if (!cellValue) {
          break;
        }

        const lowercaseColumnName = cellValue.toLowerCase();

        foundFields.push(lowercaseColumnName);

        // Check if the current column corresponds to any mandatory field
        const matchingField = mandatoryFields.find(field => lowercaseColumnName.includes(field.toLowerCase()));
        if (matchingField) {
          fieldColumnMapping[matchingField] = lowercaseColumnName;
        }
        col++;
      }

      if (mandatoryFields.every(field => foundFields.some(found => found.includes(field)))) {
        return { startRow: row + 1, columns: foundFields, fieldColumnMapping };
      }
    }
    return { startRow: undefined, columns: undefined, fieldColumnMapping: {} };
  }


  export function extractCurrencyRates(sheet: xlsx.WorkSheet): { [currency: string]: number } {
    const currencyRates: { [currency: string]: number } = {};

    const lastRow = xlsx.utils.decode_range(sheet['!ref'] || 'A1').e.r + 1;

    for (let row = 2; row <= lastRow; row++) {
      const cellAddressRate = xlsx.utils.encode_cell({ r: row, c: 0 });
      const cellAddressValue = xlsx.utils.encode_cell({ r: row, c: 1 });


      const cellValueRate = sheet[cellAddressRate] ? xlsx.utils.format_cell(sheet[cellAddressRate]) : '';
      const cellValueValue = sheet[cellAddressValue] ? xlsx.utils.format_cell(sheet[cellAddressValue]) : '';


      const isValidRate = cellValueRate.toLowerCase().includes('rate') && !isNaN(Number(cellValueValue));


      if (isValidRate) {
        const currency = cellValueRate.toLowerCase().replace('rate', '').trim().toUpperCase();
        currencyRates[currency] = Number(cellValueValue);
      } else {
        break;
      }
    }

    return currencyRates;
  }

  export function processInvoicesData(params: ProcessInvoicesDataParams): any[] {
    const { sheet, startRow, columns, fieldColumnMapping, currencyRates } = params;
    const lastRow = xlsx.utils.decode_range(sheet['!ref'] || 'A1').e.r + 1;
    const invoicesData: any[] = [];

    for (let row = startRow; row <= lastRow; row++) {
      const rowData: { [key: string]: any } = {};

      columns.forEach((column, index) => {
        const cellAddress = xlsx.utils.encode_cell({ r: row, c: index });
        const cellValue = sheet[cellAddress] ? xlsx.utils.format_cell(sheet[cellAddress]) : '';

        // Use the fieldColumnMapping to get the corresponding mandatory field
        const mandatoryField = Object.keys(fieldColumnMapping).find(field => fieldColumnMapping[field] === column);
        if (mandatoryField) {
          rowData[mandatoryField] = cellValue;
        }
      });

      // Calculate Invoice Total and add it to the record
      const totalPrice = parseFloat(rowData['total price']) || 0;
      const invoiceCurrency = rowData['invoice currency']?.toUpperCase();
      const currencyRate = currencyRates[invoiceCurrency] || 1;
      const invoiceTotal = totalPrice * currencyRate;

      rowData['invoice total'] = invoiceTotal;

      // Check if the line is relevant
      if (isRelevantLine(rowData)) {
        invoicesData.push(rowData);
      }
    }

    // Set validationErrors
    postProcessInvoicesData(invoicesData, params);

    return (invoicesData);
  }

  export function postProcessInvoicesData(invoicesData: any[], params: ProcessInvoicesDataParams): void {
    invoicesData.forEach((record) => {
      record.validationErrors = validateInvoiceRecord(record, params);
    });
  }


  export function validateInvoiceRecord(record: any, params: ProcessInvoicesDataParams): string[] {
    const { mandatoryFields } = params;
    const validationErrors: string[] = [];

    // Check if all mandatory fields are present in the record
    for (const field of mandatoryFields) {
      if (!record[field]) {
        validationErrors.push(`Missing required field: ${field}`);
      }
    }

    return validationErrors;
  }

  export function getInvoicingMonth(sheet: xlsx.WorkSheet): string | undefined {
    const cellAddress = xlsx.utils.encode_cell({ r: 0, c: 0 });
    const cellValue = sheet[cellAddress] ? xlsx.utils.format_cell(sheet[cellAddress]) : undefined;
    return cellValue;
  }
}
