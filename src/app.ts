import express, { Request, Response } from 'express';
import multer from 'multer';
import * as xlsx from 'xlsx';
import { InvoiceDataService } from './invoiceDataService';
import { InvoiceDataUtils } from './utils/invoiceDataUtils';

const app = express();
const port = 3000;

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post('/api/v1/upload', upload.single('file'), (req: Request, res: Response) => {
  try {
    if (!req.file) {
      throw new Error('No file uploaded');
    }

    const mandatoryFields = [
      'Customer',
      'Cust No',
      'Project Type',
      'Quantity',
      'Price Per Item',
      'Price Currency',
      'Total Price',
      'Invoice Currency',
      'Status',
    ];

    const dataService = new InvoiceDataService(mandatoryFields);

    const invoicingMonthParam = req.body.invoicingMonth as string;
    dataService.validateFile(req.file, mandatoryFields, invoicingMonthParam);

    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const { startRow, columns, fieldColumnMapping } = dataService.findInvoicesDataHeaderRow(sheet);
    if (!startRow || !columns) {
      return res.status(400).json({ success: false, error: 'Unable to find the start row of invoices data' });
    }
    dataService.setFieldColumnMapping(fieldColumnMapping);


    const currencyRates = dataService.extractCurrencyRates(sheet);
    dataService.setCurrencyRates(currencyRates);

    const InvoicingMonth = dataService.getInvoicingMonth(sheet);
    const invoicesData = dataService.processInvoicesData({
      sheet,
      startRow,
      columns,
    });

    res.status(200).json({
      InvoicingMonth,
      currencyRates,
      invoicesData,
    });
  } catch (error) {
    console.error(error);
    res.status(400).json({ success: false, error: error instanceof Error ? error.message : 'Unknown error occurred' });
  }
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
