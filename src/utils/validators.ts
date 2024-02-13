import * as xlsx from 'xlsx';

export function validateFile(file: Express.Multer.File): void {
  if (!file) {
    throw new Error('No file uploaded');
  }
}
