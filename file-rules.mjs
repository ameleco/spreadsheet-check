export function fileError(file) {
  if (!file) return 'Choose your spreadsheet first.';
  const name = file.name.toLowerCase();
  if (name.endsWith('.pdf') || file.type === 'application/pdf') return 'PDF files can’t be graded here. Export the original spreadsheet as an Excel (.xlsx) file, then try again.';
  if (!name.endsWith('.xlsx')) return 'Please upload an Excel (.xlsx) spreadsheet. Use File → Download → Microsoft Excel (.xlsx) in Google Sheets.';
  if (file.size === 0) return 'This file is empty. Save your spreadsheet and choose it again.';
  if (file.size > 10 * 1024 * 1024) return 'This file is larger than 10 MB. Remove unnecessary images or sheets, or ask your instructor to review it.';
  return '';
}
