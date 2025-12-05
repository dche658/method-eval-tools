const excelRangeRegex = "^([a-zA-Z\\d_-]+\\!)?(\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6})(:\\$?[A-Z]{1,3}\\$?[1-9]{1}\\d{0,6})?$"

export function isValidExcelRange(range: string): boolean {
    const regex = new RegExp(excelRangeRegex);
    return regex.test(range);
}