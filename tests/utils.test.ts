import { isValidExcelRange } from "../src/utils";
test('isValidExcelRange valid ranges', () => {
    const validRanges = [
        "A1",
        "A1:B2",
        "$A$1",
        "$A$1:$B$2",
        "AA10:ZZ100",
        "$AA$10:$ZZ$100",
        "XFD1048576",
        "$XFD$1048576",
        "Sheet1!B2:B12",
        "Sheet1!A1",
        "Sheet_2!$C$3:$D$4"
    ];
    validRanges.forEach(range => {
        expect(isValidExcelRange(range)).toBe(true);
    });
});
test('isValidExcelRange invalid ranges', () => {
    const invalidRanges = [
        "A",
        "1",
        "A1:B",
        "A1:2",
        "A0",
        "0.1",
        "A1:B0",
        "Sheet1"
    ];
    invalidRanges.forEach(range => {
        expect(isValidExcelRange(range)).toBe(false);
    });
});