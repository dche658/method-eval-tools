# Proposition

## Title

AACB Method Evaluation Tools

## Summary

Excel add-in to assist with validation or verification of clinical laboratory methods.

## Description

This Excel add-in is targetted at clinical laboratory professionals, including scientists and pathologists. It delivers specialized statistical procedures used for verifying and validating clinical pathology methods. It provides essential calculations missing from native Excel functions, including Weighted Deming regression, Passing-Bablok regression, and variance component analysis. Additionally, the add-in features dedicated utilities to help users derive or verify clinical laboratory test reference intervals. The add-in has been developed by the Australsian Association for Clinical Biochemistry and Laboratory Medicine (AACB). A method verification template that makes optimal use of the add-in is available for download from the Add-in website https://metools.chesher.id.au/assets/mvw.xlsx.

# Discoverability

## Categories

I do not know what categories are available. May need to evaluate when the submission is made.

## Keywords

- clinical laboratory Excel add-in
- lab method verification software
- Excel statistical tools clinical laboratory

## Listing Organisation

Australsian Association for Clinical Biochemistry and Laboratory Medicine

# Graphic Elements

I do not know if you wish to include any for the market place.

## Logo 

Between 216x216 and 350x350 pixels (square)

# Acquisition, Pricing, and Terms

I have no idea about this. My intent was to offer it for free as I do not plan to offer any warranty.

# Offer Information

## Terms and Conditions

MIT License Copyright (c) 2026 Australasian Association for Clinical Biochemistry and Laboratory Medicine

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice (including the next paragraph) shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Privacy Policy

I assume the AACB has one!

## Documentation

https://metools.aacb.asn.au/help/

## Learn More

https://metools.aacb.asn.au/help/

# 1120 Office Add-ins: Word, Excel, PowerPoint, and Outlook

## 1120.1 Offer requirements

All Office Add-ins must use the latest version of the Microsoft-hosted Office.js file at (https://appsforoffice.microsoft.com/lib/1/hosted/office.js). CONFIRMED

All Office Add-ins must use the latest released (not preview) version of the manifest schema for either the unified manifest for Microsoft 365 or for the add-in only manifest. CONFIRMED

Specify a valid Support URL when using the add-in only manifest, or a valid website URL when using the unified manifest for Microsoft 365. CONFIRMED

A high-resolution icon is mandatory. CONFIRMED 128x128 pixel.

Source location in the add-in only manifest or the “extensions.runtimes,code” property in the unified manifest for Microsoft 365, must point to a valid web address. CONFIRMED

The version number in the app package updates must be incremented. CONFIRMED

## 1120.2 Mobile requirements

Not applicable. Not designed for mobile devices.

## 1120.3 Functionality

As far as I can test. This should all be okay. Although I have not been able to test on Apple devices. Nevertheless, I believe it should work.

## 1120.5 Excel custom functions

Your custom functions metadata must have the helpUrl property set. CONFIRMED

### 1120.5.2 Security

To help to ensure the security of your app and users, your custom functions HTML, JavaScript, and JSON metadata files must be hosted on the same domain. CONFIRMED

### 1120.5.3 Functionality

Add-ins that contain custom functions must support add-in commands. This is to ensure that users can easily discover your add-in. CONFIRMED

Your add-in must work across all platforms that support custom functions. 

After an add-in that uses the equivalent add-ins feature is approved, all future updates to the add-in must include the appropriate markup in the manifest. This markup ensures that your custom functions save in XLL-compatible mode.

### 1120.5.4 Validation

To help ensure an efficient validation process, if your add-in contains custom functions, you must provide certification test notes for at least one custom function to validate them on submission.

Custom Function METOOLS.BOXCOX(x, lambda)

For values of x equal to 10, 20, 100, and 2000. If lambda equals 0.5, the boxcox function should return the following values: 4.325, 6.944, 18.000, and 87.443 respectively.

If lambda equals zero, then the boxcox function should return 2.303, 2.996, 4.605, 7.601 respectively.

```
const data1 = [[0], [2], [3], [4]];
const data2 = [[0], [2], [3], [7]];

const x = [10, 20, 100, 2000];
const y0_5 = [4.325, 6.944, 18.000, 87.443];
const bcLambda0 = [2.303, 2.996, 4.605, 7.601];

test("Box-Cox lambda=0.5", () => {
    for (let i=0;i<x.length;i++) {
        expect(boxcox(x[i], 0.5)).toBeCloseTo(y0_5[i]);
    }
}); 

test("Box-Cox lambda=0", () => {
    for (let i=0;i<x.length;i++) {
        expect(boxcox(x[i], 0)).toBeCloseTo(Math.log(x[i]));
        expect(boxcox(x[i], 0)).toBeCloseTo(bcLambda0[i]);
    }
});

test("Box-Cox inverse returns the original value for lambda 0", () => {
    const x = 4;
    expect(boxcoxinv(boxcox(x, 0), 0)).toBeCloseTo(x);
});

test("Box-Cox inverse returns the original value for non-zero lambda", () => {
    const x = 4;
    expect(boxcoxinv(boxcox(x, 2), 2)).toBeCloseTo(x);
});

test("Box-Cox wrapper matches the reference implementation", () => {
    expect(boxcox(9, 0)).toBeCloseTo(boxCox(9, 0));
    expect(boxcox(9, 2)).toBeCloseTo(boxCox(9, 2));
});

test("Shapiro-Wilk wrapper flattens a 2D array and returns the same stats", () => {
    const input = [[1, 2, 3], [4, 5, 6]];
    const expected = ShapiroWilkW([1, 2, 3, 4, 5, 6]);
    const result = shapirowilkw(input);

    expect(result[0][0]).toBeCloseTo(expected.w ?? NaN);
    expect(result[1][0]).toBeCloseTo(expected.p ?? NaN);
});

test("Dixon Reed Test", () => {
    let res = dixonreed(data1);
    expect(res[1][1]).toBeCloseTo(1 / 4);
    expect(res[2][1]).toBe(undefined);
    res = dixonreed(data2);
    expect(res[1][1]).toBeCloseTo(4 / 7);
    expect(res[2][1]).toBe(7);
});
```