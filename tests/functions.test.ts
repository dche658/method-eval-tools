import {
    boxcox,
    boxcoxinv,
    dixonreed,
    shapirowilkw,
} from "../src/functions/functions";
import { ShapiroWilkResult, ShapiroWilkW } from "../src/shapiro-wilk";
import {test, expect} from "vitest";

const data1 = [[0], [2], [3], [4]];
const data2 = [[0], [2], [3], [7]];
const data3 = [[10],[90],[98],[99],[100]]
test("Dixon Reed Test", () => {
    let res = dixonreed(data1);
    expect(res[1][1]).toBeCloseTo(1/4);
    expect(res[2][1]).toBe(undefined);
    console.log(res);
    res = dixonreed(data2);
    expect(res[1][1]).toBeCloseTo(4/7);
    expect(res[2][1]).toBe(7);
    console.log(res);
});

test("Dixon Reed Test with low outlier", () => {
    const res = dixonreed(data1);
    expect(res[1][0]).toBeCloseTo(1/2);
    expect(res[2][0]).toBe(0);
    expect(res[2][1]).toBe(undefined);
    console.log(res);
});

test("Dixon Reed Test data 3 with low outlier", () => {
    const res = dixonreed(data3);
    expect(res[1][0]).toBeCloseTo(8/9);
    expect(res[2][0]).toBe(10);
    expect(res[2][1]).toBe(undefined);
    console.log(res);
});

test("BOXCOX custom function", () => {
    expect(boxcox(4, 2)).toBe(7.5);
    expect(boxcox(4, 0)).toBeCloseTo(Math.log(4));
});

test("BOXCOXINV custom function", () => {
    expect(boxcoxinv(7.5, 2)).toBeCloseTo(4);
    expect(boxcoxinv(Math.log(4), 0)).toBeCloseTo(4);
});

test("SHAPIROWILKW custom function returns W and p-value as a column", () => {
    const values = [[1, 2, 3], [4, 5, 6, 7, 8]];
    const result = shapirowilkw(values);
    const flattenedValues = values.reduce((allValues, row) => allValues.concat(row), [] as number[]);
    const expected: ShapiroWilkResult = ShapiroWilkW(flattenedValues);

    expect(result).toHaveLength(2);
    expect(result[0]).toHaveLength(1);
    expect(result[1]).toHaveLength(1);
    expect(result[0][0]).toBeCloseTo(expected.w!);
    expect(result[1][0]).toBeCloseTo(expected.p!);
});