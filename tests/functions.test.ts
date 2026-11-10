import { boxcox, boxcoxinv, dixonreed, shapirowilkw } from "../src/functions/functions";
import { boxCox } from "../src/reference_intervals";
import { ShapiroWilkW } from "../src/shapiro-wilk";
import { test, expect } from "vitest";

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

