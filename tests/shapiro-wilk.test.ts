import {ShapiroWilkW, ShapiroWilkResult} from "../src/shapiro-wilk";
import { normal } from "jstat-esm";

const alp = [
    91, 96, 65, 225, 69, 62, 86, 84, 68, 73, 77, 85, 105, 51, 66, 80, 35, 52, 135, 225, 41, 38, 70, 82, 52, 48, 90, 81, 58, 86, 89, 100,
    114, 113, 90, 95, 66, 57, 103, 113, 90, 135, 75, 35, 114, 78, 66, 57, 105, 60, 85, 105, 67, 59, 112, 50, 69, 74, 71, 95, 106, 92,
    78, 146, 144, 82, 144, 170, 58, 84, 170, 84, 95, 82, 82, 67, 71, 77, 79, 181, 58, 86, 77, 258, 77, 49, 73, 68, 86, 65, 78, 95, 69,
    67, 56, 63, 94, 87, 35, 100
]

// test("Normal quantile", () => {
//     const mu = 0;
//     const sigma = 1;
//     const p = [0.025, 0.975];
//     for (let i = 0; i < p.length; i++) {
//         const res = normalQuantile(p[i], mu, sigma);
//         expect(res).toBeCloseTo(normal.inv(p[i], mu, sigma), 4);
//         //console.log(res);
//     }
// });

test("Shapiro-Wilk", () => {
    const res: ShapiroWilkResult = ShapiroWilkW(alp);
    expect(res.w).toBeCloseTo(0.811,3);
    expect(res.p).toBeCloseTo(0.0,3);
    //console.log(res);
});
