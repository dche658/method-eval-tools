import {test, expect} from "vitest";
import { density } from "../src/stats";

const sampleData = [1, 1.5, 2, 2.1, 2.5, 3, 5, 5.5, 6];
const result = density(sampleData);

test("Kernel Density Estimation (KDE) Test", () => {
    expect(result.x.length).toBe(512);
    expect(result.y.length).toBe(512);
    expect(result.bw).toBeGreaterThan(0);

    console.log(`Calculated Bandwidth (bw): ${result.bw.toFixed(4)}`);
    console.log(`First point -> x: ${result.x[0].toFixed(2)}, y: ${result.y[0].toExponential(4)}`);
    console.log(`Peak point  -> x: ${result.x[200].toFixed(2)}, y: ${result.y[200].toFixed(4)}`);
});