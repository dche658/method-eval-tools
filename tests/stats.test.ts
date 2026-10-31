import {test, expect} from "vitest";
import { density } from "../src/stats";
import { alp } from "./testdata";
import * as fs from "fs";

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

test("KDE Test with ALP data", () => {
    const sampleData = alp;
    const result = density(sampleData);
    expect(result.x.length).toBe(512);
    expect(result.y.length).toBe(512);
    expect(result.bw).toBeGreaterThan(0);
    console.log(`Calculated Bandwidth (bw): ${result.bw.toFixed(4)}`);
    //write result to csv file in local directory
    // const headers = ["x", "y"];
    // const csvRows: string[] =  [];
    // csvRows.push(headers.join(","));
    // for (let i = 0; i < result.x.length; i++) {
    //     csvRows.push(`${result.x[i]},${result.y[i]}`);
    // }
    // const csvContent = csvRows.join("\n");
    // fs.writeFileSync("output.csv", csvContent, "utf-8");
});