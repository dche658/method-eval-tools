import {test, expect} from "vitest";
import { 
    bowley, log, lognorm,
    iboxplot, interpretation,
    adjust_digits, permissible_uncertainty,
    reflim, round,
    truncated_qqplot
} from "../src/reflim";

import { alp } from "./testdata"

const truncated = [
72 ,49 ,71 ,85 ,61 ,38 ,45 ,43 ,76 ,64 ,66 ,58 ,69 ,78 ,43 ,57 ,48 ,54 ,43 ,47 ,
64 ,66 ,35 ,77 ,82 ,49 ,42 ,48 ,54 ,66 ,49 ,63 ,64 ,48 ,74 ,46 ,81 ,60 ,71 ,50 ,
63 ,71 ,61 ,42 ,47 ,43 ,51 ,61 ,38 ,74 ,83 ,55 ,58 ,43 ,41 ,62 ,46 ,53 ,58 ,59 ,
53 ,55 ,81 ,71 ,63 ,60 ,61 ,81 ,48 ,78 ,75 ,81 ,48 ,65 ,63 ,57 ,44 ,66 ,57 ,56 ,
39 ,42 ,63 ,64 ,52 ,49 ,66 ,45 ,50 ,37 ,38 ,47 ,41 ,77 ,80 ,39 ,86 ,67 ,58 ,74 ,
37 ,50 ,50 ,51 ,52 ,41 ,64 ,40 ,45 ,47 ,70 ,68 ,86 ,51 ,38 ,49 ,39 ,55 ,62 ,36 ,
55 ,66 ,60 ,64 ,72 ,38 ,48 ,44 ,65 ,55 ,67 ,69 ,42 ,55 ,63 ,77 ,56 ,35 ,73 ,53 ,
70 ,59 ,42 ,36 ,60 ,60 ,58 ,56 ,48 ,73 ,79 ,47 ,47 ,77 ,43 ,57 ,45 ,49 ,44 ,48 ,
66 ,42 ,44 ,45 ,58 ,74 ,36 ,62 ,66 ,51 ,53 ,65 ,40 ,66 ,62 ,44 ,42 ,62 ,51 ,55 ,
40 ,41 ,50 ,50 ,54 ,78 ,54 ,61 ,45 ,58 ,51 ,51 ,82 ,55 ,44 ,62 ,59 ,43 ,79 ,50 ,
36 ,68 ,58 ,75 ,63 ,56 ,46 ,42 ,71 ,51 ,51 ,40 ,39 ,49 ,83 ,64 ,45 ,56 ,64 ,73 ,
48 ,47 ,74 ,50 ,55 ,49 ,56 ,62 ,81 ,52 ,40 ,62 ,48 ,55 ,55 ,45 ,62 ,36 ,37 ,56 ,
55 ,42 ,61 ,39 ,53 ,84 ,61 ,75 ,42]

test("Bowley", () => {
    let bs = bowley(alp);
    expect(bs).toEqual(0)
});

test("Bowley Difference", () => {
    let bs = bowley(alp)
    let bs_log = bowley(log(alp));
    let diff = bs - bs_log
    expect(diff).toBeGreaterThan(0.05)
});

test("Lognorm", () => {
    const res = lognorm(alp)
    expect(res.lognormal).toBeTruthy()
    expect(res.BowleySkewness.normal).toEqual(0)
    expect(res.BowleySkewness.delta).toBeGreaterThanOrEqual(0.05)
})

test("adjust_digits", () => {
    const res = adjust_digits(45.6789);
    //console.log(res.digits)
    expect(res.digits).toEqual(1);
    expect(res.x_round).toEqual(45.7);
});

test("round", () => {
    const val = 45.67389;
    const r1 = round(val,1);
    const r2 = round(val,2);
    const r3 = round(val,3);
    expect(r1).toEqual(45.7);
    expect(r2).toEqual(45.67);
    expect(r3).toEqual(45.674); 
});

test("iBoxPlot", () => {
    const ibp = iboxplot(alp, null, 2.5, true)
    expect(ibp["truncation.points"].lower).toEqual(35.0);
    expect(ibp["truncation.points"].upper).toEqual(86.0);
    //console.log(ibp)
});

test("truncated_qqplot", () => {
    const res = truncated_qqplot(truncated);
    expect(res.mean).toBeCloseTo(4.006914, 6);
    expect(res.sd).toBeCloseTo(0.263591, 6);
    //console.log(res)
});

test("permissible_uncertainty", () => {
    const pu = permissible_uncertainty(32.8, 92.2, true);
    expect(pu.lower_lim_low).toBeCloseTo(30.3);
    expect(pu.lower_lim_upp).toBeCloseTo(35.3);
    expect(pu.upper_lim_low).toBeCloseTo(86.6);
    expect(pu.upper_lim_upp).toBeCloseTo(97.8);
    //console.log(pu)
});

test("interpretation", ()=>{
    const ip = interpretation([32.8, 92.2], [30, 110])
    //console.log(ip)
});

test("reflim", () => {
    const res = reflim(alp, true, [30,110])
    expect(res.limits.lower_lim).toBeCloseTo(32.8);
    expect(res.limits.upper_lim).toBeCloseTo(92.2);
    console.log(res);
});

test("Miscellaneous", () => {
    const arr = [23, 45];
    //console.log(Array.isArray(arr) && arr.length==2);
});