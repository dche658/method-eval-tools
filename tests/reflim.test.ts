import {test, expect} from "vitest";
import { 
    bowley, log, lognorm,
    iboxplot, interpretation,
    adjust_digits, permissible_uncertainty,
    range, reflim, round, seq,
    truncated_qqplot
} from "../src/reflim";

const alp = [
    72 ,49 ,98 ,71 ,85 ,61 ,92 ,38 ,45 ,43 ,112 ,76 ,64 ,66 ,58 ,69 ,78 ,29 ,43 ,
    57 ,102 ,48 ,54 ,43 ,47 ,34 ,64 ,66 ,35 ,77 ,34 ,82 ,49 ,42 ,48 ,54 ,66 ,49 ,
    63 ,64 ,48 ,114 ,74 ,46 ,81 ,60 ,34 ,71 ,50 ,63 ,71 ,61 ,42 ,32 ,47 ,31 ,43 ,
    51 ,93 ,33 ,61 ,38 ,74 ,83 ,55 ,58 ,43 ,93 ,24 ,41 ,62 ,46 ,31 ,53 ,58 ,59 ,
    53 ,55 ,81 ,71 ,63 ,60 ,61 ,81 ,48 ,78 ,75 ,81 ,48 ,65 ,63 ,57 ,44 ,66 ,57 ,
    56 ,104 ,32 ,39 ,42 ,63 ,64 ,31 ,33 ,52 ,49 ,66 ,45 ,50 ,30 ,37 ,38 ,47 ,118 ,
    41 ,77 ,95 ,80 ,39 ,86 ,67 ,31 ,58 ,89 ,74 ,37 ,50 ,50 ,51 ,52 ,41 ,34 ,64 ,
    40 ,45 ,47 ,33 ,112 ,70 ,68 ,86 ,51 ,38 ,49 ,39 ,55 ,62 ,36 ,55 ,66 ,60 ,64 ,
    34 ,29 ,72 ,38 ,48 ,44 ,65 ,55 ,67 ,69 ,42 ,55 ,63 ,77 ,56 ,35 ,73 ,53 ,70 ,
    59 ,42 ,36 ,60 ,60 ,58 ,97 ,56 ,48 ,73 ,79 ,47 ,95 ,47 ,33 ,77 ,43 ,57 ,45 ,
    49 ,44 ,48 ,66 ,42 ,111 ,44 ,32 ,182 ,45 ,58 ,74 ,34 ,36 ,62 ,66 ,51 ,53 ,65 ,
    40 ,66 ,62 ,44 ,42 ,62 ,112 ,51 ,94 ,55 ,92 ,40 ,41 ,50 ,50 ,54 ,78 ,54 ,61 ,
    45 ,98 ,58 ,148 ,26 ,34 ,51 ,51 ,82 ,55 ,44 ,62 ,59 ,43 ,79 ,50 ,36 ,68 ,58 ,
    75 ,63 ,56 ,46 ,42 ,71 ,26 ,51 ,51 ,40 ,39 ,31 ,49 ,83 ,64 ,45 ,119 ,56 ,64 ,
    73 ,48 ,29 ,47 ,74 ,50 ,55 ,49 ,56 ,62 ,81 ,52 ,40 ,62 ,48 ,97 ,55 ,55 ,45 ,
    62 ,36 ,37 ,56 ,55 ,94 ,42 ,61 ,39 ,53 ,84 ,28 ,61 ,75 ,42
];

const age = [
    33 ,45 ,59 ,53 ,29 ,19 ,40 ,45 ,18 ,29 ,62 ,27 ,68 ,49 ,26 ,64 ,40 ,30 ,52 ,
    24 ,60 ,46 ,50 ,52 ,40 ,46 ,64 ,37 ,76 ,37 ,43 ,53 ,67 ,50 ,30 ,73 ,33 ,74 ,
    36 ,70 ,67 ,38 ,69 ,26 ,42 ,55 ,48 ,63 ,70 ,20 ,65 ,27 ,35 ,18 ,71 ,33 ,55 ,
    51 ,38 ,37 ,43 ,57 ,46 ,60 ,30 ,31 ,42 ,63 ,76 ,58 ,27 ,27 ,78 ,19 ,67 ,34 ,
    28 ,52 ,43 ,27 ,49 ,32 ,26 ,43 ,42 ,79 ,26 ,65 ,60 ,29 ,55 ,60 ,31 ,63 ,49 ,
    45 ,68 ,74 ,30 ,78 ,38 ,25 ,54 ,78 ,51 ,75 ,77 ,30 ,57 ,76 ,31 ,56 ,27 ,50 ,
    43 ,59 ,32 ,29 ,61 ,34 ,35 ,33 ,65 ,19 ,60 ,36 ,65 ,37 ,30 ,64 ,25 ,34 ,18 ,
    72 ,73 ,23 ,41 ,42 ,59 ,57 ,51 ,30 ,55 ,47 ,73 ,36 ,18 ,35 ,76 ,22 ,57 ,43 ,
    62 ,62 ,41 ,68 ,75 ,56 ,63 ,35 ,42 ,34 ,20 ,39 ,47 ,54 ,36 ,27 ,71 ,49 ,31 ,
    49 ,40 ,76 ,79 ,61 ,79 ,70 ,42 ,23 ,75 ,42 ,26 ,37 ,61 ,78 ,23 ,55 ,60 ,33 ,
    21 ,43 ,30 ,37 ,20 ,78 ,18 ,38 ,20 ,44 ,76 ,31 ,64 ,47 ,56 ,69 ,78 ,31 ,45 ,
    69 ,19 ,26 ,20 ,45 ,37 ,29 ,42 ,18 ,59 ,73 ,29 ,70 ,31 ,55 ,44 ,61 ,52 ,45 ,
    51 ,33 ,72 ,72 ,42 ,44 ,49 ,55 ,56 ,21 ,20 ,27 ,25 ,59 ,51 ,29 ,70 ,27 ,27 ,
    34 ,26 ,42 ,70 ,71 ,63 ,27 ,30 ,27 ,46 ,70 ,76 ,36 ,61 ,50 ,76 ,29 ,64 ,31 ,
    61 ,27 ,34 ,32 ,58 ,45 ,20 ,37 ,76 ,76 ,19 ,65 ,20 ,39 ,55 ,79 ,52 ,61 ,25 ,
    63 ,23 ,55 ,45 ,50 ,45 ,45 ,65 ,24 ,66 ,22 ,58 ,32 ,76 ,24
]

const sex = [
    "M" ,"M" ,"F" ,"M" ,"F" ,"F" ,"F" ,"F" ,"M" ,"M" ,"F" ,"F" ,"F" ,"M" ,"M" ,
    "M" ,"F" ,"M" ,"F" ,"F" ,"M" ,"F" ,"F" ,"M" ,"F" ,"F" ,"M" ,"F" ,"F" ,"F" ,
    "F" ,"F" ,"F" ,"M" ,"M" ,"F" ,"F" ,"M" ,"M" ,"M" ,"M" ,"F" ,"M" ,"M" ,"F" ,
    "F" ,"F" ,"F" ,"F" ,"F" ,"M" ,"M" ,"M" ,"F" ,"F" ,"F" ,"F" ,"F" ,"M" ,"M" ,
    "F" ,"M" ,"F" ,"F" ,"F" ,"M" ,"M" ,"M" ,"M" ,"F" ,"F" ,"F" ,"F" ,"M" ,"M" ,
    "M" ,"F" ,"M" ,"M" ,"M" ,"F" ,"M" ,"F" ,"F" ,"F" ,"M" ,"M" ,"F" ,"F" ,"M" ,
    "F" ,"M" ,"M" ,"F" ,"F" ,"M" ,"M" ,"M" ,"M" ,"F" ,"M" ,"M" ,"F" ,"F" ,"F" ,
    "F" ,"F" ,"M" ,"M" ,"F" ,"M" ,"F" ,"M" ,"M" ,"F" ,"M" ,"F" ,"F" ,"M" ,"M" ,
    "F" ,"M" ,"M" ,"M" ,"F" ,"F" ,"M" ,"F" ,"F" ,"F" ,"M" ,"F" ,"M" ,"M" ,"M" ,
    "M" ,"M" ,"F" ,"F" ,"M" ,"F" ,"F" ,"F" ,"F" ,"F" ,"M" ,"F" ,"M" ,"F" ,"M" ,
    "F" ,"F" ,"M" ,"M" ,"F" ,"F" ,"F" ,"F" ,"M" ,"M" ,"M" ,"F" ,"F" ,"F" ,"M" ,
    "M" ,"M" ,"F" ,"F" ,"M" ,"F" ,"M" ,"F" ,"F" ,"M" ,"M" ,"F" ,"F" ,"F" ,"F" ,
    "F" ,"M" ,"M" ,"M" ,"M" ,"M" ,"M" ,"M" ,"F" ,"F" ,"M" ,"F" ,"F" ,"M" ,"F" ,
    "F" ,"M" ,"F" ,"M" ,"M" ,"M" ,"F" ,"F" ,"F" ,"M" ,"F" ,"M" ,"F" ,"M" ,"F" ,
    "F" ,"F" ,"F" ,"F" ,"F" ,"M" ,"M" ,"M" ,"F" ,"F" ,"F" ,"M" ,"F" ,"M" ,"F" ,
    "F" ,"F" ,"F" ,"F" ,"F" ,"M" ,"F" ,"F" ,"M" ,"F" ,"M" ,"F" ,"M" ,"M" ,"F" ,
    "M" ,"M" ,"F" ,"M" ,"F" ,"M" ,"F" ,"M" ,"F" ,"M" ,"M" ,"M" ,"F" ,"M" ,"M" ,
    "M" ,"M" ,"F" ,"M" ,"M" ,"F" ,"M" ,"F" ,"M" ,"F" ,"F" ,"F" ,"F" ,"F" ,"F" ,
    "F" ,"M" ,"F" ,"F" ,"F" ,"F" ,"M" ,"M" ,"M" ,"F" ,"F" ,"F" ,"F" ,"F" ,"F" ,
    "M" ,"F" ,"M" ,"M" ,"M" ,"M" ,"M" ,"F" ,"M" ,"M" ,"M" ,"F" ,"F" ,"F" ,"M"
]

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

test("seq", () => {
    const arr = seq(1, 10, 0.1);
    expect(arr.length).toEqual(91);
    expect(arr[0]).toEqual(1);
    expect(arr[90]).toEqual(10);
    expect(arr[1]).toEqual(1.1);
    expect(arr[89]).toEqual(9.9);
    //console.log(arr)
});

test("range", () => {
    const arr = range(1,10,90);
    expect(arr.length).toEqual(90);
    expect(arr[0]).toEqual(1);
    expect(arr[89]).toEqual(10);
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