import {
    JackknifeConfidenceInterval,
    BootstrapConfidenceInterval,
    PassingBablokRegression,
    DemingRegression,
    WeightedDemingRegression,
    DEFAULT_ERROR_RATIO
} from "../src/regression"

const x = [10377.5, 4056, 2654, 4747, 1459.5, 5880, 3871, 2461, 1802, 1607.5,
    4329, 7911.5, 1798.5, 6504, 9506.5, 4781, 2122, 17859, 7064, 963.5,
    408.5, 5623, 4923.5, 2745, 1738.5, 7057, 3001.5, 895.5, 12155.5, 1290.5,
    943, 11312.5, 2700, 3561, 1711, 1006, 1180.5, 9007, 4438, 3606];

const y = [10741.5, 3848.5, 2655.5, 5190.5, 1554.5, 6025, 3861.5, 2533.5, 1858, 1636, 
    4341.5, 8185, 1753.5, 6793, 9056, 4724.5, 2133.5, 16576.5, 7106.5, 986.5, 
    481, 5603, 5046, 2799.5, 1721.5, 7547.5, 3009.5, 986.5, 11712, 1416, 
    998.5, 11698.5, 2829, 3776.5, 1899, 1042, 1616, 9380.5, 4332.5, 3308.5];

test("Passing Bablok Regression", () => {
    let regression = new PassingBablokRegression();
    let reg = regression.calculate(x, y);
    //console.log(reg);
    // expected output from R mcr package version 1.3.3.1
    expect(reg.slope).toBeCloseTo(0.9987917, 4);
    expect(reg.intercept).toBeCloseTo(57.2280851, 3);
    expect(reg.slopeLCL).toBeCloseTo(0.9710559, 4);
    expect(reg.slopeUCL).toBeCloseTo(1.027696, 4);
    expect(reg.interceptLCL).toBeCloseTo(-8.774619, 3);
    expect(reg.interceptUCL).toBeCloseTo(125.435582, 3);
});

test("Weighted Deming Regression with jackknife confidence interval", () => {
    let regression = new WeightedDemingRegression();
    let reg = regression.calculate(x, y);
    let ci = new JackknifeConfidenceInterval(x, y, regression);
    let ciRes = ci.calculate();
    //console.log(ciRes);
    // expected output from R mcr package version 1.3.3.1
    expect(ciRes.slope).toBeCloseTo(0.9949122, 4); // output from R mcr package
    expect(ciRes.intercept).toBeCloseTo(80.3878188, 3);
    expect(ciRes.slopeLCL).toBeCloseTo(0.9751623, 4);
    expect(ciRes.slopeUCL).toBeCloseTo(1.014662, 4);
    expect(ciRes.interceptLCL).toBeCloseTo(44.8202707, 3);
    expect(ciRes.interceptUCL).toBeCloseTo(115.955367, 3);
    expect(ciRes.slopeSE).toBeCloseTo(0.009755964, 5);
    expect(ciRes.interceptSE).toBeCloseTo(17.569477709, 3);
});
