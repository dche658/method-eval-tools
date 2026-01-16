import { LeastSquaresRegression, LeastSquaresConfidenceInterval } from "../src/lsregression";
import { expect, test } from '@jest/globals';
import { BootstrapConfidenceInterval } from "../src/regression";

//Mendenhall WM, Sincich TL. 2016. Statistics for Engineering
//and the Sciences 6th ed. CRC Press, Boca Raton. Table 10.1 p 486
const x = [1, 2, 3, 4, 5];
const y = [1, 1, 2, 2, 4];

test("LeastSquaresRegression", ()=>{
    const regression = new LeastSquaresRegression();
    const reg = regression.calculate(x, y);
    //console.log(reg);
    expect(reg.slope).toBeCloseTo(0.7);
    expect(reg.intercept).toBeCloseTo(-0.1);
});

//Consistent with Figure 10.10 in Mendenhall 2016, showing output
//for SAS
test("LeastSquaresConfidenceInterval", ()=>{
    const ci = new LeastSquaresConfidenceInterval();
    const ciRes = ci.calculate(x, y);
    //console.log(ciRes);
    expect(ciRes.slope).toBeCloseTo(0.7);
    expect(ciRes.intercept).toBeCloseTo(-0.1);
    expect(ciRes.slopeSE).toBeCloseTo(0.191485);
    expect(ciRes.interceptSE).toBeCloseTo(0.635085);
    expect(ciRes.slopeLCL).toBeCloseTo(0.0906079);
    expect(ciRes.slopeUCL).toBeCloseTo(1.30939207);
    expect(ciRes.interceptLCL).toBeCloseTo(-2.12112485);
    expect(ciRes.interceptUCL).toBeCloseTo(1.92112485);
});

test("LS BootstrapConfidenceInterval", ()=>{
    const regression = new LeastSquaresRegression();
    const ci = new BootstrapConfidenceInterval(x, y, regression);
    const ciRes = ci.calculate();
    console.log(ciRes);
    expect(ciRes.slope).toBeCloseTo(0.7);
    expect(ciRes.intercept).toBeCloseTo(-0.1);
});
