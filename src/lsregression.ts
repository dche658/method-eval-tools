/**
 * Simple Least Squares Regression
 * 
 * Included for completeness but not currently included in the add-in
 */
import { Regression, ConfidenceIntervalModel } from "./regression";
import { studentt } from "jstat-esm";

class LeastSquaresRegression implements Regression {
  calculate(x: number[], y: number[]) {
    let sum_xy = 0;
    let sum_xx = 0;
    let sum_x = 0;
    let sum_y = 0;
    let n = x.length;
    for (let i = 0; i < n; i++) {
      sum_xy += x[i] * y[i];
      sum_xx += x[i] * x[i];
      sum_x += x[i];
      sum_y += y[i];
    }
    const ss_xy = sum_xy - (sum_x * sum_y) / n;
    const ss_xx = sum_xx - (sum_x * sum_x) / n;
    let b1 = ss_xy / ss_xx;
    let b0 = (sum_y - b1 * sum_x) / n;

    return {
      slope: b1,
      intercept: b0,
      slopeLCL: NaN,
      slopeUCL: NaN,
      interceptLCL: NaN,
      interceptUCL: NaN,
    };
  }
}

class LeastSquaresConfidenceInterval {
  alpha: number; //alpha for upper and lower confidence limits

  constructor(alpha = 0.05) {
    this.alpha = alpha;
  }

  calculate(x: number[], y: number[]): ConfidenceIntervalModel {
    const regression = new LeastSquaresRegression();
    const reg = regression.calculate(x, y);
    const n = x.length;
    let sum_xx = 0;
    let sum_xy = 0;
    let sum_yy = 0;
    let sum_y = 0;
    let sum_x = 0;

    for (let i = 0; i < n; i++) {
      sum_xx += x[i] * x[i];
      sum_xy += x[i] * y[i];
      sum_yy += y[i] * y[i];
      sum_y += y[i];
      sum_x += x[i];
      //sse += (y[i] - reg.intercept - reg.slope * x[i]) * (y[i] - reg.intercept - reg.slope * x[i]);
    }
    //sum of squares
    const ss_xx = sum_xx - (sum_x * sum_x) / n;
    const ss_xy = sum_xy - (sum_x * sum_y) / n;
    const ss_yy = sum_yy - (sum_y * sum_y) / n;

    //Calculate SSE to reduce rounding errors
    //Mendenhall WM, Sincich TL. 2016. Statistics for Engineering
    //and the Sciences 6th ed. CRC Press, Boca Raton. page 503.
    const sse = ss_yy - ss_xy * reg.slope;
    const s = Math.sqrt(sse / (n - 2));
    const s_slope = s / Math.sqrt(ss_xx);
    const z = studentt.inv(1 - this.alpha / 2, n - 2);

    //Mendenhall WM, Sincich TL. 2016. Statistics for Engineering
    //and the Sciences 6th ed. CRC Press, Boca Raton. page 503.
    const s_intercept = s * Math.sqrt(1 / n + Math.pow(sum_x / n, 2) / ss_xx);

    // console.log(`ss_xx = ${ss_xx}, ss_xy = ${ss_xy}, ss_yy = ${ss_yy}`);
    // console.log(`sse = ${sse}; s = ${s}`);
    // console.log(`s_slope = ${s_slope}, z=${z}`);

    const slope_lcl = reg.slope - z * s_slope;
    const slope_ucl = reg.slope + z * s_slope;
    const intercept_lcl = reg.intercept - z * s_intercept;
    const intercept_ucl = reg.intercept + z * s_intercept;

    return {
      slope: reg.slope,
      intercept: reg.intercept,
      slopeSE: s_slope,
      interceptSE: s_intercept,
      slopeLCL: slope_lcl,
      slopeUCL: slope_ucl,
      interceptLCL: intercept_lcl,
      interceptUCL: intercept_ucl,
    };
  }
}

export { LeastSquaresRegression, LeastSquaresConfidenceInterval };
