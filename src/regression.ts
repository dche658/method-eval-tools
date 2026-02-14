/**
 * Small library for performing regression analysis on data from method
 * comparison studies.
 *
 * It has routines for Deming, Weighted Deming, and Passing-Bablok regression.
 * The default confidence interval of Deming and Weighted Deming regression
 * uses a leave-one-out Jackknife procedure.
 * The default confidence interval for Passing-Bablock regression is a
 * non-parametric estimate.
 *
 * Acknowledgement: A lot of the procedures are based on the algorithms
 * in the mcr package for R by Sergej Potapov 2021
 *
 * @author Douglas Chesher
 *
 * Created: August 2025.
 */

import { normal, studentt, mean, stdev } from "jstat-esm";

const CI_METHOD_JACKKNIFE = "jackknife";
const CI_METHOD_BOOTSTRAP = "bootstrap";
const CI_METHOD_NONPARAMETRIC = "np";
const CI_METHOD_DEFAULT = "default";
const REG_METHOD_DEMING = "Deming";
const REG_METHOD_WDEMING = "WDeming";
const REG_METHOD_PABA = "PaBa";

const DEFAULT_ERROR_RATIO = 1;
const DEFAULT_ITER_MAX = 30;
const DEFAULT_THRESHOLD = 0.000001;
const DEFAULT_ALPHA = 0.05;
const DEFAULT_BOOTSTRAP_N = 10000;

const PI4 = Math.PI / 4;

/**
 * Sum of the sqaured deviations.
 *
 * Sum of (x_i - mean(x))^2
 * @param arr array of values
 * @returns sum of deviations from the mean squared
 */
function devsq(arr: number[]): number {
  const avg = mean(arr);
  let sum = 0;
  for (let i = 0; i < arr.length; i++) {
    sum += (arr[i] - avg) * (arr[i] - avg);
  }
  return sum;
} //devsq

/* Calculate the q th quantile for the given probabilities.
 *
 * arr must be an array of numeric values
 * probs is an array of probabilities ranging from 0 to 1
 */
function quantile(arr: number[], probs: number[]): number[] | undefined {
  if (!Array.isArray(arr) || arr.length === 0) {
    return undefined; // Handle empty or invalid input
  }

  for (let q of probs) {
    if (q < 0 || q > 1) {
      throw new RangeError(`Quantiles must be between 0 and 1. q=${q}`);
    }
  }

  const sortedArr = [...arr].sort((a, b) => a - b); // Create a sorted copy
  const n = sortedArr.length;
  const quantiles = new Array<number>(probs.length);

  let getQuantile = function (sortedArr: number[], n: number, p: number): number {
    if (p === 1) {
      return sortedArr[n - 1]; // Maximum value
    }
    if (p === 0) {
      return sortedArr[0]; // Minimum value
    }

    const index = (n - 1) * p;
    const lowerIndex = Math.floor(index);
    const upperIndex = Math.ceil(index);

    // Interpolate the value if the index is not an integer.
    if (lowerIndex === upperIndex) {
      return sortedArr[lowerIndex]; // Exact match
    } else {
      // Linear interpolation
      const lowerValue = sortedArr[lowerIndex];
      const upperValue = sortedArr[upperIndex];
      return lowerValue + (upperValue - lowerValue) * (index - lowerIndex);
    }
  };

  for (let i = 0; i < probs.length; i++) {
    quantiles[i] = getQuantile(sortedArr, n, probs[i]);
  }

  return quantiles;
} //quantile

/* Object returned by calculate method of a Regression instance */
interface RegressionModel {
  slope: number;
  intercept: number;
  slopeLCL: number;
  slopeUCL: number;
  interceptLCL: number;
  interceptUCL: number;
}

/* Base class for performing regression analysis */
interface Regression {
  calculate(x: number[], y: number[]): RegressionModel;
}

/**
 * Ordinary Deming Regression
 * Assumes a constant SD for the measuring range
 * Error ratio is the SD for comparison method to SD of
 * the reference method.
 */
class DemingRegression implements Regression {
  private errorRatio: number;

  constructor(errorRatio = DEFAULT_ERROR_RATIO) {
    this.errorRatio = errorRatio;
  }

  /** Calculate regression
   * @param x must be an array of numeric values
   * @param y must be an array of numeric values
   */
  calculateDeming(x: number[], y: number[]): RegressionModel {
    // Simple validation
    if (x.length !== y.length || x.length === 0) {
      throw new Error("Input arrays must have the same non-zero length.");
    }
    if (!(typeof this.errorRatio === "number" && this.errorRatio > 0)) {
      throw new Error("Error ratio must be a positive number.");
    }
    if (Math.min(...x) < 0 || Math.min(...y) < 0) {
      throw new Error("Input arrays must contain only non-negative numbers.");
    }
    const n = x.length;
    let lambda = this.errorRatio;
    let mean_x = mean(x);
    let mean_y = mean(y);
    let u = devsq(x);
    let q = devsq(y);

    let p = 0;
    for (let i = 0; i < n; i++) {
      p = p + (x[i] - mean_x) * (y[i] - mean_y);
    }

    let b1 =
      (lambda * q - u + Math.sqrt(Math.pow(u - lambda * q, 2) + 4 * lambda * Math.pow(p, 2))) /
      (2 * lambda * p);
    let b0 = mean_y - b1 * mean_x;
    return {
      slope: b1,
      intercept: b0,
      slopeLCL: NaN,
      slopeUCL: NaN,
      interceptLCL: NaN,
      interceptUCL: NaN,
    };
  }

  calculate(
    x: number[],
    y: number[]
  ): {
    slope: number;
    intercept: number;
    slopeLCL: number;
    slopeUCL: number;
    interceptLCL: number;
    interceptUCL: number;
  } {
    return this.calculateDeming(x, y);
  }
} //Deming

/* Weighted Deming Regression
 *
 * Assumes constant CV across the measuring range of the method
 */
class WeightedDemingRegression implements Regression {
  private errorRatio: number;
  private iterMax: number;
  private threshold: number;

  /*
   * errorRatio ratio between squared measurement errors of reference- and test method, necessary for Deming regression (Default is 1).
   * iterMax maximal number of iterations.
   * threshold threshold value.
   */
  constructor(
    errorRatio = DEFAULT_ERROR_RATIO,
    iterMax = DEFAULT_ITER_MAX,
    threshold = DEFAULT_THRESHOLD
  ) {
    this.errorRatio = errorRatio;
    this.iterMax = iterMax;
    this.threshold = threshold;
  }

  /*
   * x measurement values of reference method as an Array.
   * Y measurement values of test method as an Array.
   */
  calculateWeightedDeming(x: number[], y: number[]): RegressionModel {
    if (x.length !== y.length || x.length === 0) {
      throw new Error("Input arrays must have the same non-zero length.");
    }
    if (!(typeof this.errorRatio === "number" && this.errorRatio > 0)) {
      throw new Error("Error ratio must be a positive number.");
    }
    if (!(typeof this.iterMax === "number" && this.iterMax > 0)) {
      throw new Error("Maximum number of iterations must be a positive number.");
    }
    if (!(typeof this.threshold === "number" && this.threshold > 0)) {
      throw new Error("Threshold must be a positive number.");
    }
    if (Math.min(...x) < 0 || Math.min(...y) < 0) {
      throw new Error("Input arrays must contain only non-negative numbers.");
    }

    /* Calculate starting values using a simple Deming regression */
    let demingRegression = new DemingRegression(this.errorRatio);
    let initial = demingRegression.calculate(x, y);
    let b1 = initial.slope;
    let b0 = initial.intercept;

    const n = x.length;
    let lambda = this.errorRatio;

    let d = new Array<number>(n);
    let x_hat = new Array<number>(n);
    let y_hat = new Array<number>(n);
    let w = new Array<number>(n);
    let w_x = new Array<number>(n);
    let w_y = new Array<number>(n);
    let w_u = new Array<number>(n);
    let w_q = new Array<number>(n);
    let w_p = new Array<number>(n);
    let iter = 0;

    // Iterate until convergence is reached or max interations is reached.
    while (iter <= this.iterMax) {
      iter += 1;
      if (iter == this.iterMax) console.log(`No convergence after ${this.iterMax} iterations.`);
      let sumW = 0;
      let wx = 0;
      let wy = 0;
      let wu = 0;
      let wq = 0;
      let wp = 0;

      /* Calculate the weighted sums */
      for (let i = 0; i < n; i++) {
        d[i] = y[i] - (b0 + b1 * x[i]);
        x_hat[i] = x[i] + (lambda * b1 * d[i]) / (1 + lambda * b1 * b1);
        y_hat[i] = y[i] - d[i] / (1 + lambda * Math.pow(b1, 2));
        w[i] = Math.pow((x_hat[i] + lambda * y_hat[i]) / (1 + lambda), -2);
        w_x[i] = w[i] * x[i];
        w_y[i] = w[i] * y[i];
        sumW += w[i];
        wx += w_x[i];
        wy += w_y[i];
      }
      wx = wx / sumW;
      wy = wy / sumW;

      //if (iter===1) console.log(`wx = ${wx}, wy = ${wy}, sumW = ${sumW}, d = ${d}, x_hat = ${x_hat}, y_hat = ${y_hat}, w = ${w}, w_x = ${w_x}, w_y = ${w_y}`);

      /* Calculate the regression parameters on the weighted values */
      for (let i = 0; i < n; i++) {
        w_u[i] = w[i] * Math.pow(x[i] - wx, 2);
        w_q[i] = w[i] * Math.pow(y[i] - wy, 2);
        w_p[i] = w[i] * (x[i] - wx) * (y[i] - wy);
        wu += w_u[i];
        wq += w_q[i];
        wp += w_p[i];
      }
      let b_1 =
        (lambda * wq -
          wu +
          Math.sqrt(Math.pow(wu - lambda * wq, 2) + 4 * lambda * Math.pow(wp, 2))) /
        (2 * lambda * wp);
      let b_0 = wy - b_1 * wx;

      // Check for convergence
      if (Math.abs(b1 - b_1) < this.threshold && Math.abs(b0 - b_0) < this.threshold) {
        //console.log(`sumW = ${sumW}, wx = ${wx}, wy = ${wy}, wu = ${wu}, wq = ${wq}, wp = ${wp}, b1 = ${b1}, b0 = ${b0}, iter = ${iter}`);
        //console.log(`w_x = ${w_x}, w_y = ${w_y}`);
        break;
      }
      b1 = b_1;
      b0 = b_0;
    }

    return {
      slope: b1,
      intercept: b0,
      slopeLCL: NaN,
      slopeUCL: NaN,
      interceptLCL: NaN,
      interceptUCL: NaN,
    };
  } // #calculate

  calculate(x: number[], y: number[]): RegressionModel {
    return this.calculateWeightedDeming(x, y);
  }
} //WeightedDeming

interface AngleMatrixModel {
  matrix: number[];
  nAllItems: number;
  nNeg: number;
  nNeg2: number;
  nPos: number;
  nPos2: number;
}

class PassingBablokRegression implements Regression {
  // Initially based on mcr package for R by Sergej Potapov 2021
  // but only implemented for positively correlated data and to always
  // calculate the non-parametric confidence intervals using the method of Passing and Bablock
  //
  // Passing H, Bablock W. A new biometrical procedure for testing the equality of measurements
  // from two different analytical methods. Applications of linear regression procedures for
  // method comparison studies in clinical chemistry, part I.
  // J Clin Chem Clin Biochem. 1983;21:709-720.
  private alpha: number;
  private positiveCorrelated: boolean;

  constructor(alpha = DEFAULT_ALPHA, positiveCorrelated = true) {
    this.alpha = alpha;
    this.positiveCorrelated = positiveCorrelated;
  }

  calculatePaBa(x: number[], y: number[]): RegressionModel {
    let slope = 0;
    let slopeL = 0;
    let slopeU = 0;
    let intercept = 0;
    let interceptL = 0;
    let interceptU = 0;
    let n = x.length;
    let angleMatrixObj = this.calcAngleMatrix(x, y);

    let matrix = angleMatrixObj.matrix;
    let nAllItems = angleMatrixObj.nAllItems;
    let nNeg = angleMatrixObj.nNeg;
    let nNeg2 = angleMatrixObj.nNeg2;
    let nPos = angleMatrixObj.nPos;
    let nPos2 = angleMatrixObj.nPos2;
    let offset = this.positiveCorrelated ? nNeg + nNeg2 : -1 * (nPos + nPos2);
    let nValIndex2 = nAllItems + offset;
    let lowestIdx = 0;
    let angle = 0;
    matrix.sort((a, b) => a - b); //compare function to ensure sort by numeric values, not lexicographically, which is the default

    let half = Math.floor((nValIndex2 + 1) / 2.0);
    if (nValIndex2 % 2 === 0) {
      slope = (matrix[half - 1] + matrix[half]) / 2.0;
    } else {
      slope = matrix[half - 1];
    }
    slope = Math.tan(slope);

    let z = normal.inv(1 - this.alpha / 2, 0, 1);
    let t = n * (n - 1) * (2 * n + 5);
    let dConf = Math.round(z * Math.sqrt(t / 18.0));
    let ml = Math.floor(nAllItems - dConf + offset);
    let mlIdx = Math.floor((ml + 1) / 2.0);
    if (this.positiveCorrelated) {
      lowestIdx = 2 * (nNeg - nNeg2) + 1;
    } else {
      lowestIdx = 2 * (nPos - nPos2) + 1;
    }
    if (mlIdx >= lowestIdx) {
      if (ml % 2 === 0) {
        angle = (matrix[mlIdx - 1] + matrix[mlIdx]) / 2.0;
      } else {
        angle = matrix[mlIdx - 1];
      }
      slopeL = Math.tan(angle);
    } else {
      slopeL = NaN;
    }
    let m2 = Math.floor(nAllItems + dConf + offset);
    let m2Idx = Math.floor((m2 + 1) / 2.0);
    if (m2Idx <= nAllItems) {
      if (m2 % 2 === 0) {
        angle = (matrix[m2Idx - 1] + matrix[m2Idx]) / 2.0;
      } else {
        angle = matrix[m2Idx - 1];
      }
      slopeU = Math.tan(angle);
    } else {
      slopeU = NaN;
    }
    let interceptArr = new Array(n);
    let interceptLArr = new Array(n);
    let interceptUArr = new Array(n);
    for (let i = 0; i < n; i++) {
      interceptArr[i] = y[i] - slope * x[i];
      interceptLArr[i] = y[i] - slopeU * x[i];
      interceptUArr[i] = y[i] - slopeL * x[i];
    }
    interceptArr.sort((a, b) => a - b);
    interceptLArr.sort((a, b) => a - b);
    interceptUArr.sort((a, b) => a - b);
    half = Math.floor(n / 2.0) - 1;

    //console.log(`Intercepts: ${interceptArr}; half=${half}`)

    if (n % 2 === 0) {
      intercept = (interceptArr[half] + interceptArr[half + 1]) / 2.0;
      interceptL = (interceptLArr[half] + interceptLArr[half + 1]) / 2.0;
      interceptU = (interceptUArr[half] + interceptUArr[half + 1]) / 2.0;
    } else {
      intercept = interceptArr[half];
      interceptL = interceptLArr[half];
      interceptU = interceptUArr[half];
    }
    return {
      slope: slope,
      intercept: intercept,
      slopeLCL: slopeL,
      slopeUCL: slopeU,
      interceptLCL: interceptL,
      interceptUCL: interceptU,
    };
  }

  calcAngleMatrix(x: number[], y: number[]): AngleMatrixModel {
    let nrows = x.length;
    let ncols = y.length;
    let nData = nrows * ncols;
    let matrix = new Array<number>(nData - 1);
    let nAllItems = 0;
    let nNeg = 0;
    let nNeg2 = 0;
    let nPos = 0;
    let nPos2 = 0;
    for (let j = 0; j < ncols; j++) {
      for (let k = 0; k < nrows; k++) {
        if (k < j) {
          //consider pairs of points. Reciprocal of pairs set to 500
          matrix[j * nrows + k] = 500;
        } else {
          //calculate differences
          let dx = this.calcDiff(x[k], x[j]);
          let dy = this.calcDiff(y[k], y[j]);
          if (dx !== 0) {
            //calculate slope if dx not zero
            nAllItems += 1;
            matrix[j * nrows + k] = Math.atan(dy / dx);
          } else if (dy !== 0) {
            //dx is zero and dy not zero
            nAllItems += 1;
            if (this.positiveCorrelated) {
              matrix[j * nrows + k] = Math.PI / 2; //second point directly above first
            } else {
              matrix[j * nrows + k] = (-1 * Math.PI) / 2; //second point directly below first
            }
          } else {
            //dx and dy are zero. Points are colocated.
            matrix[j * nrows + k] = 500;
          }
        }
        //
        if (matrix[j * nrows + k] < 500) {
          if (matrix[j * nrows + k] <= -1 * PI4) {
            nNeg += 1;
            if (matrix[j * nrows + k] < -1 * PI4) {
              nNeg2 += 1;
            }
          }
          if (matrix[j * nrows + k] >= PI4) {
            nPos += 1;
            if (matrix[j * nrows + k] > PI4) {
              nPos2 += 1;
            }
          }
        }
      } //next k
    } //next j
    return {
      matrix: matrix,
      nAllItems: nAllItems,
      nNeg: nNeg,
      nNeg2: nNeg2,
      nPos: nPos,
      nPos2: nPos2,
    };
  }

  calcDiff(a: number, b: number, eps = 0.000000000001): number {
    //Calculate difference between two numeric values that
    //gives exactly zero for very small relative differences.
    //Copied from mcr package for R by Sergej Potapov 2021
    let delta = a - b;
    if (Math.abs(delta) < eps * ((Math.abs(a) + Math.abs(b)) / 2)) {
      return 0;
    } else {
      return delta;
    }
  }

  calculate(x: number[], y: number[]): RegressionModel {
    return this.calculatePaBa(x, y);
  }
} //PassingBablokRegression

interface ConfidenceIntervalModel {
  slope: number;
  intercept: number;
  slopeSE: number;
  interceptSE: number;
  slopeLCL: number;
  slopeUCL: number;
  interceptLCL: number;
  interceptUCL: number;
}

/* Calculate a confidence interval using a jackknife procedure.
 *
 */
class JackknifeConfidenceInterval {
  private regression: Regression;
  private alpha: number;
  private x: number[];
  private y: number[];

  constructor(x: number[], y: number[], regression: Regression, alpha = DEFAULT_ALPHA) {
    this.x = x;
    this.y = y;
    this.regression = regression;
    this.alpha = alpha;
  }

  calculate(): ConfidenceIntervalModel {
    let n = this.x.length;
    let size = n - 1;
    if (size <= 1) throw new Error("Sample size must be greater than 2");
    let vx = new Array<number>(size);
    let vy = new Array<number>(size);
    let b_1 = new Array<number>(n);
    let b_0 = new Array<number>(n);

    let reg = this.regression.calculate(this.x, this.y);
    let b1 = reg.slope;
    let b0 = reg.intercept;

    //Iterate over the pool leaving one out
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < n; j++) {
        if (j < i) {
          vx[j] = this.x[j];
          vy[j] = this.y[j];
        } else if (j > i) {
          vx[j - 1] = this.x[j];
          vy[j - 1] = this.y[j];
        }
      }
      let regRes = this.regression.calculate(vx, vy);
      b_1[i] = regRes.slope;
      b_0[i] = regRes.intercept;
    }
    //console.log(`size vx = ${vx.length}, size b_1 = ${b_1.length}, b_1 = ${b_1}`);

    let se_b1 = this.linnetSE(b_1, b1);
    let se_b0 = this.linnetSE(b_0, b0);

    let z = studentt.inv(1 - this.alpha / 2, size - 1);

    return {
      slope: b1,
      intercept: b0,
      slopeSE: se_b1,
      interceptSE: se_b0,
      slopeLCL: b1 - z * se_b1,
      slopeUCL: b1 + z * se_b1,
      interceptLCL: b0 - z * se_b0,
      interceptUCL: b0 + z * se_b0,
    };
  }

  /*
   * Calculate the standard error using the procedure of Linnet
   */
  linnetSE(b_jack: number[], b_global: number): number {
    let n = b_jack.length; //number of data points
    let d_b = new Array<number>(n); //deviations for estimating the standard error
    for (let i = 0; i < n; i++) {
      d_b[i] = n * b_global - (n - 1) * b_jack[i];
    }
    //console.log(`d_b = ${d_b}`, `n = ${n}`);
    return stdev(d_b, true) / Math.sqrt(n);
  } //linnetSE
} //JackknifeConfidenceInterval

/* Calculate a confidence interval using a bootstrap procedure.
 *
 */
class BootstrapConfidenceInterval {
  private x: number[];
  private y: number[];
  private regression: Regression;
  private alpha: number;
  private bootstrapN: number;

  constructor(
    x: number[],
    y: number[],
    regression: Regression,
    bootstrapN = DEFAULT_BOOTSTRAP_N,
    alpha = DEFAULT_ALPHA
  ) {
    this.x = x;
    this.y = y;
    this.regression = regression;
    this.bootstrapN = bootstrapN;
    this.alpha = alpha;
  }
  calculate(): ConfidenceIntervalModel {
    const global_coefficents = this.regression.calculate(this.x, this.y);

    const coefficients = this.calculateBootstrapRegression(this.x, this.y, this.bootstrapN);

    const slopeCI = quantile(coefficients.b1, [this.alpha / 2, 1 - this.alpha / 2]);
    const interceptCI = quantile(coefficients.b0, [this.alpha / 2, 1 - this.alpha / 2]);

    return {
      slope: global_coefficents.slope,
      intercept: global_coefficents.intercept,
      slopeSE: NaN,
      interceptSE: NaN,
      slopeLCL: slopeCI === undefined ? NaN : slopeCI[0],
      slopeUCL: slopeCI === undefined ? NaN : slopeCI[1],
      interceptLCL: interceptCI === undefined ? NaN : interceptCI[0],
      interceptUCL: interceptCI === undefined ? NaN : interceptCI[1],
    };
  }

  //Calculate regression on each of the bootstrap samples
  //and return arrays containing the slopes (b1) and intercepts (b0)
  calculateBootstrapRegression(
    x: number[],
    y: number[],
    bootstrapN: number
  ): { b1: number[]; b0: number[] } {
    const b1 = new Array<number>(bootstrapN);
    const b0 = new Array<number>(bootstrapN);
    for (let i = 0; i < bootstrapN; i++) {
      let bs = this.getBootstrapSamples(x, y);
      let reg = this.regression.calculate(bs.x, bs.y);
      b1[i] = reg.slope;
      b0[i] = reg.intercept;
    }
    b1.sort((a, b) => a - b);
    b0.sort((a, b) => a - b);
    return {
      b1: b1,
      b0: b0,
    };
  }

  //Sample with replacement
  getBootstrapSamples(x: number[], y: number[]): { x: number[]; y: number[] } {
    const sx = new Array<number>(x.length);
    const sy = new Array<number>(y.length);
    for (let i = 0; i < x.length; i++) {
      let index = Math.floor(Math.random() * x.length);
      sx[i] = x[index];
      sy[i] = y[index];
    }
    return {
      x: sx,
      y: sy,
    };
  }
}

class MethodCompRegression implements Regression {
  private regressionMethod: string;
  private errorRatio: number;
  private iterMax: number;
  private threshold: number;
  private alpha: number;
  private ciMethod: string;
  private bootstrapN: number;

  constructor(
    regressionMethod: string,
    errorRatio: number = DEFAULT_ERROR_RATIO,
    iterMax: number = DEFAULT_ITER_MAX,
    threshold: number = DEFAULT_THRESHOLD,
    alpha: number = DEFAULT_ALPHA,
    ciMethod: string = CI_METHOD_DEFAULT,
    bootstrapN: number = DEFAULT_BOOTSTRAP_N
  ) {
    this.regressionMethod = regressionMethod;
    this.errorRatio = errorRatio;
    this.iterMax = iterMax;
    this.threshold = threshold;
    this.alpha = alpha;
    this.ciMethod = ciMethod;
    this.bootstrapN = bootstrapN;
  }

  calculate(x: number[], y: number[]): ConfidenceIntervalModel {
    let res = {
      slope: NaN,
      intercept: NaN,
      slopeLCL: NaN,
      slopeUCL: NaN,
      interceptLCL: NaN,
      interceptUCL: NaN,
      slopeSE: NaN,
      interceptSE: NaN,
    };
    if (this.regressionMethod === REG_METHOD_DEMING) {
      let regression = new DemingRegression(this.errorRatio);
      let reg = regression.calculate(x, y);
      if (this.ciMethod === CI_METHOD_JACKKNIFE || this.ciMethod === CI_METHOD_DEFAULT) {
        let ci = new JackknifeConfidenceInterval(x, y, regression, this.alpha);
        let ciRes = ci.calculate();
        res.slope = reg.slope;
        res.intercept = reg.intercept;
        res.slopeLCL = ciRes.slopeLCL;
        res.slopeUCL = ciRes.slopeUCL;
        res.interceptLCL = ciRes.interceptLCL;
        res.interceptUCL = ciRes.interceptUCL;
        res.slopeSE = ciRes.slopeSE;
        res.interceptSE = ciRes.interceptSE;
      } else if (this.ciMethod === CI_METHOD_BOOTSTRAP) {
        let ci = new BootstrapConfidenceInterval(x, y, regression, this.bootstrapN, this.alpha);
        let ciRes = ci.calculate();
        res.slope = reg.slope;
        res.intercept = reg.intercept;
        res.slopeLCL = ciRes.slopeLCL;
        res.slopeUCL = ciRes.slopeUCL;
        res.interceptLCL = ciRes.interceptLCL;
        res.interceptUCL = ciRes.interceptUCL;
      }
    } else if (this.regressionMethod === REG_METHOD_WDEMING) {
      let regression = new WeightedDemingRegression(this.errorRatio, this.iterMax, this.threshold);
      let reg = regression.calculate(x, y);
      if (this.ciMethod === CI_METHOD_JACKKNIFE || this.ciMethod === CI_METHOD_DEFAULT) {
        let ci = new JackknifeConfidenceInterval(x, y, regression, this.alpha);
        let ciRes = ci.calculate();
        res.slope = reg.slope;
        res.intercept = reg.intercept;
        res.slopeLCL = ciRes.slopeLCL;
        res.slopeUCL = ciRes.slopeUCL;
        res.interceptLCL = ciRes.interceptLCL;
        res.interceptUCL = ciRes.interceptUCL;
        res.slopeSE = ciRes.slopeSE;
        res.interceptSE = ciRes.interceptSE;
      } else if (this.ciMethod === CI_METHOD_BOOTSTRAP) {
        let ci = new BootstrapConfidenceInterval(x, y, regression, this.bootstrapN, this.alpha);
        let ciRes = ci.calculate();
        res.slope = reg.slope;
        res.intercept = reg.intercept;
        res.slopeLCL = ciRes.slopeLCL;
        res.slopeUCL = ciRes.slopeUCL;
        res.interceptLCL = ciRes.interceptLCL;
        res.interceptUCL = ciRes.interceptUCL;
      }
    } else if (this.regressionMethod === REG_METHOD_PABA) {
      let regression = new PassingBablokRegression(this.alpha);
      let reg = regression.calculate(x, y);
      if (this.ciMethod === CI_METHOD_DEFAULT) {
        res.slope = reg.slope;
        res.intercept = reg.intercept;
        res.slopeLCL = reg.slopeLCL;
        res.slopeUCL = reg.slopeUCL;
        res.interceptLCL = reg.interceptLCL;
        res.interceptUCL = reg.interceptUCL;
      } else if (this.ciMethod === CI_METHOD_BOOTSTRAP) {
        let ci = new BootstrapConfidenceInterval(x, y, regression, this.bootstrapN, this.alpha);
        let ciRes = ci.calculate();
        res.slope = reg.slope;
        res.intercept = reg.intercept;
        res.slopeLCL = ciRes.slopeLCL;
        res.slopeUCL = ciRes.slopeUCL;
        res.interceptLCL = ciRes.interceptLCL;
        res.interceptUCL = ciRes.interceptUCL;
      }
    } else {
      throw new Error(`Unknown regression method: ${this.regressionMethod}`);
    }
    return res;
  }
} // MethodCompRegression

export {
  CI_METHOD_BOOTSTRAP,
  CI_METHOD_JACKKNIFE,
  CI_METHOD_NONPARAMETRIC,
  CI_METHOD_DEFAULT,
  REG_METHOD_DEMING,
  REG_METHOD_WDEMING,
  REG_METHOD_PABA,
  DEFAULT_ALPHA,
  DEFAULT_ERROR_RATIO,
  DEFAULT_ITER_MAX,
  DEFAULT_THRESHOLD,
  DEFAULT_BOOTSTRAP_N,
  DemingRegression,
  WeightedDemingRegression,
  PassingBablokRegression,
  JackknifeConfidenceInterval,
  BootstrapConfidenceInterval,
  MethodCompRegression,
  RegressionModel,
  ConfidenceIntervalModel,
  Regression,
  quantile,
};
