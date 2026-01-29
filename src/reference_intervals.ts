/**
* Robust algorithm for calculating reference intervals for
* small data sets.
* 
* Based on the R implementation in the referenceIntervals 
* package by Daniel Finnegan.
* 
* Horn PS, Pesce AJ, Copeland BE. A robust approach to 
* reference interval estimation and evaluation. Clinical Chemistry.
* 1998. 44(3):622-631.
* 
* This method is applicable to symmetric distributions. If data
* is skewed then two approaches can be taken. One is transform
* the data using a Box-Cox or log transformation. The second
* is to calculate the upper reference limit by reflecting the
* data above the median. For example, suppose we have the following
* data set.
* [21, 26, 34, 41, 50, 55, 56, 61, 70]. Median is 50, so the
* reflected data set is [33, 39, 44, 45, 50, 56, 61, 70]. 
* Where reflected value = 2*median - x. eg. For 55 the
* reflected value = 100 - 55 = 45. For the a skewed population
* Horn recommends using the non-parametric estimator.
* 
* Horn, Paul S and Pesce, Amadeo J. 2005. Reference Intervals: A
* User's Guide. AACC Press, Washington DC.
* https://archive.org/details/referenceinterva0000horn/mode/2up
* 
*/

import { studentt, jStat } from 'jstat-esm';

import * as shapiro_wilk from './shapiro-wilk.js'

function median_absolute_deviation(data: number[], median: number): number {
    let dev = [data.length]
    for (let i = 0; i < data.length; i++) {
        dev[i] = Math.abs(data[i] - median);
    }
    return jStat.median(dev);
}

function robust(data: number[], alpha: number = 0.05): number[] {
    const n = data.length;
    data.sort((a, b) => a - b);
    const median = jStat.median(data);
    let tbi = median;
    let tbi_new = 10000;
    const c = 3.7
    let mad = median_absolute_deviation(data, median);
    mad = mad / 0.6745;
    let ui = [n]
    let wi = [n]
    let bi = [n]
    let diff = 1;
    do {
        for (let i = 0; i < n; i++) {
            ui[i] = (data[i] - tbi) / (c * mad)
            ui[i] = ui[i] < -1 ? 1 : ui[i];
            ui[i] = ui[i] > 1 ? 1 : ui[i];
            wi[i] = Math.pow(1 - Math.pow(ui[i], 2), 2);
            bi[i] = wi[i] * data[i];
        }
        tbi_new = jStat.sum(bi) / jStat.sum(wi);
        diff = Math.abs(tbi_new - tbi);
        tbi = tbi_new;
        //console.log(tbi);

    } while (diff > 0.000001);

    //console.log(`tbi = ${tbi}`);

    for (let i = 0; i < n; i++) {
        ui[i] = (data[i] - median) / (205.6 * mad);
    }
    const sbi205_6 = sbi(ui, mad, 205.6);
    //console.log(`sbi205.6 = ${sbi205_6}`);
    for (let i = 0; i < n; i++) {
        ui[i] = (data[i] - median) / (3.7 * mad);
    }
    const sbi3_7 = sbi(ui, mad, 3.7);
    //console.log(`sbi3.7 = ${sbi3_7}`);
    for (let i = 0; i < n; i++) {
        ui[i] = (data[i] - tbi) / (3.7 * sbi3_7);
    }
    const st3_7 = st(ui, sbi3_7, 3.7);
    //console.log(`st3.7 = ${st3_7}`);

    const t_statistic = studentt.inv(1 - alpha / 2, n - 1);
    const margin = t_statistic * Math.sqrt(Math.pow(sbi205_6, 2) + Math.pow(st3_7, 2));
    const robustLower = tbi - margin;
    const robustUpper = tbi + margin;
    return [robustLower, robustUpper];
}

function sbi(ui: number[], mad: number, c: number): number {
    let sub = [];
    let ai = [];
    let bi = [];
    for (let i = 0; i < ui.length; i++) {
        if (ui[i] > -1 && ui[i] < 1) {
            ai.push(Math.pow(1 - Math.pow(ui[i], 2), 4) * Math.pow(ui[i], 2));
            bi.push((1 - Math.pow(ui[i], 2)) * (1 - 5 * Math.pow(ui[i], 2)));
        }
    }
    let sum_a = jStat.sum(ai);
    let sum_b = jStat.sum(bi);
    //console.log(`sum_a = ${sum_a}, sum_b = ${sum_b}`);
    let d = jStat.max([1, sum_b - 1])
    let sbi = c * mad * Math.sqrt((ui.length * sum_a) / (sum_b * d));
    return sbi;
}

function st(ui: number[], mad: number, c: number): number {
    let ai = [];
    let bi = [];
    for (let i = 0; i < ui.length; i++) {
        if (ui[i] > -1 && ui[i] < 1) {
            ai.push(Math.pow(1 - Math.pow(ui[i], 2), 4) * Math.pow(ui[i], 2));
            bi.push((1 - Math.pow(ui[i], 2)) * (1 - 5 * Math.pow(ui[i], 2)));
        }
    }

    let sum_a = jStat.sum(ai);
    let sum_b = jStat.sum(bi);
    let d = jStat.max([1, sum_b - 1])
    let st = c * mad * Math.sqrt((sum_a) / (sum_b * d));
    //console.log(`c = ${c}, mad = ${mad}, sum_a = ${sum_a}, sum_b = ${sum_b}, d = ${d}, st = ${st}`);
    return st;
}

function sample_with_replacement(data: number[]): number[] {
    const indexes = Array.from({ length: data.length }, () => {
        return Math.floor(Math.random() * data.length);
    });
    let sample = [];
    for (let i = 0; i < data.length; i++) {
        sample.push(data[indexes[i]]);
    }
    return sample;
}

function bootstrap_confidence_interval(data: number[], alpha: number = 0.05, confAlpha = 0.1, n: number = 5000): [number[], number[]] {
    const lowerLimit: number[] = [n];
    const upperLimit: number[] = [n];
    for (let i = 0; i < n; i++) {
        let sample = sample_with_replacement(data);
        const ri = robust(sample, alpha);
        lowerLimit[i] = ri[0];
        upperLimit[i] = ri[1];
    }
    lowerLimit.sort((a, b) => a - b);
    upperLimit.sort((a, b) => a - b);
    const lower_limit_ci: number[] = [
        jStat.percentile(lowerLimit, confAlpha / 2),
        jStat.percentile(lowerLimit, 1 - confAlpha / 2)
    ];
    const upper_limit_ci: number[] = [
        jStat.percentile(upperLimit, confAlpha / 2),
        jStat.percentile(upperLimit, 1 - confAlpha / 2)
    ];
    return [lower_limit_ci, upper_limit_ci];
}

/**
 * Horn, P. S. (1988). A Biweight Prediction Interval for Random Samples. 
 * Journal of the American Statistical Association, 83(401), 249–256. 
 * https://doi.org/10.1080/01621459.1988.10478593
 * 
 * Calculation of the C2 tuning factor
 * 
 * Note: Had to go to the original reference to get the coefficients
 * to six decimal places. Otherwise it was causing rounding issues.
 * 
 * @param alpha 
 * @returns c2
 */
function c2(alpha: number): number {
    return 1 / (0.581734 - 0.607227 * (1 - alpha))
}

/** Equivalent to the range and seq functions in python and R
 * respectively
 * 
 * @param start lower limit of range (inclusive) 
 * @param stop upper limit of range
 * @param step increment
 * @returns array of sequential values.
 */
const range = (start: number, stop: number, step: number): number[] => {
    const n = Math.ceil((stop - start) / step);
    return Array.from({ length: n }, (_, i) => start + i * step);
};

/**
 * Box-Cox transformation using maximum likelihood estimation.
 * 
 * This algorithm was was generated by ChatGPT on 27/1/2026 
 * (free version GPT-5.2).

 * 
 * @param {number} x - positive data point
 * @param {number} lambda - transformation parameter
 * @returns {number}
 */
function boxCox(x: number, lambda: number): number {
  if (lambda === 0) {
    return Math.log(x);
  }
  return (Math.pow(x, lambda) - 1) / lambda;
}

/**
 * Compute log-likelihood for Box-Cox transformed data
 *
 * Generated by ChatGPT 27/1/2026
 * 
 * @param {number[]} data - original positive data
 * @param {number} lambda - transformation parameter
 * @returns {number}
 */
function boxCoxLogLikelihood(data: number[], lambda: number): number {
  const n = data.length;

  // Transform data
  const transformed = data.map(x => boxCox(x, lambda));

  // Mean
  const mean =
    transformed.reduce((sum, v) => sum + v, 0) / n;

  // Variance
  const variance =
    transformed.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / n;

  // Jacobian term
  const logJacobian =
    (lambda - 1) * data.reduce((sum, x) => sum + Math.log(x), 0);

  // Log-likelihood (up to additive constant)
  return -0.5 * n * Math.log(variance) + logJacobian;
}

/**
 * Estimate lambda via grid search MLE
 * 
 * This algorithm was was generated by ChatGPT on 27/1/2026 
 * (free version GPT-5.2).
 * 
 * It performs a simple grid search and calculates the log likelihood for
 * each potential lambda assessed. This simple approach is generally all that
 * is required when applied to reference intervals as the measurand 
 * concentrations are always positive. Also, for a suitably selected 
 * reference population, the results will almost certainly have a single 
 * mode.
 * 
 * @param {number[]} data - positive univariate data
 * @param {number} minLambda
 * @param {number} maxLambda
 * @param {number} step
 * @returns {{lambda: number, logLikelihood: number}}
 */
function estimateBoxCoxLambda(
  data: number[],
  minLambda = -2,
  maxLambda = 2,
  step = 0.01
): {lambda: number, logLikelihood: number} {
  if (data.some(x => x <= 0)) {
    throw new Error("Box-Cox requires strictly positive data.");
  }

  let bestLambda = null;
  let bestLL = -Infinity;

  for (let lambda = minLambda; lambda <= maxLambda; lambda += step) {
    const ll = boxCoxLogLikelihood(data, lambda);
    if (ll > bestLL) {
      bestLL = ll;
      bestLambda = lambda;
    }
  }

  return {
    lambda: Number(bestLambda.toFixed(4)),
    logLikelihood: bestLL
  };
}

interface BoxCoxTransform {
    lambda: number;
    transformedData: number[];
    shapiroWilkValue: number;
    maximumLikelihoodValue: number;
}

/** Estimate a box-cox transformation of supplied data using the Shapiro-Wilk
 * goodness of fit test for normality.
 * 
 * Asar Ö, Ilk O, Dag O. Estimating Box-Cox power transformation parameter via 
 * goodness-of-fit tests. Communications in Statistics - Simulation and
 * Computation 2017;46:91-105. Available at: 
 * https://doi.org/10.1080/03610918.2014.957839.
 * 
 * @param x data to be transformed
 * @param lambda search space for lambda (min, max, step)
 * @returns array [estimated lambda, transformed data, shapiro-wilk value ]
 */
function boxcoxsw(x: number[], lambda = [-2, 2, 0.01]): BoxCoxTransform {
    let lambda_arr = range(lambda[0], lambda[1], lambda[2]);
    const n = lambda_arr.length;
    const store2: number[][] = [];
    const store3: number[] = [];
    for (let i = 0; i < n; i++) {
        let t: number[] = []
        for (let j = 0; j < x.length; j++) {
            let v = boxCox(x[j], lambda_arr[i]);
            t.push(v);
            // let v = (lambda_arr[i] != 0) ? (Math.pow(x[j], lambda_arr[i]) - 1) / lambda_arr[i] : Math.log(x[j]);
            // t.push(v);
        }
        store2.push(t);
        store3.push(shapiro_wilk.ShapiroWilkW(t));
    }
    let k = store3.indexOf(Math.max(...store3));
    return {
        lambda: lambda_arr[k], 
        transformedData: store2[k],
        shapiroWilkValue: store3[k],
        maximumLikelihoodValue: NaN
    };
}

/** Estimate a box-cox transformation of supplied data using one
 * of two approaches.
 * 
 * The default is to calculate the goodness of fit using the 
 * Shapiro-Wilk test of normality.
 * 
 * The second is by maximum likelihood estimation
 * will be implemented at a later date.
 * 
 * @param x data to be transformed
 * @param lambda search space for lambda (min, max, step)
 * @param method for estimating lambda. May be one of ["sw","mle"]. Default method is Shapiro-Wilk.
 * @returns BoxCoxTransform object
 */
function boxcoxfit(x: number[], lambda = [-2, 2, 0.01], method="sw"): BoxCoxTransform {
    if (method === "sw") {
        return boxcoxsw(x, lambda);
    } else if (method === "mle") {
        const res = estimateBoxCoxLambda(x, lambda[0], lambda[1], lambda[2]);
        const transformed: number[] = [];
        for (let i = 0; i < x.length; i++) {
            transformed.push(boxCox(x[i], res.lambda));
        }
        return {
            lambda: res.lambda,
            transformedData: transformed,
            shapiroWilkValue: NaN,
            maximumLikelihoodValue: res.logLikelihood
        };
    } else {
        throw new Error("Unknown method: " + method);
    }
}

/** Inverse Box-Cox transformation.
 * 
 * @param x data to be transformed
 * @param lambda 
 * @returns an array containing the transformed data which is the same length as x.
 */
function boxcoxinv(x: number[], lambda: number): number[] {
    const n = x.length;
    const y: number[] = [n];
    for (let i = 0; i < n; i++) {
        if (lambda != 0) {
            y[i] = Math.pow((lambda * x[i]) + 1, lambda);
        } else {
            y[i] = Math.exp(x[i]);
        }
    }
    return y;
}


export {
    robust,
    median_absolute_deviation,
    sample_with_replacement,
    bootstrap_confidence_interval,
    c2,
    boxcoxfit,
    boxcoxinv
}