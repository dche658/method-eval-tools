/**
 * Reference intervals
 * 
 * @author Douglas Chesher
 */

import { studentt, normal } from 'jstat-esm';
import jStat from 'jstat-esm';

import * as shapiro_wilk from './shapiro-wilk.js'

import { quantile } from './regression'

const NP = "np"; //non parametric
const ROBUST = "robust";

/**
 * Calculate MAD for robust algorithm
 * @param data 
 * @param median 
 * @returns 
 */
function median_absolute_deviation(data: number[], median: number): number {
    let dev = [data.length]
    for (let i = 0; i < data.length; i++) {
        dev[i] = Math.abs(data[i] - median);
    }
    return jStat.median(dev);
}

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
 * 
 * @param data array of values
 * @param alpha Default 0.05 for 95% confidence interval.
 * @returns reference limits and their confidence intervals
 */
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

/**
 * Take a random sample with replacement
 * @param data array of values
 * @returns results of sampling as an array with the same length as the original data.
 */
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

/**
 * Calculate the bootstrap confidence interval for the robust method
 * @param data array of values
 * @param alpha Default 0.05 for 95% confidence interval. 
 * @param confAlpha Alpha for the confidence limits for the reference limit. Default is 0.1 for 90% confidence
 * @param n number of bootstrap samples. Default is 5000.
 * @returns 
 */
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
 * Reference Interval results object
 */
interface ReferenceInterval {
    lowerLimit: number;
    upperLimit: number;
    lowerLimitCI: number[];
    upperLimitCI: number[];
}

/**
 * Calculate the reference interval based on the supplied univariate data.
 * 
 * @param data an array of values
 * @param alpha Default 0.05 for 95% confidence interval.
 * @param confAlpha Alpha for the confidence limits for the reference limit. Default is 0.1 for 90% confidence
 * @param method "robust" or "np" only
 * @param bootstrap_n. Number of bootstrap samples. Only used by robust method at present but could be used for the non-parametric method. 
 * @returns reference limits and their confidence intervals
 */
function ref_limit(data: number[], alpha: number = 0.05, confAlpha = 0.1, method=ROBUST, bootstrap_n: number = 5000): ReferenceInterval {
    let refint:ReferenceInterval;
    if (method === ROBUST) {
        const ri = robust(data, alpha);
        const ci = bootstrap_confidence_interval(data, alpha, confAlpha, bootstrap_n);
        refint = {
            lowerLimit: ri[0],
            upperLimit: ri[1],
            lowerLimitCI: ci[0],
            upperLimitCI: ci[1]
        }
    } else if (method === NP) {
        refint = non_parametric(data, alpha, confAlpha);
    } else {
        throw new Error("Unknown method: " + method);
    }
    return refint;
}

/**
 * Non-parametric method for determining reference intervals.
 * The method used for calculating the confidence limits for the reference intervals is different to that
 * used by CLSI. CLSI C28-A3 references Reed AH, Henry RJ, Mason WB. Influence of statistical method used 
 * on the resulting estimate of normal range. Clin Chem. 1971 Apr;17(4):275-84. PMID: 5552364. However,
 * this paper provides an inequality to determine the upper and lower ranks and then lists the values in 
 * Table 3. In contrast Campbell and Gardner provide an alternative approch for their numerical derivation,
 * which is used below.
 * 
 * @param data population data
 * @param alpha proportion in confidence limit. Default is 0.05 for 95%
 * @param confAlpha level of significance for the reference limit confidence intervals. Default is 0.1 for 90%.
 * @returns reference limits and confidence intervals.
 */
function non_parametric(data: number[], alpha: number = 0.05, confAlpha = 0.1): ReferenceInterval {
    const n = data.length;
    data.sort((a, b) => a - b);
    if (confAlpha <= 0.05 && n < 153) {
        throw new Error("At least 153 subjects are required to calculate 95% confidence intervals using a non-parametric method.");
    }
    if (confAlpha <= 0.05 && n < 120) {
        throw new Error("At least 120 subjects are required to calculate 90% confidence intervals using a non-parametric method.")
    }
    const lowerLimitRank = Math.round(n * alpha / 2);
    const lowerLimitIndex = lowerLimitRank < 1 ? 0 : lowerLimitRank - 1;
    const lowerLimit = data[lowerLimitIndex];
    const upperLimitRank = Math.round(n * (1 - alpha / 2));
    const upperLimitIndex = upperLimitRank > n ? n - 1 : upperLimitRank - 1;
    const upperLimit = data[upperLimitIndex];

    const lowerLimitLCLRank = lower_confidence_limit_rank(n, alpha/2, confAlpha);
    const lowerLimitLCLIndex = lowerLimitLCLRank < 1 ? 0 : lowerLimitLCLRank - 1;
    const lowerLimitLCL = data[lowerLimitLCLIndex];
    const lowerLimitUCLRank = upper_confidence_limit_rank(n, alpha/2, confAlpha);
    const lowerLimitUCLIndex = lowerLimitUCLRank > n ? n - 1 : lowerLimitUCLRank - 1;
    const lowerLimitUCL = data[lowerLimitUCLIndex];

    const upperLimitLCLRank = lower_confidence_limit_rank(n, 1 - alpha/2, confAlpha);
    const upperLimitLCLIndex = upperLimitLCLRank < 1 ? 0 : upperLimitLCLRank - 1;
    const upperLimitLCL = data[upperLimitLCLIndex];
    const upperLimitUCLRank = upper_confidence_limit_rank(n, 1 - alpha/2, confAlpha);
    const upperLimitUCLIndex = upperLimitUCLRank > n ? n - 1 : upperLimitUCLRank - 1;
    const upperLimitUCL = data[upperLimitUCLIndex];

    return {
        lowerLimit: lowerLimit,
        upperLimit: upperLimit,
        lowerLimitCI: [lowerLimitLCL, lowerLimitUCL],
        upperLimitCI: [upperLimitLCL, upperLimitUCL]
    };

}

/**
 * Campbell MJ, Gardner MJ. Calculating confidence intervals for some non-parametric analyses. 
 * Br Med J (Clin Res Ed). 1988 May 21;296(6634):1454-6. doi: 10.1136/bmj.296.6634.1454. 
 * PMID: 3132290; PMCID: PMC2545906.
 * 
 * @param n 
 * @param q 
 * @param alpha 
 */
function lower_confidence_limit_rank(n: number, q: number = 0.025, alpha = 0.1): number {
    let r = (n*q)-(normal.inv(1-alpha/2, 0, 1)*Math.sqrt(n*q*(1-q)));
    r = Math.round(r);
    return r;
}

/**
 * Campbell MJ, Gardner MJ. Calculating confidence intervals for some non-parametric analyses. 
 * Br Med J (Clin Res Ed). 1988 May 21;296(6634):1454-6. doi: 10.1136/bmj.296.6634.1454. 
 * PMID: 3132290; PMCID: PMC2545906.
 * 
 * @param n 
 * @param q 
 * @param alpha 
 */
function upper_confidence_limit_rank(n: number, q: number = 0.975, alpha = 0.1): number {
    let s = 1 + (n*q) + normal.inv(1-alpha/2, 0, 1)*Math.sqrt(n*q*(1-q));
    s = Math.round(s);
    return s;
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
    shapiroWilkValue: {w: number, p: number};
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
    const store4: {w: number, p: number}[] = [];
    for (let i = 0; i < n; i++) {
        let t: number[] = []
        for (let j = 0; j < x.length; j++) {
            let v = boxCox(x[j], lambda_arr[i]);
            t.push(v);
            // let v = (lambda_arr[i] != 0) ? (Math.pow(x[j], lambda_arr[i]) - 1) / lambda_arr[i] : Math.log(x[j]);
            // t.push(v);
        }
        store2.push(t);
        const sw = shapiro_wilk.ShapiroWilkW(t);
        store3.push(sw.w);
        store4.push(sw);
    }
    let k = store3.indexOf(Math.max(...store3));
    return {
        lambda: lambda_arr[k], 
        transformedData: store2[k],
        shapiroWilkValue: store4[k],
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
            shapiroWilkValue: {w: NaN, p: NaN},
            maximumLikelihoodValue: res.logLikelihood
        };
    } else {
        throw new Error("Unknown method: " + method);
    }
}


/** Adjusted Fisher Pearson Coefficient of Skewness.
 * 
 * G1 > 0 Skewed to the right
 * G1 < 0 Skewed to the left
 * G1 = 0 Symmetric 
 * 
 * Equation for G1 taken from
 * 1.3.5.11 Measures of Skewness and Kurtosis. In
 * NIST/SEMATECH e-Handbook of Statistical Methods.
 * https://www.itl.nist.gov/div898/handbook/eda/section3/eda35b.htm
 * Accessed 2/2/2026
 *  
 * 
 * @param data 
 * @returns G1
 */
function adjusted_fisher_pearson_coefficient(data: number[]): number {
    const n = data.length;
    const mean = jStat.mean(data);
    const s = jStat.stdev(data);
    let sumDevX = 0;
    for (let i = 0; i < n; i++) {
        sumDevX += Math.pow(data[i] - mean, 3)/n;
    }
    const g1 = Math.sqrt(n*(n-1))/(n-2) * sumDevX / Math.pow(s, 3);
    return g1;
}

/**
 * Variance of G1 assuming a normal population
 * 
 * Equation from Wikipedia https://en.wikipedia.org/wiki/Skewness
 * Accessed 2/2/2026
 * 
 * 
 * @param n sample size
 * @returns var(G1)
 */
function var_g1(n:number): number {
    return 6*n*(n-1)/((n-2)*(n+1)*(n+3))
}

/**
 * Assess whether G1 is close to zero and consistent with
 * a normal population without any significant skew.
 * 
 * Significance at 90% level by default.
 * 
 * @param n sample size
 * @param g1 statistic
 * @param alpha significance, default 0.1 for 90%
 * @returns 
 */
function is_consistent_with_normal(n: number, g1: number, alpha: number = 0.1): boolean {
    const z = normal.inv(1-alpha/2, 0, 1);
    const varg1 = var_g1(n);
    const cl = z * Math.sqrt(varg1);
    //console.log(`g1 ci= ${cl}`);
    if (Math.abs(g1) < z * Math.sqrt(varg1)) {
        return true;
    } else {
        return false;
    }
}


export {
    robust,
    median_absolute_deviation,
    sample_with_replacement,
    bootstrap_confidence_interval,
    c2,
    boxcoxfit,
    BoxCoxTransform,
    ref_limit,
    NP,
    ROBUST,
    ReferenceInterval,
    adjusted_fisher_pearson_coefficient,
    is_consistent_with_normal
}