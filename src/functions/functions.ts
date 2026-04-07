/**
 * Custom Functions
 */
import { boxCox } from "../reference_intervals";
import { ShapiroWilkW } from "../shapiro-wilk";

/**
 * Inverse of the Box-Cox transformation.
 *
 * @customfunction BOXCOXINV
 * @param {number} x value to be transformed
 * @param {number} lambda exponent
 * @returns {number} inverse of the Box-Cox transformation
 */
export function boxcoxinv(x: number, lambda: number): number {
  let y: number;
  if (lambda != 0) {
    y = Math.exp(Math.log(lambda * x + 1) / lambda);
  } else {
    y = Math.exp(x);
  }
  return y;
}

/**
 * Box-Cox transformation
 *
 * @customfunction BOXCOX
 * @param {number} x value to be transformed
 * @param {number} lambda exponent
 * @returns {number} transformed value
 */
export function boxcox(x: number, lambda: number): number {
  return boxCox(x, lambda);
}

/**
 * Shapiro-Wilk W statistic for normality.
 *
 * @customfunction SHAPIROWILKW
 * @param {number[][]} x array of values
 * @returns {any[][]} Shapiro-Wilk w and associated p value.
 */
export function shapirowilkw(x: number[][]): any[][] {
  const arr: number[] = [];
  x.forEach((row) => {
    row.forEach((value) => {
      arr.push(value);
    });
  });
  const sw = ShapiroWilkW(arr);
  return [[sw.w], [sw.p]];
}

/**
 * Dixon-Reed test for outliers
 * 
 * Reed AH, Henry RJ, Mason WB. Influence of statistical method 
 * used on the resulting estimate of normal range. Clinical Chemistry
 * 1971;17:275-284.
 * 
 * @customfunction DIXONREED
 * @param {number[][]} x array of values
 * @returns {any[][]} Array containing ratio and outlier
 * 
 */
export function dixonreed(x: number[][]): any[][] {
  const arr: number[] = [];
  const outliers: any[] = [];
  const ratios: number[] = [];
  x.forEach((row) => {
    row.forEach((value) => {
      arr.push(value);
    });
  });
  arr.sort((a,b) => a - b);
  let r = 0;
  const n = arr.length - 1;
  if (arr.length > 2) {
    //Check for low outliers
    r = (arr[1] - arr[0]) / (arr[n] - arr[0]);
    ratios.push(r);
    outliers.push(r > 1/3 ? arr[n]: undefined);
    //Check for high outliers
    r = (arr[n] - arr[n-1]) / (arr[n] - arr[0]);
    ratios.push(r);
    outliers.push(r > 1/3 ? arr[n]: undefined);
  }
  return [["Low", "High"],[ratios[0],ratios[1]],[outliers[0], outliers[1]]];
}
