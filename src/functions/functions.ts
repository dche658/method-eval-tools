/**
 * Custom Functions
 */
import {boxCox} from '../reference_intervals.js'
import {ShapiroWilkW} from '../shapiro-wilk.js'

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
        y = Math.exp(Math.log((lambda * x) + 1) / lambda);
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
 * @param {number[]} x array of values
 * @returns {number[]} Shapiro-Wilk w and associated p value.
 */
export function shapirowilkw(x: number[]): number[] {
    const sw = ShapiroWilkW(x);
    return [sw.w, sw.p];
}
