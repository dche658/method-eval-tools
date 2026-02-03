/**
 * Inverse of the Box-Cox transformation.
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

