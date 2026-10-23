/** @module reflim
 * 
 * Javascript implimentation of reflimR package for R.
 *
 */

import { arange, seq, quantiles, min, max, normal, median, quartiles } from 'jstat-esm';

import { LeastSquaresRegression } from './lsregression';

/**
 * Bowley quantile function
 * @param x array of numbers
 * @param alpha default is 0.25
 * @returns 
 */
function bowley(x: number[], alpha = 0.25): number {
    let q = quantiles(x, [alpha, 0.5, 1 - alpha]);
    let bs = ((q[0] - 2 * q[1] + q[2]) / (q[2] - q[0]));
    return bs
}

/**
 * Vectorised log (base e) function
 * @param x array of numbers
 * @returns natural logarithm of the submitted values
 */
function log(x: number[]): number[] {
    const y = x.map(x => Math.log(x));
    return y;
}

/**
 * Interquartile range
 * @param x 
 * @returns 
 */
function iqr(x: number[]) {
    const q = quartiles(x);
    return q[2] - q[0];
}

/**
 * Check skewness of the data.
 * 
 * @param x 
 * @param cutoff 
 * @param alpha 
 * @returns 
 */
function lognorm(x: number[], cutoff = 0.05, alpha = 0.25, digits: null | number = 3) {
    const lowest = min(x)
    if (lowest < 0) {
        throw new Error(`Minimum value ${lowest} is less than 0`);
    }
    const bs = [bowley(x, alpha), bowley(log(x), alpha)]
    let lognormal = false;
    if (bs[0] < 0) {
        lognormal = false;
    } else {
        lognormal = (bs[0] - bs[1]) >= cutoff
    }
    let delta = bs[0] - bs[1];
    if (digits != null) {
        bs[0] = round(bs[0], digits);
        bs[1] = round(bs[1], digits);
        delta = round(delta, digits);
    }
    // Unable to implement the lognorm plot in reflimR with out the kernel density function
    return {
        "lognormal": lognormal,
        "BowleySkewness": {
            "normal": bs[0],
            "lognormal": bs[1],
            "delta": delta
        }
    }
}

function truncate(x: number[], qf: number): number[] {
    const q = quantiles(x, [0.25, 0.5, 0.75])
    const var1 = q[1] - q[0];
    const var2 = q[2] - q[1];
    const v = Math.min(var1, var2);
    const lim = [q[1] - qf * v, q[1] + qf * v];
    const subset = x.filter(val => val >= lim[0] && val <= lim[1])
    return subset;
}

function print_progress(x: number[], lognormal = false) {
    let y = x;
    if (lognormal) {
        y = x.map(v => Math.log(v));
    }
    return {
        "n": y.length,
        "min": min(y),
        "max": max(y)
    };
}

function adjust_digits(x: number) {
    if (x == 0) {
        return {
            "x_round": 0,
            "digits": 0
        }
    }
    let xx = Math.abs(x);
    let digits = 2 - Math.floor(Math.log10(xx));
    if (digits < 0) {
        digits = 0;
    }
    xx = Math.round(x * Math.pow(10, digits)) / Math.pow(10, digits);
    return {
        "x_round": xx,
        "digits": digits
    }
}

function round(x: number, digits: number) {
    return Math.round(x * Math.pow(10, digits)) / Math.pow(10, digits);
}

function iboxplot(x: number[], lognormal: boolean | null = null, perc_trunc = 2.5, apply_rounding = true) {
    let xx = x;
    const digits = adjust_digits(median(x)).digits;
    const n = x.length;
    if (lognormal == null) {
        lognormal = lognorm(x).lognormal;
    }
    if (lognormal) {
        xx = x.map(v => Math.log(v))
    }
    let progress = [{
        "n": x.length,
        "min": min(x),
        "max": max(x)
    }];
    const q_trunc = perc_trunc / 100;
    const steps = 5;
    for (let i = 0; i < steps; i++) {
        const target_quantile = 1 - (q_trunc * i / steps);
        const alpha = (2 * q_trunc) * (i - 1) / steps;
        const qf = normal.inv(target_quantile, 0, 1) / normal.inv(0.75 - 0.25 * alpha, 0, 1);
        xx = truncate(xx, qf);
        progress = [...progress, print_progress(xx, lognormal)];
    }

    let n0 = 1;
    let n1 = 0;
    let qf = normal.inv(1 - q_trunc, 0, 1) / normal.inv(0.75 - 0.5 * q_trunc, 0, 1);
    let i = 0;
    while (n0 > n1) {
        i += 1;
        n0 = xx.length;
        xx = truncate(xx, qf)
        n1 = xx.length
        progress = [...progress, print_progress(xx, lognormal)];
    }
    if (lognormal) {
        xx = xx.map(x => Math.exp(x));
    }
    const lim = {
        "lower": min(xx),
        "upper": max(xx)
    }
    if (apply_rounding) {
        lim.lower = round(lim.lower, digits);
        lim.upper = round(lim.upper, digits)
    }
    let prop = round(xx.length * 100 / (1 - (2 * perc_trunc / 100)) / n, 1);
    if (prop > 100) prop = 100;
    progress.forEach(p => {
        p.min = round(p.min, digits);
        p.max = round(p.max, digits);
    });
    return {
        "trunc": xx,
        "truncation.points": lim,
        "lognormal": lognormal,
        "perc.norm": prop,
        "progress": progress
    }
}

interface QQPlotResult {
    "mean": number,
    "sd": number,
    "lower_lim": number,
    "upper_lim": number,
    "lognormal": boolean | null,
    "chart_data": {
        "xlim": number[],
        "ylim": number[],
        "data_y": number[],
        "data_x": number[],
        "main": string,
        "xlab": string,
        "ylab": string
    },
}

function truncated_qqplot(x_trunc: number[], lognormal: boolean | null = null, perc_trunc = 2.5, n_min = 200,
    apply_rounding = true, plot_it = true, main = "Q-Q plot", xlab = "theoretical quantiles", ylab = "sample quantiles") {
    const n = x_trunc.length;
    if (n < n_min) console.log(`Truncated data set has ${n} values, which is less than the minimum requirement of ${n_min}`)
    let n_quantiles = 100;
    if (n < n_quantiles) n_quantiles = n;
    const digits = adjust_digits(median(x_trunc)).digits;
    if (lognormal == null) {
        lognormal = lognorm(x_trunc).lognormal;
    }
    if (lognormal) {
        x_trunc = x_trunc.map(x => Math.log(x));
    }
    let p1: number[] = seq(perc_trunc / 100, 1 - perc_trunc / 100, n_quantiles);
    let p2: number[] = seq(0, 1, n_quantiles);
    //console.log(p1,p2);
    let x_ax: number[] = p1.map(x => normal.inv(x, 0, 1));
    let y_ax: number[] = quantiles(x_trunc, p2, 1, 1); 
    //alphap and betap of quantiles set to 1 and 1 to get same results as R.
    //see https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.mstats.mquantiles.html
    //R type 7: https://en.wikipedia.org/wiki/Quantile
    //console.log(x_ax, y_ax);
    let central_part_from = Math.floor(0.05 * n_quantiles);
    let central_part_to = Math.ceil(0.95 * n_quantiles);
    let central_part: number[] = arange(central_part_from, central_part_to+1, 1);
    //console.log(central_part);
    let y: number[] = central_part.map(i => y_ax[i-1]);
    let x: number[] = central_part.map(i => x_ax[i-1]);
    //console.log(x,y);
    const ols = new LeastSquaresRegression();
    const reg = ols.calculate(x, y);
    //console.log(reg);
    let result: QQPlotResult = {
        "mean": reg.intercept,
        "sd": reg.slope,
        "lower_lim": reg.intercept - 1.96 * reg.slope,
        "upper_lim": reg.intercept + 1.96 * reg.slope,
        "lognormal": lognormal,
        "chart_data": {
            "xlim": [],
            "ylim": [],
            "data_y": y_ax,
            "data_x": x_ax,
            "main": main,
            "xlab": xlab,
            "ylab": ylab
        }
    }
    if (lognormal) {
        result.lower_lim = Math.exp(result.lower_lim);
        result.upper_lim = Math.exp(result.upper_lim);
    }
    if (apply_rounding) {
        result.lower_lim = round(result.lower_lim, digits);
        result.upper_lim = round(result.upper_lim, digits);
    }
    // console.log(n_quantiles);
    // console.log(x_trunc);
    // console.log(lognormal)
    // console.log(x_ax);
    // console.log(y_ax);
    if (plot_it) {
        const ll = lognormal ? Math.log(result.lower_lim) : result.lower_lim;
        const ul = lognormal ? Math.log(result.upper_lim) : result.upper_lim;
        const diff = ul - ll;
        const sig = adjust_digits(ll).digits;
        result.chart_data.xlim = [-3, 3];
        result.chart_data.ylim = [round(ll - 0.2 * diff, sig), round(ul + 0.2 * diff, sig)];
    }
    return result;
}

interface ConfIntResult {
    "lower_lim_low": number | undefined,
    "lower_lim_upp": number | undefined,
    "upper_lim_low": number | undefined,
    "upper_lim_upp": number | undefined,
    "n": number
}
function conf_int95(n: number, lower_limit: number, upper_limit: number, lognormal: boolean | null = true, apply_rounding = true) {
    if (upper_limit <= lower_limit) {
        throw new Error("upper.limit must be higher than lower.limit.");
    }
    if (lognormal) {
        lower_limit = Math.log(lower_limit);
        upper_limit = Math.log(upper_limit);
    }
    const sigma = (upper_limit - lower_limit) / 3.92
    let result: ConfIntResult = {
        "lower_lim_low": undefined,
        "lower_lim_upp": undefined,
        "upper_lim_low": undefined,
        "upper_lim_upp": undefined,
        "n": 0
    }

    const diff_outer = sigma * 5.81 / (Math.sqrt(n) + 0.66)
    const diff_inner = sigma * 7.26 / (Math.sqrt(n) - 5.58)
    result.lower_lim_low = lower_limit - diff_outer;
    result.lower_lim_upp = lower_limit + diff_inner;
    result.upper_lim_low = upper_limit - diff_inner;
    result.upper_lim_upp = upper_limit + diff_outer;
    result.n = n;
    if (lognormal) {
        result.lower_lim_low = Math.exp(result.lower_lim_low);
        result.lower_lim_upp = Math.exp(result.lower_lim_upp);
        result.upper_lim_low = Math.exp(result.upper_lim_low);
        result.upper_lim_upp = Math.exp(result.upper_lim_upp);
    }
    if (apply_rounding) {
        const digits = adjust_digits(result.lower_lim_low).digits;
        result.lower_lim_low = round(result.lower_lim_low, digits);
        result.lower_lim_upp = round(result.lower_lim_upp, digits);
        result.upper_lim_low = round(result.upper_lim_low, digits);
        result.upper_lim_upp = round(result.upper_lim_upp, digits);
    }
    return result;
}

interface RefLimResult {
    "stats": {
        "mean": undefined | number,
        "sd": undefined | number,
        "n_total": number,
        "n_trunc": number,
    },
    "lognormal": boolean | null,
    "limits": {
        "lower_lim": undefined | number,
        "upper_lim": undefined | number,
        "lower_lim_low": undefined | number,
        "lower_lim_upp": undefined | number,
        "upper_lim_low": undefined | number,
        "upper_lim_upp": undefined | number
    },
    "targets": {
        "lower_lim": undefined | number,
        "upper_lim": undefined | number,
        "lower_lim_low": undefined | number,
        "lower_lim_upp": undefined | number,
        "upper_lim_low": undefined | number,
        "upper_lim_upp": undefined | number
    },
    "perc_norm": undefined | number,
    "confidence_int": ConfIntResult,
    "interpretation": InterpResult | null,
    "remarks": string,
    "res_lognorm": object | null,
    "res_trunc": object | null,
    "res_qq": QQPlotResult | null,
    "dev_lim": string[]
}

interface Limits {
    "lower_lim": undefined | number,
    "upper_lim": undefined | number,
    "lower_lim_low": undefined | number,
    "lower_lim_upp": undefined | number,
    "upper_lim_low": undefined | number,
    "upper_lim_upp": undefined | number
}

interface PermissibleUncertaintyResult {
    "lower_lim_low": number,
    "lower_lim_upp": number,
    "upper_lim_low": number,
    "upper_lim_upp": number,
}

function permissible_uncertainty(lower_limit: number, upper_limit: number, apply_rounding: boolean | null = true): PermissibleUncertaintyResult {
    if (lower_limit <= 0 || upper_limit <= 0 || lower_limit >= upper_limit) {
        throw new Error("Please check lower and upper limits: only positive values allowed.");
    }

    const g = Math.sqrt(lower_limit * upper_limit);
    //console.log(g);
    const slog = (Math.log(upper_limit) - Math.log(lower_limit)) / 3.92;
    //console.log(slog);
    const CV_E = 100 * (Math.sqrt(Math.exp(slog * slog) - 1));
    //console.log(CV_E);
    const f = Math.sqrt(CV_E - 0.25);
    //console.log(f);
    const pU1 = (f * (1.024 * lower_limit + 0.256 * g)) / 100;
    //console.log(pU1);
    const pU2 = (f * (1.024 * upper_limit + 0.256 * g)) / 100;
    //console.log(pU2);

    const result: PermissibleUncertaintyResult = {
        "lower_lim_low": lower_limit - pU1,
        "lower_lim_upp": lower_limit + pU1,
        "upper_lim_low": upper_limit - pU2,
        "upper_lim_upp": upper_limit + pU2,
    }
    if (apply_rounding) {
        const digits = adjust_digits(result.lower_lim_low).digits;
        result.lower_lim_low = round(result.lower_lim_low, digits);
        result.lower_lim_upp = round(result.lower_lim_upp, digits);
        result.upper_lim_low = round(result.upper_lim_low, digits);
        result.upper_lim_upp = round(result.upper_lim_upp, digits);
    }
    return (result)
}

interface InterpResult {
    "tol_lim": PermissibleUncertaintyResult,
    "tol_tar": PermissibleUncertaintyResult,
    "col_lim": string[],
    "col_tar": string[],
    "dev_lim": string[]
}

function interpretation(limits: number[], targets: number[]): InterpResult {
    if (limits.length != 2) throw new Error("Limits must be a vector of length 2.");
    if (limits[0] >= limits[1]) throw new Error("The upper limit must be greater than the lower limit.");
    if (limits[0] <= 0) throw new Error("Only positive target values are allowed.");
    if (targets.length != 2) throw new Error("Limits must be a vector of length 2.");
    if (targets[0] >= targets[1]) throw new Error("The upper target must be greater than the lower target.");
    if (targets[0] <= 0) throw new Error("Only positive target values are allowed.");
    const GREY = "grey";
    const GREEN = "green";
    const YELLOW = "yellow";
    const RED = "red";
    const tol_lim: PermissibleUncertaintyResult = permissible_uncertainty(limits[0], limits[1]);
    const col_lim: string[] = [GREY, GREY];
    const dev_lim = ["within tolerance", "within tolerance"];
    const tol_tar: PermissibleUncertaintyResult = permissible_uncertainty(targets[0], targets[1]);
    const col_tar = [...col_lim]; //clone
    if (limits[0] >= tol_tar.lower_lim_low && limits[0] <= tol_tar.lower_lim_upp) {
        col_lim[0] = GREEN;
    } else {
        if ((tol_lim.lower_lim_low >= tol_tar.lower_lim_low && tol_lim.lower_lim_low <= tol_tar.lower_lim_upp) ||
            (tol_lim.lower_lim_upp >= tol_tar.lower_lim_low && tol_lim.lower_lim_upp <= tol_tar.lower_lim_upp)) {
            col_lim[0] = YELLOW
        } else { col_lim[0] = RED; }
    }
    if (limits[1] >= tol_tar.upper_lim_low && limits[1] <= tol_tar.upper_lim_upp) {
        col_lim[1] = GREEN;
    } else {
        if ((tol_lim.upper_lim_low >= tol_tar.upper_lim_low && tol_lim.upper_lim_low <= tol_tar.upper_lim_upp) ||
            (tol_lim.upper_lim_upp >= tol_tar.upper_lim_low && tol_lim.upper_lim_upp <= tol_tar.upper_lim_upp)) {
            col_lim[1] = YELLOW;
        } else {
            col_lim[1] = RED;
        }
    }
    for (let i = 0; i < 2; i++) {
        if (col_lim[i] == YELLOW) {
            if (limits[i] < targets[i]) {
                dev_lim[i] = "slightly decreased"
            } else {
                dev_lim[i] = "slightly increased"
            }
        }
        if (col_lim[i] == RED) {
            if (limits[i] < targets[i]) {
                dev_lim[i] = "markedly decreased"
            } else {
                dev_lim[i] = "markedly increased"
            }
        }
    }
    return {
        "tol_lim": tol_lim,
        "tol_tar": tol_tar,
        "col_lim": col_lim,
        "col_tar": col_tar,
        "dev_lim": dev_lim
    };
}

function reflim(x: number[], lognormal: boolean | null = null, targets: number[] | null = null, perc_trunc = 2.5, n_min = 200, apply_rounding = true) {
    if (min(x) <= 0) {
        throw new Error("Only positive values are allowed")
    }
    if (targets != null && !Array.isArray(targets)) {
        throw new Error("Target must be an array of length two");
    }
    const result: RefLimResult = {
        "stats": {
            "mean": undefined,
            "sd": undefined,
            "n_total": 0,
            "n_trunc": 0,
        },
        "lognormal": lognormal,
        "limits": {
            "lower_lim": undefined,
            "upper_lim": undefined,
            "lower_lim_low": undefined,
            "lower_lim_upp": undefined,
            "upper_lim_low": undefined,
            "upper_lim_upp": undefined
        },
        "targets": {
            "lower_lim": undefined,
            "upper_lim": undefined,
            "lower_lim_low": undefined,
            "lower_lim_upp": undefined,
            "upper_lim_low": undefined,
            "upper_lim_upp": undefined
        },
        "perc_norm": undefined,
        "confidence_int": {
            "lower_lim_low": undefined,
            "lower_lim_upp": undefined,
            "upper_lim_low": undefined,
            "upper_lim_upp": undefined,
            "n": 0
        },
        "interpretation": null,
        "remarks": "",
        "res_lognorm": null,
        "res_trunc": null,
        "res_qq": null,
        "dev_lim": ["",""],
    }
    let xx = x;
    const n = xx.length;
    const digits = adjust_digits(median(xx)).digits;
    let plot_logtype = false;
    if (lognormal == null) {
        plot_logtype = true;
        lognormal = lognorm(xx).lognormal;
    } else {
        plot_logtype = false;
    }
    result.res_lognorm = lognorm(xx); //in reflimR reflim.R but never used.
    let res_trunc = iboxplot(xx, lognormal, perc_trunc, apply_rounding);
    result.res_trunc = res_trunc;
    const n_trunc = res_trunc.trunc.length;
    if (n_trunc < 40) {
        console.log(`Warning: n = ${n_trunc} after truncation. The absolute limit for reference limit estimation is 40.`);
        result.stats.n_total = n;
        result.stats.n_trunc = n_trunc;
        result.remarks = "n < 40 after truncation."
        return result;
    }
    let res_qq = truncated_qqplot(res_trunc.trunc, lognormal, perc_trunc, n_min, apply_rounding);
    result.res_qq = res_qq;
    let res_ci = conf_int95(n, res_qq.lower_lim, res_qq.upper_lim, lognormal, apply_rounding);
    let res_pu: PermissibleUncertaintyResult | null = null;
    if (res_qq.lower_lim > 0) {
        res_pu = permissible_uncertainty(res_qq.lower_lim, res_qq.upper_lim, apply_rounding);
    } else {
        console.log("Estimated lower limit <= 0. No tolerance limits calculated. No graphics produced.");
        targets = null;
        result.remarks = "Lower limit <= 0";
    }
    const res_lim: Limits = {
        "lower_lim": res_qq.lower_lim,
        "upper_lim": res_qq.upper_lim,
        "lower_lim_low": res_pu == null ? undefined : res_pu.lower_lim_low,
        "lower_lim_upp": res_pu == null ? undefined : res_pu.lower_lim_upp,
        "upper_lim_low": res_pu == null ? undefined : res_pu.upper_lim_low,
        "upper_lim_upp": res_pu == null ? undefined : res_pu.upper_lim_upp
    }
    const res_tar: Limits = {
        "lower_lim": undefined,
        "upper_lim": undefined,
        "lower_lim_low": undefined,
        "lower_lim_upp": undefined,
        "upper_lim_low": undefined,
        "upper_lim_upp": undefined
    }
    if (targets != null) {
        const limits: number[] = [
            res_lim.lower_lim == undefined? 0: res_lim.lower_lim,
            res_lim.upper_lim == undefined? 0: res_lim.upper_lim
        ];
        const ip = interpretation(limits, targets);
        res_tar.lower_lim = targets[0];
        res_tar.upper_lim = targets[1];
        res_tar.lower_lim_low = ip.tol_tar.lower_lim_low;
        res_tar.lower_lim_upp = ip.tol_tar.lower_lim_upp;
        res_tar.upper_lim_low = ip.tol_tar.upper_lim_low;
        res_tar.upper_lim_upp = ip.tol_tar.upper_lim_upp;
        if (apply_rounding) {
            res_tar.lower_lim = round(res_tar.lower_lim, digits);
            res_tar.upper_lim = round(res_tar.upper_lim, digits);
        }
        result.dev_lim = ip.dev_lim;
        result.interpretation = ip;
    }
    result.stats.mean = round(res_qq.mean, 3);
    result.stats.sd = round(res_qq.sd,3);
    result.stats.n_total = n;
    result.stats.n_trunc = n_trunc;
    result.lognormal = lognormal;
    result.limits = res_lim;
    result.targets = res_tar;
    result.perc_norm = res_trunc['perc.norm'];
    result.confidence_int = res_ci

    return result;
}

export {
    adjust_digits,
    bowley,
    iboxplot,
    interpretation,
    iqr,
    log,
    lognorm,
    permissible_uncertainty,
    reflim,
    round,
    truncated_qqplot,
    QQPlotResult,
    ConfIntResult,
    RefLimResult,
    InterpResult,
    Limits
}