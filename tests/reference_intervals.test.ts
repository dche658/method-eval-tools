import jStat from 'jstat-esm';
import { normal } from 'jstat-esm';
import { test, expect } from '@jest/globals';
import { ShapiroWilkW } from '../src/shapiro-wilk';
import {
    robust,
    median_absolute_deviation,
    sample_with_replacement,
    bootstrap_confidence_interval,
    c2,
    boxcoxfit,
    adjusted_fisher_pearson_coefficient,
    is_consistent_with_normal,
} from '../src/reference_intervals';


let data = [74.64731, 77.25439, 81.64253, 82.37364, 84.96763, 85.83717, 86.15544, 86.25899, 86.32812, 88.06514, 88.30495,
    89.39985, 89.50754, 90.04025, 90.47054, 90.62720, 90.67660, 91.00039, 91.03980, 91.35050, 92.34147, 92.43501, 92.64297,
    92.81908, 92.93594, 93.40938, 93.61053, 93.61496, 93.82136, 94.99198, 95.25909, 95.26355, 95.58339, 95.79880, 96.03610,
    96.40177, 97.05810, 97.35304, 97.48031, 97.50916, 97.86968, 97.93678, 98.00245, 98.28156, 98.36116, 99.09475, 99.45793,
    99.47634, 99.56519, 99.59689, 99.80310, 99.96620, 100.06790, 100.66260, 100.81482, 101.97113, 101.97667, 102.93476,
    103.21854, 103.23386, 103.77116, 103.89888, 103.91929, 104.34439, 104.74232, 105.31758, 105.32821, 105.60509, 105.84881,
    106.14772, 106.54333, 106.59670, 106.69538, 106.79779, 106.84138, 107.19005, 107.61806, 107.85569, 107.96401, 108.27283,
    109.18719, 109.34009, 109.99233, 110.18195, 112.10702, 112.13682, 112.30508, 112.43540, 112.72758, 114.60769, 114.67982,
    115.08818, 115.35477, 119.49361, 119.61944, 120.02999, 120.53776, 120.61930, 120.66959, 127.04795]

let gamma = [0.406097253, 1.927136715, 0.639342131, 3.336699138, 1.093880030, 2.700616161, 0.001766422, 0.206998554, 0.741494558, 1.299483392, 0.430274110
    , 6.577475983, 1.849348710, 2.952894997, 1.669499855, 0.084587693, 0.044020240, 0.263098605, 0.897018817, 1.444843507, 0.571559751, 0.022303623
    , 0.577537269, 0.081302691, 3.848764057, 2.353167018, 0.268868977, 4.917589395, 5.917678560, 0.476190803, 2.116104897, 0.271883166, 6.748716188
    , 3.892478150, 0.379680808, 0.638750983, 0.429443774, 1.402422600, 2.583584083, 1.201288079, 7.326501270, 2.799731953, 2.377142473, 1.149476840
    , 0.231626102, 1.921274065, 0.598562140, 0.239498301, 0.653548973, 0.556557971, 0.768732829, 0.006114799, 2.619890722, 0.129391948, 1.929628875
    , 0.207203269, 2.419582215, 0.624902881, 1.577329301, 0.321728885, 1.251751451, 1.853064964, 6.198788198, 3.536307202, 1.879690833, 0.171038473
    , 1.305210881, 1.572385810, 2.961620531, 4.628528289, 2.302651890, 7.267263659, 2.685723730, 1.377514846, 3.700843146, 3.813869127, 0.208235206
    , 2.690799919, 4.856524489, 1.060161667, 0.154099355, 0.724712726, 3.173177290, 3.473125648, 5.631928889, 7.268106849, 0.778306458, 4.512081244
    , 0.014876456, 0.803038472, 5.894150192, 1.528405472, 2.461929148, 4.952444982, 0.035199990, 3.395572744, 2.567366288, 0.789625627, 5.653114144
    , 0.151561768];

const alp = [
    91, 96, 65, 225, 69, 62, 86, 84, 68, 73, 77, 85, 105, 51, 66, 80, 35, 52, 135, 225, 41, 38, 70, 82, 52, 48, 90, 81, 58, 86, 89, 100,
    114, 113, 90, 95, 66, 57, 103, 113, 90, 135, 75, 35, 114, 78, 66, 57, 105, 60, 85, 105, 67, 59, 112, 50, 69, 74, 71, 95, 106, 92,
    78, 146, 144, 82, 144, 170, 58, 84, 170, 84, 95, 82, 82, 67, 71, 77, 79, 181, 58, 86, 77, 258, 77, 49, 73, 68, 86, 65, 78, 95, 69,
    67, 56, 63, 94, 87, 35, 100
]

test("Robust", () => {
    const res = robust(data);
    expect(res[0]).toBeCloseTo(79.80121, 5);
    expect(res[1]).toBeCloseTo(121.06286, 5);
});

test("Get bootstrap sample", () => {
    const bootstrapN = 10;
    const collection = []
    const data = [10, 20, 30, 40, 50];
    for (let i = 0; i < bootstrapN; i++) {
        const sample = sample_with_replacement(data);
        collection.push(sample);
    }
    expect(collection[0].length).toBe(data.length);
    expect(collection.length).toBe(bootstrapN);
    //console.log(collection);
});

test("Bootstrap confidence interval", () => {
    const ref = robust(data);
    //console.log(ref);
    const res = bootstrap_confidence_interval(data);
    //console.log(res);
    expect(res[0][0]).toBeCloseTo(77, 0);
    expect(res[0][1]).toBeCloseTo(83, 0);
    expect(res[1][0]).toBeCloseTo(118, 0);
    expect(res[1][1]).toBeCloseTo(124, 0);
});

test("Median absolute deviation", () => {
    const data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    const res = median_absolute_deviation(data, jStat.median(data));
    expect(res).toBeCloseTo(2.5);
});

test("C2", () => {
    expect(c2(0.10)).toBeCloseTo(28.4, 1);
    expect(c2(0.05)).toBeCloseTo(205.4, 1);
});

test("Boxcox fit Shapiro-Wilk", () => {
    const res = boxcoxfit(gamma, [-2, 2, 0.01], "sw");
    //console.log(res);
    expect(res.lambda).toBeCloseTo(0.32);
    expect(res.transformedData.length).toBe(100);
    expect(res.shapiroWilkValue.w).toBeCloseTo(0.9817, 4);
    const mean = res.transformedData.reduce((a, b) => a + b, 0) / res.transformedData.length;
    const res_trans = robust(res.transformedData);
    //console.log(boxcoxinv(res_trans, res.lambda));
});

test("Boxcox fit MLE", () => {
    const res = boxcoxfit(gamma, [-2, 2, 0.01], "mle");
    //console.log(res);
    expect(res.lambda).toBeCloseTo(0.31);
    expect(res.transformedData.length).toBe(100);
});

test("Normal cdf", () => {
    expect(normal.cdf(0, 0, 1)).toBeCloseTo(0.5);
    expect(normal.cdf(1, 0, 1)).toBeCloseTo(0.841344746);
    expect(normal.cdf(-1.96, 0, 1)).toBeCloseTo(0.025, 3);
    expect(normal.cdf(1.96, 0, 1)).toBeCloseTo(0.975, 3);
});

test("Shapiro-Wilk", () => {
    const res = ShapiroWilkW(alp);
    expect(res.w).toBeCloseTo(0.811,3);
    expect(res.p).toBeCloseTo(0.0,3);
    //console.log(res);
});

test("Adjusted Fisher-Pearson coefficient", () => {
    const res = adjusted_fisher_pearson_coefficient(alp);
    console.log(res);
    let isNormal = is_consistent_with_normal(alp.length, res);
    expect(isNormal).toBe(false);
    const fit = boxcoxfit(alp, [-2, 2, 0.01], "sw");
    const res2 = adjusted_fisher_pearson_coefficient(fit.transformedData);
    isNormal = is_consistent_with_normal(fit.transformedData.length, res2);
    expect(isNormal).toBe(true);
    console.log(res2);
});

