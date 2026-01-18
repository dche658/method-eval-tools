import jStat from 'jstat-esm';
import { test, expect} from '@jest/globals';
import { 
    robust, 
    median_absolute_deviation,
    sample_with_replacement,
    bootstrap_confidence_interval,
    c2
} from '../src/reference_intervals';


let data = [74.64731, 77.25439, 81.64253, 82.37364, 84.96763, 85.83717, 86.15544, 86.25899, 86.32812, 88.06514, 88.30495, 89.39985, 89.50754, 90.04025, 90.47054, 90.62720, 90.67660, 91.00039, 91.03980, 91.35050, 92.34147, 92.43501, 92.64297, 92.81908, 92.93594, 93.40938, 93.61053, 93.61496, 93.82136, 94.99198, 95.25909, 95.26355, 95.58339, 95.79880, 96.03610, 96.40177, 97.05810, 97.35304, 97.48031, 97.50916, 97.86968, 97.93678, 98.00245, 98.28156, 98.36116, 99.09475, 99.45793, 99.47634, 99.56519, 99.59689, 99.80310, 99.96620, 100.06790, 100.66260, 100.81482, 101.97113, 101.97667, 102.93476, 103.21854, 103.23386, 103.77116, 103.89888, 103.91929, 104.34439, 104.74232, 105.31758, 105.32821, 105.60509, 105.84881, 106.14772, 106.54333, 106.59670, 106.69538, 106.79779, 106.84138, 107.19005, 107.61806, 107.85569, 107.96401, 108.27283, 109.18719, 109.34009, 109.99233, 110.18195, 112.10702, 112.13682, 112.30508, 112.43540, 112.72758, 114.60769, 114.67982, 115.08818, 115.35477, 119.49361, 119.61944, 120.02999, 120.53776, 120.61930, 120.66959, 127.04795]

test("Robust", () => {
    const res = robust(data);
    expect(res[0]).toBeCloseTo(79.80121,5);
    expect(res[1]).toBeCloseTo(121.06286,5);
});

test("Get bootstrap sample", () => {
    const bootstrapN = 10;
    const collection = []
    const data = [10, 20, 30, 40, 50];
    for (let i = 0; i < bootstrapN; i++) {
        const sample =  sample_with_replacement(data);
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
    expect(res[0][0]).toBeCloseTo(77,0);
    expect(res[0][1]).toBeCloseTo(83,0);
    expect(res[1][0]).toBeCloseTo(118,0);
    expect(res[1][1]).toBeCloseTo(124,0);
});

test("Median absolute deviation", () => {
    const data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    const res = median_absolute_deviation(data, jStat.median(data));
    expect(res).toBeCloseTo(2.5);
});

test("C2", () => {
    expect(c2(0.10)).toBeCloseTo(28.4,1);
    expect(c2(0.05)).toBeCloseTo(205.4,1);
});