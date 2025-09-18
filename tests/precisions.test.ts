/* npm install -D typescript jest ts-jest @types/jest @types/node
 */
import { TwoFactorNestedAnova, TwoFactorVarianceAnalysis, TwoFactorVariance } from '../src/precision';

const days = ["Day 1", "Day 1", "Day 1", "Day 1", "Day 2", "Day 2", "Day 2", "Day 2",
    "Day 3", "Day 3", "Day 3", "Day 3", "Day 4", "Day 4", "Day 4", "Day 4",
    "Day 5", "Day 5", "Day 5", "Day 5", "Day 6", "Day 6", "Day 6", "Day 6",
    "Day 7", "Day 7", "Day 7", "Day 7", "Day 8", "Day 8", "Day 8", "Day 8",
    "Day 9", "Day 9", "Day 9", "Day 9", "Day 10", "Day 10", "Day 10", "Day 10",
    "Day 11", "Day 11", "Day 11", "Day 11", "Day 12", "Day 12", "Day 12", "Day 12",
    "Day 13", "Day 13", "Day 13", "Day 13", "Day 14", "Day 14", "Day 14", "Day 14",
    "Day 15", "Day 15", "Day 15", "Day 15", "Day 16", "Day 16", "Day 16", "Day 16",
    "Day 17", "Day 17", "Day 17", "Day 17", "Day 18", "Day 18", "Day 18", "Day 18",
    "Day 19", "Day 19", "Day 19", "Day 19", "Day 20", "Day 20", "Day 20", "Day 20"];

const runs = ["Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2",
    "Run 1", "Run 1", "Run 2", "Run 2", "Run 1", "Run 1", "Run 2", "Run 2"];

const values = [105.9125495, 103.4594439, 100.3746204, 103.7659139, 101.5233903, 101.3347999, 102.7242272, 99.32447708,
    104.6393885, 94.140947, 95.33198048, 98.00188127, 104.5948706, 98.5307138, 96.03949179, 93.73456291,
    103.8169936, 100.9089366, 95.24712655, 103.279433, 98.57880475, 100.6577332, 104.5750405, 102.1076445,
    98.11422559, 104.7763511, 102.4655525, 99.93848875, 103.0618642, 102.9668452, 90.75061241, 93.40087747,
    100.5039457, 99.86293922, 98.93478169, 102.1839088, 95.95539457, 108.4717926, 99.49379525, 95.7559523,
    106.643512, 98.85220955, 99.49343338, 98.63464539, 99.26235495, 99.05364823, 106.0952489, 105.2823297,
    100.6536363, 105.123515, 94.8370386, 97.99274921, 99.00280737, 103.7450497, 103.0861776, 104.1824404,
    100.3202806, 100.6459121, 100.1346292, 93.84683307, 99.54365554, 95.55137483, 98.56707898, 94.78143328,
    102.014383, 100.126421, 102.443221, 97.48245746, 103.3301168, 100.1661557, 102.3016949, 96.31915012,
    96.54984086, 99.87260073, 102.8303454, 100.0121366, 106.0571248, 104.1900122, 100.0701146, 102.8510929];

const ep15_a1 = [242,246,245,246,243,242,238,238,247,239,241,240,249,241,250,245,246,242,243,240,
    244,245,251,247,241,246,245,247,245,245,243,245,243,239,244,245,244,246,247,239,
    252,251,247,241,249,248,251,246,242,240,251,245,246,249,248,240,247,248,245,246,
    240,238,239,242,241,244,245,248,244,244,237,242,241,239,247,245,247,240,245,242];

test('TwoFactorNestedAnova with balanced data', () => {
    const twoFactorNestedAnova = new TwoFactorNestedAnova(days, runs, values);
    const anova = twoFactorNestedAnova.calculate();
    //console.log(anova);
    expect(anova.mean).toBeCloseTo(100.3899, 4);
    expect(anova.ssa).toBeCloseTo(248.5388, 4);
    expect(anova.sse).toBeCloseTo(400.5985, 4);
    expect(anova.ssb).toBeCloseTo(370.4565, 4);
    expect(anova.dfT).toBe(79);
    expect(anova.dfA).toBe(19);
    expect(anova.dfE).toBe(40);
    expect(anova.dfB).toBe(20);
});

test('TwoFactorNestedAnova with unbalanced data', () => {
    const days2 = [...days]; //clone the array
    days2.splice(39,1);
    days2.splice(6,2);
    const runs2 = [...runs]; //clone the array
    runs2.splice(39,1);
    runs2.splice(6,2);
    const values2 = [...values];
    values2.splice(39,1);
    values2.splice(6,2);
    const twoFactorNestedAnova = new TwoFactorNestedAnova(days2, runs2, values2);
    expect(() => twoFactorNestedAnova.calculate()).toThrow(Error);
});

function round4(x: number): number {
    return Math.round((x + Number.EPSILON) * 10000) / 10000;
}

test('TwoFactorVarianceAnalysis with balanced data', () => {
    const twoFactorVarianceAnalysis = new TwoFactorVarianceAnalysis(days, runs, values, 0.05);
    const variance = twoFactorVarianceAnalysis.calculate();
    //console.log(variance);
    expect(variance.vT).toBeCloseTo(14.2689, 4);
    expect(variance.vA).toBeCloseTo(0.0, 4);
    expect(variance.vB).toBeCloseTo(4.2539, 4);
    expect(variance.vE).toBeCloseTo(10.0150, 4);
    expect(variance.dfA).toBe(19);
    expect(variance.dfB).toBe(20);
    expect(variance.dfE).toBe(40);
});

test('EP05 Table A1 with balanced data', () => {
    const twoFactorVarianceAnalysis = new TwoFactorVarianceAnalysis(days, runs, ep15_a1, 0.05);
    const variance = twoFactorVarianceAnalysis.calculate();
    //console.log(variance);
    expect(variance.vT).toBeCloseTo(12.934, 3);
    expect(variance.vA).toBeCloseTo(1.959, 3);
    expect(variance.vB).toBeCloseTo(3.075, 3);
    expect(variance.vE).toBeCloseTo(7.90, 2);
    expect(variance.dfA).toBe(19);
    expect(variance.dfB).toBe(20);
    expect(variance.dfE).toBe(40);
    expect(Number(variance.dfWL.toFixed(1))).toBeCloseTo(64.8);
});

test('TwoFactorVarianceAnalysis with unbalanced data', () => {
    const days2 = [...days]; //clone the array
    days2.splice(39,1);
    days2.splice(6,2);
    const runs2 = [...runs]; //clone the array
    runs2.splice(39,1);
    runs2.splice(6,2);
    const values2 = [...values];
    values2.splice(39,1);
    values2.splice(6,2);
    const twoFactorVarianceAnalysis = new TwoFactorVarianceAnalysis(days2, runs2, values2, 0.05);
    expect(() => twoFactorVarianceAnalysis.calculate()).toThrow(Error);
});
