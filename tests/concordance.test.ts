import { ContingencyTableBuilder, ConcordanceCalculator } from "../src/concordance";

const x = [10377.5, 4056, 2654, 4747, 1459.5, 5880, 3871, 2461, 1802, 1607.5,
    4329, 7911.5, 1798.5, 6504, 9506.5, 4781, 2122, 17859, 7064, 963.5,
    408.5, 5623, 4923.5, 2745, 1738.5, 7057, 3001.5, 895.5, 12155.5, 1290.5,
    943, 11312.5, 2700, 3561, 1711, 1006, 1180.5, 9007, 4438, 3606];

const y = [10741.5, 3848.5, 2655.5, 5190.5, 1554.5, 6025, 3861.5, 2533.5, 1858, 1636, 
    4341.5, 8185, 1753.5, 6793, 9056, 4724.5, 2133.5, 16576.5, 7106.5, 986.5, 
    481, 5603, 5046, 2799.5, 1721.5, 7547.5, 3009.5, 986.5, 11712, 1416, 
    998.5, 11698.5, 2829, 3776.5, 1899, 1042, 1616, 9380.5, 4332.5, 3308.5];

const thresholds = [
    [2000, 2010],
    [4000, 4010],
    [6000, 6010],
]

test('ContingencyTableBuilder', () => {
    const builder = new ContingencyTableBuilder(thresholds);
    const table = builder.build(x, y);
    expect(table.values[0][0]).toBe(13);
    expect(table.values[0][1]).toBe(0);
    expect(table.values[0][2]).toBe(0);
    expect(table.values[0][3]).toBe(0);
    expect(table.values[1][1]).toBe(9);
    expect(table.values[2][1]).toBe(1);
    expect(table.values[2][2]).toBe(6);
    expect(table.values[2][3]).toBe(1);
    expect(table.values[3][3]).toBe(10);
});

test('ConcordanceCalculator', () => {
    const builder = new ContingencyTableBuilder(thresholds);
    const table = builder.build(x, y);
    const calculator = new ConcordanceCalculator(table);
    const concordance = calculator.calculate();
    console.log(concordance);
    expect(concordance.overallAgreement).toBeCloseTo(0.95,2);
    expect(concordance.expectedAgreement).toBeCloseTo(0.2606,4);
    expect(concordance.kappa).toBeCloseTo(0.9324, 4);
    expect(concordance.seKappa).toBeCloseTo(0.0462, 4);
});
