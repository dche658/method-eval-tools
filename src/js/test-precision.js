import {
  OneFactorAnova,
  OneFactorVarianceAnalysis,
  TwoFactorAnova,
  TwoFactorVarianceAnalysis,
} from "./precision.js";

const factorA = [
  "Day 1",
  "Day 1",
  "Day 1",
  "Day 1",
  "Day 2",
  "Day 2",
  "Day 2",
  "Day 2",
  "Day 3",
  "Day 3",
  "Day 3",
  "Day 3",
  "Day 4",
  "Day 4",
  "Day 4",
  "Day 4",
];

const factorB = [
  "Run 1",
  "Run 1",
  "Run 2",
  "Run 2",
  "Run 1",
  "Run 1",
  "Run 2",
  "Run 2",
  "Run 1",
  "Run 1",
  "Run 2",
  "Run 2",
  "Run 1",
  "Run 1",
  "Run 2",
  "Run 2",
];

const values = [242, 246, 245, 246, 243, 242, 238, 238, 247, 239, 241, 240, 249, 241, 250, 245];

const oneFactorAnova = new OneFactorAnova(factorA, values);
const ofa = oneFactorAnova.calculate();
console.log(`ANOVA ${JSON.stringify(ofa)}`);

const oneFactorVarianceAnalysis = new OneFactorVarianceAnalysis(factorA, values, 1);
const resOfa = oneFactorVarianceAnalysis.calculate();
console.log(`Variance Analysis ${JSON.stringify(resOfa)}`);

const twoFactorAnova = new TwoFactorAnova(factorA, factorB, values);

console.log(JSON.stringify(twoFactorAnova.calculate()));

const twoFactorVarianceAnalysis = new TwoFactorVarianceAnalysis(factorA, factorB, values, 0.05);
const res = twoFactorVarianceAnalysis.calculate();
console.log(JSON.stringify(res));

let oneFactorVCAExpected = {
  p: 4,
  n: 16,
  mean: 243.25,
  ssTotal: 211,
  sst: 90,
  sse: 121,
  dfTotal: 15,
  dfE: 12,
  dfT: 3,
  mst: 30,
  mse: 10.0833,
  sn2: 64,
  alpha1: 0.25,
  alpha2: 0.75,
  num: 226.8789,
  den1: 18.75,
  den2: 4.766,
  dfWL: 9.6479,
  sE: 3.175,
  sT: 2.231,
  sWL: 3.881,
  fE: 1.3237,
  fWL: 1.3243,
};

//
