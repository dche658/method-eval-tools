/* Statistical procedures for evaluating imprecision
 *
 * Two classes have been implemented to perform a either a one way or
 * two way ANOVA. These classes are in turn used by the OneFactorVarianceAnalysis
 * and TwoFactorVarianceAnalysis classes to perform variance component
 * analysis. However, the ANOVA classes can be used independently if needed.
 * 
 * The class OneFactorVarianceAnalysis is for a one factor design typically used for
 * verification studies, and calculates an upper verification limit.
 * The class TwoFactorVarianceAnalysis is for a two factor design typically used for
 * validation studies and calculates confidence limits instead of a verification
 * limit.
 */

//import { average } from "./regression.js";
import jStat from "./jstat-1.9.6.min.js";

/* Class to perform one factor anova
 * Algorithm from Mendenhall WM, Sincich TL. 2016. Statistics for Engineering
 * and the Sciences 6th ed. CRC Press, Boca Raton. p752.
 */
export class OneFactorAnova {
  constructor(factorA, values) {
    this.dict = {};
    this.factorA = factorA;
    this.values = values;

    if (factorA.length !== values.length) {
      throw new Error("Factor A and values must have the same length.");
    }
    //Add values to dictionary by factor
    for (let i = 0; i < factorA.length; i++) {
      if (!(factorA[i] in this.dict)) {
        this.dict[factorA[i]] = [];
      }
      this.dict[factorA[i]].push(values[i]);
    }
  }

  calculate() {
    const numGroups = Object.keys(this.dict).length;
    const totalObservations = values.length;
    const groupSizes = [];
    const groupSums = [];
    for (let key in this.dict) {
      groupSizes.push(this.dict[key].length);
      groupSums.push(this.dict[key].reduce((a, b) => a + b, 0));
    }

    const sumGroupSizesSquared = groupSizes.reduce((a, b) => a + Math.pow(b, 2), 0);

    const sumAllValues = this.values.reduce((a, b) => a + b, 0);
    const grandMean = jStat.mean(values);
    const sumSquaredAllValues = this.values.reduce((a, b) => a + Math.pow(b, 2), 0);
    const correctedMean = Math.pow(sumAllValues, 2) / this.values.length;
    const ssTotal = sumSquaredAllValues - correctedMean;
    let sst = 0;
    for (let i = 0; i < groupSizes.length; i++) {
      sst += Math.pow(groupSums[i], 2) / groupSizes[i];
    }
    sst = sst - correctedMean;
    const sse = ssTotal - sst;
    const mst = sst / (numGroups - 1);
    const mse = sse / (totalObservations - numGroups);
    const f = mst / mse;

    return {
      mean: grandMean,
      ssTotal: ssTotal,
      sst: sst,
      sse: sse,
      dfTotal: totalObservations - 1,
      dfE: totalObservations - numGroups,
      dfT: numGroups - 1,
      mst: mst,
      mse: mse,
      f: f,
      n: totalObservations,
      p: numGroups,
      ssGroupN: sumGroupSizesSquared,
    };
  }
} //End OneFactorAnova

export class OneFactorVarianceAnalysis {
  constructor(factorA, values, numLevels, alpha = 0.05) {
    this.factorA = factorA;
    this.values = values;
    this.alpha = alpha;
    this.numLevels = numLevels;
    this.fRepeatability = 0;
    this.fWL = 0;
    this.isCalculated = false;
  }

  calculate() {
    const oneFactorAnova = new OneFactorAnova(this.factorA, this.values);
    const anova = oneFactorAnova.calculate();
    const vE = anova.mse;
    const sE = Math.sqrt(vE);
    const cvE = sE / anova.mean;
    const vB = Math.max(0, (anova.mst - anova.mse) / anova.p);
    const sB = Math.sqrt(vB);
    const cvB = sB / anova.mean;
    const s_WL = Math.sqrt(vE + vB);
    const cv_WL = s_WL / anova.mean;
    const alpha1 = 1 / anova.p;
    const alpha2 = (anova.p - 1) / anova.p;
    const numerator = Math.pow(alpha1 * anova.mst + alpha2 * anova.mse, 2);
    const denominator1 = Math.pow(alpha1 * anova.mst, 2) / anova.dfT;
    const denominator2 = Math.pow(alpha2 * anova.mse, 2) / anova.dfE;
    const dfWL = numerator / (denominator1 + denominator2);

    //One tail test with Bonferroni correction for the number of levels. CLSI EP05 approach.
    //Analyse It just calculates for each level
    const chisqRepeatability = jStat.chisquare.inv(1 - this.alpha / this.numLevels, anova.dfE);
    const chisqWL = jStat.chisquare.inv(1 - this.alpha / this.numLevels, dfWL);
    this.fRepeatability = chisqRepeatability / anova.dfE;
    this.fWL = chisqWL / dfWL;
    this.isCalculated = true;
    return {
      anova: anova,
      vE: vE,
      vB: vB,
      sE: sE,
      cvE: cvE,
      sB: sB,
      cvB: cvB,
      s_WL: s_WL,
      cv_WL: cv_WL,
      chisqRepeatability: chisqRepeatability,
      chisqWL: chisqWL,
      fRepeatability: this.fRepeatability,
      fWL: this.fWL,
    };
  }

  /* Get upper verification limit for repeatability
   */
  getUVLRepeatability(cvClaim) {
    if (this.isCalculated === false) {
      this.calculate();
    }
    return cvClaim * this.fRepeatability;
  }

  /* Get upper verification limit for within lab imprecision
   */
  getUVLWL(cvClaim) {
    if (this.isCalculated === false) {
      this.calculate();
    }
    return cvClaim * this.fWL;
  }
}

/* Class to perform two factor anova
 * For the CLSI EP protocol of
 * n replicates per run x r runs per day x d days
 */
export class TwoFactorAnova {
  constructor(factorA, factorB, values) {
    this.factorA = factorA;
    this.factorB = factorB;
    this.values = values;
    if (factorA.length !== factorB.length || factorA.length !== values.length) {
      throw new Error("Factor A, B and values must have the same length.");
    }
    this.dict = {}; //dictionary of factor A nested with factor B
    this.numFactorA = 0; //number of days
    this.numFactorB = 0; //number of runs
    this.numReps = 0; //number of replicates per run
    this.dictB = {}; //dictionary of factor B

    //Add values to dictionary by factor
    for (let i = 0; i < factorA.length; i++) {
      if (!(factorA[i] in this.dict)) {
        this.dict[factorA[i]] = {};
        this.numFactorA++;
      }
      let group = this.dict[factorA[i]];
      if (!(factorB[i] in group)) {
        group[factorB[i]] = [];
      }
      // Add values for factor A: factor B combinations
      group[factorB[i]].push(values[i]);
      if (!(factorB[i] in this.dictB)) {
        this.dictB[factorB[i]] = [];
      }
      // Add values for factor B
      this.dictB[factorB[i]].push(values[i]);
    }
    this.numFactorB = 0; //Number of runs r per day
    for (let key in this.dict) {
      let factorB = this.dict[key];
      this.numFactorB = Math.max(this.numFactorB, Object.keys(factorB).length);
      for (let key2 in factorB) {
        this.numReps = Math.max(this.numReps, factorB[key2].length);
      }
    }
    this.grandMean = jStat.mean(values);
  }

  /* Total Sum of Squares (SST):
   * Measures the total variability in the data. Calculated by summing the squared
   * differences between each data point and the grand mean.
   */
  calculateSST() {
    let sst = 0;
    for (let i = 0; i < this.values.length; i++) {
      sst += Math.pow(this.values[i] - this.grandMean, 2);
    }
    return sst;
  }

  /* Sum of squares for factor B
   * Measures the variability due to Factor B. Calculated by summing the squared
   * differences between the means of each level of Factor B and the grand mean,
   * weighted by the number of observations in each level.
   */
  calculateSSB() {
    let ssb = 0;
    for (let key in this.dictB) {
      let group = this.dictB[key]; //get factor A group: an object

      ssb += group.length * Math.pow(jStat.mean(Object.values(group)) - this.grandMean, 2);
    }
    return ssb;
  }

  /* Sum of squares for factor A
   * Calculated by summing the squared differences between the means of each
   * level of Factor A and the grand mean, weighted by the number of observations
   * in each level.
   */
  calculateSSA() {
    let ssa = 0;
    for (let key in this.dict) {
      let group = this.dict[key]; //get factor A group
      let groupMean = jStat.mean(Object.values(group).flat());
      let groupN = Object.values(group).flat().length;
      ssa += groupN * Math.pow(groupMean - this.grandMean, 2);
    }
    return ssa;
  }

  /* Sum of squares of the interaction between factor A and factor B
   */
  calculateSSAB(sst, ssa, ssb, sse) {
    return sst - ssa - ssb - sse;
  }

  /* Sum of squares for the error components
   * For each cell corresponding to factorA and factorB, calculate the sum of the squared difference
   * between each value in the cell and the mean for that cell.
   */
  calculateSSE() {
    let sse = 0;
    for (let key in this.dict) {
      let group = this.dict[key]; //get factor A group
      for (let key2 in group) {
        let groupB = group[key2]; //get factor B group
        for (let i = 0; i < groupB.length; i++) {
          sse += Math.pow(groupB[i] - jStat.mean(Object.values(groupB).flat()), 2);
        }
      }
    }
    return sse;
  }

  calculateDfT() {
    return this.values.length - 1;
  }

  calculateDfA() {
    return this.numFactorA - 1;
  }

  calculateDfB() {
    return this.numFactorB - 1;
  }
  calculateDfAB() {
    return (this.numFactorA - 1) * (this.numFactorB - 1);
  }
  calculateDfE() {
    return this.values.length - this.numFactorA * this.numFactorB;
  }

  calculate() {
    let sst = this.calculateSST();

    let ssa = this.calculateSSA();
    let ssb = this.calculateSSB();
    let sse = this.calculateSSE();
    let ssab = this.calculateSSAB(sst, ssa, ssb, sse);
    let dfT = this.calculateDfT();
    let dfA = this.calculateDfA();
    let dfB = this.calculateDfB();
    let dfAB = this.calculateDfAB();
    let dfE = this.calculateDfE();

    let msa = ssa / dfA;
    let msb = ssb / dfB;
    let msab = ssab / dfAB;
    let mse = sse / dfE;
    let fAB = msab / mse;
    let fA = msa / mse;
    let fB = msb / mse;

    return {
      mean: this.grandMean,
      sst: sst,
      ssa: ssa,
      ssb: ssb,
      sse: sse,
      ssab: ssab,
      dfT: dfT,
      dfA: dfA,
      dfB: dfB,
      dfAB: dfAB,
      dfE: dfE,
      msa: msa,
      msb: msb,
      msab: msab,
      mse: mse,
      fAB: fAB,
      fA: fA,
      fB: fB,
      n: this.totalObservations,
      nA: this.numFactorA,
      nB: this.numFactorB,
      nE: this.numReps,
    };
  }
} //End TwoFactorAnova

export class TwoFactorVarianceAnalysis {
  /*
   * factorA as an Array (Days)
   * factorB as an Array (Runs)
   * values as an Array
   * alpha for calculation of confidence limits. Default is 0.05
   */
  constructor(factorA, factorB, values, numLevels, alpha = 0.05) {
    this.factorA = factorA;
    this.factorB = factorB;
    this.values = values;
    this.numLevels = numLevels;
    this.alpha = alpha;
    this.fRepeatability = 0;
    this.fWL = 0;
    this.isCalculated = false;
  }

  calculate() {
    const twoFactorAnova = new TwoFactorAnova(this.factorA, this.factorB, this.values);
    const anova = twoFactorAnova.calculate();

    //variance for factor A: Between day
    const vA = Math.max(0, (anova.msa - anova.msab) / (anova.nB * anova.nE));

    //variance for error: within run repeatability
    const vE = anova.mse;

    //variance of interaction: between run variance
    const vAB = Math.max(0, (anova.msab - anova.mse) / anova.nE);

    //between day SD
    const sA = Math.sqrt(vA);

    //between run SD
    const sAB = Math.sqrt(vAB);

    //repeatability: within run SD
    const sE = Math.sqrt(vE);

    //within lab SD
    const sT = Math.sqrt(Math.pow(sA, 2) + Math.pow(sAB, 2) + Math.pow(sE, 2));

    //between day CV
    const cvA = sA / anova.mean;

    //between run CV
    const cvAB = sAB / anova.mean;

    //repeatability: within run CV
    const cvE = sE / anova.mean;

    //within lab CV
    const cvT = sT / anova.mean;

    //Calculations for Satterswaithe estimate of the degrees of freedom
    const alphaA = 1 / (anova.nB * anova.nE);
    const alphaAB = (anova.nE - 1) / (anova.nE * anova.nB);
    const alphaE = (anova.nE - 1) / anova.nE;
    const numerator = Math.pow(alphaA * anova.msa + alphaAB * anova.msab + alphaE * anova.mse, 2);
    const denominator1 = Math.pow(alphaA * anova.msa, 2) / anova.dfA;
    const denominator2 = Math.pow(alphaAB * anova.msab, 2) / anova.dfAB;
    const denominator3 = Math.pow(alphaE * anova.mse, 2) / anova.dfE;
    const dfWL = numerator / (denominator1 + denominator2 + denominator3);
    const sWL_UCL = this.getUpperConfidenceLimit(sT, dfWL, this.alpha);
    const sWL_LCL = this.getLowerConfidenceLimit(sT, dfWL, this.alpha);
    const sE_UCL = this.getUpperConfidenceLimit(sE, anova.dfE, this.alpha);
    const sE_LCL = this.getLowerConfidenceLimit(sE, anova.dfE, this.alpha);
    const cvWL_UCL = sWL_UCL / anova.mean;
    const cvWL_LCL = sWL_LCL / anova.mean;
    const cvE_UCL = sE_UCL / anova.mean;
    const cvE_LCL = sE_LCL / anova.mean;

    this.isCalculated = true;

    return {
      anova: anova,
      vA: vA,
      vAB: vAB,
      vE: vE,
      sA: sA,
      sAB: sAB,
      sE: sE,
      sWL: sT,
      cvA: cvA,
      cvAB: cvAB,
      cvE: cvE,
      cvT: cvT,
      dfWL: dfWL,
      sWL_LCL: sWL_LCL,
      sWL_UCL: sWL_UCL,
      sE_LCL: sE_LCL,
      sE_UCL: sE_UCL,
      cvWL_LCL: cvWL_LCL,
      cvWL_UCL: cvWL_UCL,
      cvE_LCL: cvE_LCL,
      cvE_UCL: cvE_UCL,
    };
  }

  /* Confidence Limits
   *
   * jStat.chisquare.inv(1-this.alpha/numLevels, dfWL);
   * MS Excel produces different results because it floors the decimal to an integer
   * and dfWL of 9.8 is rounded down to 9
   */
  getUpperConfidenceLimit(sd, df, alpha) {
    return sd * Math.sqrt(df / jStat.chisquare.inv(1 - alpha / 2, df));
  }
  getLowerConfidenceLimit(sd, df, alpha) {
    return sd * Math.sqrt(df / jStat.chisquare.inv(alpha / 2, df));
  }
}
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
