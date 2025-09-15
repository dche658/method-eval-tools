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

const { mean, chisquare } = require('jstat-esm');

/* Class to perform one factor anova
 * Algorithm from Mendenhall WM, Sincich TL. 2016. Statistics for Engineering
 * and the Sciences 6th ed. CRC Press, Boca Raton. p752.
 */
export class OneFactorAnova {
  private dict: { [key: string]: number[]; };
  private factorA: (string | number)[];
  private values: number[];
  constructor(factorA: (string | number)[], values: number[]) {
    this.dict = {};
    this.factorA = factorA;
    this.values = values;

    if (factorA.length !== values.length) {
      throw new Error("Factor A and values must have the same length.");
    }
    //Add values to dictionary by factor
    for (let i = 0; i < factorA.length; i++) {
      if (!(factorA[i] in this.dict)) {
        this.dict[factorA[i]] = new Array<number>();
      }
      this.dict[factorA[i]].push(values[i]);
    }
  }

  calculate(): {
    mean: number,
    ssTotal: number,
    sst: number,
    sse: number,
    dfTotal: number,
    dfE: number,
    dfT: number,
    mst: number,
    mse: number,
    f: number,
    n: number,
    p: number,
    sn2: number,
  } {
    const numGroups = Object.keys(this.dict).length;
    const totalObservations = this.values.length;
    const groupSizes = [];
    const groupSums = [];
    for (let key in this.dict) {
      groupSizes.push(this.dict[key].length);
      groupSums.push(this.dict[key].reduce((a, b) => a + b, 0));
    }

    const sumGroupSizesSquared = groupSizes.reduce((a, b) => a + Math.pow(b, 2), 0);

    const sumAllValues = this.values.reduce((a, b) => a + b, 0);
    const grandMean = mean(this.values);
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
      sn2: sumGroupSizesSquared,
    };
  }
} //End OneFactorAnova

export interface OneFactorVariance {
  mean: number,
  ssTotal: number,
  sst: number,
  sse: number,
  dfTotal: number,
  dfE: number,
  dfT: number,
  mst: number,
  mse: number,
  f: number,
  n: number,
  p: number,
  sn2: number,
  vE: number,
  vB: number,
  sE: number,
  cvE: number,
  sB: number,
  cvB: number,
  s_WL: number,
  cv_WL: number,
  dfWL: number,
  chisqRepeatability: number,
  chisqWL: number,
  fRepeatability: number,
  fWL: number
}

export class OneFactorVarianceAnalysis {
  private factorA: (string | number)[];
  private values: number[];
  private alpha: number;
  private numLevels: number;
  private fRepeatability: number;
  private fWL: number;
  private isCalculated: boolean;

  constructor(factorA: (string | number)[], values: number[], numLevels: number, alpha = 0.05) {
    this.factorA = factorA;
    this.values = values;
    this.alpha = alpha;
    this.numLevels = numLevels;
    this.fRepeatability = 0;
    this.fWL = 0;
    this.isCalculated = false;
  }

  calculate(): OneFactorVariance {
    const oneFactorAnova = new OneFactorAnova(this.factorA, this.values);
    const anova = oneFactorAnova.calculate();
    const vE = anova.mse;
    const sE = Math.sqrt(vE);
    const cvE = sE / anova.mean;
    const n0 = (anova.n - anova.sn2 / anova.n) / (anova.p - 1);
    const vB = Math.max(0, (anova.mst - anova.mse) / n0);
    const sB = Math.sqrt(vB);
    const cvB = sB / anova.mean;
    const s_WL = Math.sqrt(vE + vB);
    const cv_WL = s_WL / anova.mean;
    const alpha1 = 1 / n0; // 1 / anova.p;
    const alpha2 = 1 - alpha1; //(n0-1)/n0;
    const numerator = Math.pow(alpha1 * anova.mst + alpha2 * anova.mse, 2);
    const denominator1 = Math.pow(alpha1 * anova.mst, 2) / anova.dfT;
    const denominator2 = Math.pow(alpha2 * anova.mse, 2) / anova.dfE;
    const dfWL = numerator / (denominator1 + denominator2);

    //One tail test with Bonferroni correction for the number of levels. CLSI EP05 approach.
    //Analyse It just calculates for each level
    const chisqRepeatability = chisquare.inv(1 - this.alpha / this.numLevels, anova.dfE);
    const chisqWL = chisquare.inv(1 - this.alpha / this.numLevels, dfWL);
    this.fRepeatability = chisqRepeatability / anova.dfE;
    this.fWL = chisqWL / dfWL;
    this.isCalculated = true;
    return {
      ...anova,
      ...{
        vE: vE,
        vB: vB,
        sE: sE,
        cvE: cvE,
        sB: sB,
        cvB: cvB,
        s_WL: s_WL,
        cv_WL: cv_WL,
        dfWL: dfWL,
        chisqRepeatability: chisqRepeatability,
        chisqWL: chisqWL,
        fRepeatability: this.fRepeatability,
        fWL: this.fWL,
      },
    };
  }

  /* Get upper verification limit for repeatability
   */
  getUVLRepeatability(cvClaim: number): number {
    if (this.isCalculated === false) {
      this.calculate();
    }
    return cvClaim * this.fRepeatability;
  }

  /* Get upper verification limit for within lab imprecision
   */
  getUVLWL(cvClaim: number): number {
    if (this.isCalculated === false) {
      this.calculate();
    }
    return cvClaim * this.fWL;
  }
} //OneFactorVarianceAnalysis

/* Class to perform two factor anova
 * For the CLSI EP protocol of
 * n replicates per run x r runs per day x d days
 */
export class TwoFactorAnova {
  private factorA: (string | number)[];
  private factorB: (string | number)[];
  private values: number[];
  private dict: { [key: string]: { [key: string]: number[] } };
  private numFactorA: number;
  private numFactorB: number;
  private numReps: number;
  private dictB: { [key: string]: number[] };
  private grandMean: number;

  constructor(factorA: (string | number)[], factorB: (string | number)[], values: number[]) {
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
    this.grandMean = mean(values);
  }

  /* Total Sum of Squares (SST):
   * Measures the total variability in the data. Calculated by summing the squared
   * differences between each data point and the grand mean.
   */
  calculateSST(): number {
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
  calculateSSB(): number {
    let ssb = 0;
    for (let key in this.dictB) {
      let group = this.dictB[key]; //get factor A group: an object

      ssb += group.length * Math.pow(mean(Object.values(group)) - this.grandMean, 2);
    }
    return ssb;
  }

  /* Sum of squares for factor A
   * Calculated by summing the squared differences between the means of each
   * level of Factor A and the grand mean, weighted by the number of observations
   * in each level.
   */
  calculateSSA(): number {
    let ssa = 0;
    for (let key in this.dict) {
      let group = this.dict[key]; //get factor A group
      let groupMean = mean(Object.values(group).flat());
      let groupN = Object.values(group).flat().length;
      ssa += groupN * Math.pow(groupMean - this.grandMean, 2);
    }
    return ssa;
  }

  /* Sum of squares of the interaction between factor A and factor B
   */
  calculateSSAB(sst: number, ssa: number, ssb: number, sse: number): number {
    return sst - ssa - ssb - sse;
  }

  /* Sum of squares for the error components
   * For each cell corresponding to factorA and factorB, calculate the sum of the squared difference
   * between each value in the cell and the mean for that cell.
   */
  calculateSSE(): number {
    let sse = 0;
    for (let key in this.dict) {
      let group = this.dict[key]; //get factor A group
      for (let key2 in group) {
        let groupB = group[key2]; //get factor B group
        for (let i = 0; i < groupB.length; i++) {
          sse += Math.pow(groupB[i] - mean(Object.values(groupB).flat()), 2);
        }
      }
    }
    return sse;
  }

  calculateDfT(): number {
    return this.values.length - 1;
  }

  calculateDfA(): number {
    return this.numFactorA - 1;
  }

  calculateDfB(): number {
    return this.numFactorB - 1;
  }
  calculateDfAB(): number {
    return (this.numFactorA - 1) * (this.numFactorB - 1);
  }
  calculateDfE(): number {
    return this.values.length - this.numFactorA * this.numFactorB;
  }

  calculate(): {
    mean: number,
    sst: number,
    ssa: number,
    ssb: number,
    sse: number,
    ssab: number,
    dfT: number,
    dfA: number,
    dfB: number,
    dfAB: number,
    dfE: number,
    msa: number,
    msb: number,
    msab: number,
    mse: number,
    fAB: number,
    fA: number,
    fB: number,
    n: number,
    nA: number,
    nB: number,
    nE: number,
  } {
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
      n: this.values.length,
      nA: this.numFactorA,
      nB: this.numFactorB,
      nE: this.numReps,
    };
  }
} //End TwoFactorAnova

export interface TwoFactorVariance {
  mean: number,
  sst: number,
  ssa: number,
  ssb: number,
  sse: number,
  ssab: number,
  dfT: number,
  dfA: number,
  dfB: number,
  dfAB: number,
  dfE: number,
  msa: number,
  msb: number,
  msab: number,
  mse: number,
  fAB: number,
  fA: number,
  fB: number,
  n: number,
  nA: number,
  nB: number,
  nE: number,
  vA: number,
  vAB: number,
  vE: number,
  sA: number,
  sAB: number,
  sE: number,
  sWL: number,
  cvA: number,
  cvAB: number,
  cvE: number,
  cvT: number,
  dfWL: number,
  sWL_LCL: number,
  sWL_UCL: number,
  sE_LCL: number,
  sE_UCL: number,
  cvWL_LCL: number,
  cvWL_UCL: number,
  cvE_LCL: number,
  cvE_UCL: number,
}

export class TwoFactorVarianceAnalysis {
  private factorA: (string | number)[];
  private factorB: (string | number)[];
  private values: number[];
  private numLevels: number;
  private alpha: number;
  private fRepeatability: number;
  private fWL: number;
  private isCalculated: boolean;

  /*
   * factorA as an Array (Days)
   * factorB as an Array (Runs)
   * values as an Array
   * alpha for calculation of confidence limits. Default is 0.05
   */
  constructor(factorA: (string | number)[], factorB: (string | number)[], values: number[], numLevels: number, alpha = 0.05) {
    this.factorA = factorA;
    this.factorB = factorB;
    this.values = values;
    this.numLevels = numLevels;
    this.alpha = alpha;
    this.fRepeatability = 0;
    this.fWL = 0;
    this.isCalculated = false;
  }

  calculate(): TwoFactorVariance {
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
      ...anova,
      ...{
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
      },
    };
  }

  /* Confidence Limits
   *
   * jStat.chisquare.inv(1-this.alpha/numLevels, dfWL);
   * MS Excel produces different results because it floors the decimal to an integer
   * and dfWL of 9.8 is rounded down to 9
   */
  getUpperConfidenceLimit(sd: number, df: number, alpha: number) {
    let ucl = sd * Math.sqrt(df / chisquare.inv(alpha / 2, df));
    return ucl;
  }
  getLowerConfidenceLimit(sd: number, df: number, alpha: number) {
    let lcl = sd * Math.sqrt(df / chisquare.inv(1 - alpha / 2, df));
    return lcl;
  }
} //TwoFactorVarianceAnalysis
