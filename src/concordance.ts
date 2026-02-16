/**
 * Classes to build contingency tables and calculate Cohen's Kappa
 * statistic
 *
 * @author Douglas Chesher
 *
 */

import { normal } from "jstat-esm";

const X_COL = 0; // column index for x data
const Y_COL = 1; // column index for y data

/**
 * Representation of a contingency tables
 * values is a two dimensional array with the summed values
 * xLabels are the row labels found in the left hand column of the table
 * yLabels are the column labels in the top row
 * table is a formatted version of the table
 */
interface ContingencyTable {
  values: number[][];
  xLabels: string[];
  ylabels: string[];
  table: (string | number)[][];
}

/**
 * Class to build the contingency table from data in two numeric arrays
 * and the stated thresholds. The arrays are assumed to represent two
 * columns of results from the measurement of each sample (rows) by two
 * different methods. The number of thresholds must be the same for each
 * method. The thresholds must be in numeric order.
 */
class ContingencyTableBuilder {
  // Thresholds as [n rows][2 columns]
  private thresholds: number[][];

  /**
   * 
   * @param thresholds diagnostic thresholds in an n by 2 array
   */
  constructor(thresholds: number[][]) {
    this.thresholds = thresholds;
  }

  /**
   * Build the contingency table
   * 
   * @param x quantitative results for method 1.
   * @param y quantitative results for method 2.
   * @returns
   */
  build(x: number[], y: number[]): ContingencyTable {
    let values: number[][] = this.initializeValues(this.thresholds.length + 1);
    // tabulate data
    for (let i = 0; i < x.length; i++) {
      let rowIndex = this.getCategoryIndex(X_COL, x[i]);
      let colIndex = this.getCategoryIndex(Y_COL, y[i]);
      values[rowIndex][colIndex]++;
    }
    // build labels
    let xLabels: string[] = this.getLabels(X_COL);
    let yLabels: string[] = this.getLabels(Y_COL);
    // console.log(xLabels);
    // console.log(yLabels);
    // console.log(values);
    // create a formatted table
    const table = formatContingencyTable(values, xLabels, yLabels);
    return {
      values: values,
      xLabels: xLabels,
      ylabels: yLabels,
      table: table,
    };
  }

  /**
   * Build the label strings based on the thresholds
   * Less than lowest threshold: "< t1"
   * Between thresholds ">= t1 - < t2"
   * Greater than highest threshold ">= t2"
   * 
   * @param col Which column of the diagnostic thresholds array to use. Must be 0 or 1.
   * @returns
   * 
   * @internal
   */
  getLabels(col: number): string[] {
    //console.log(this.thresholds);
    let labels: string[] = new Array<string>(this.thresholds.length + 1);
    for (let row = 0; row <= this.thresholds.length; row++) {
      if (row === 0) {
        labels[row] = `<${this.thresholds[row][col].toString()}`;
      } else if (row === this.thresholds.length) {
        labels[row] = `≥${this.thresholds[row - 1][col].toString()}`;
      } else {
        labels[row] =
          `≥${this.thresholds[row - 1][col].toString()}-<${this.thresholds[row][col].toString()}`;
      }
    }
    return labels;
  }

  /**
   * Compare value to the thresholds for the stated column. It assumes the thresholds
   * are in numeric order from smallest to largest.
   *
   * @param col Which column of the diagnostic thresholds array to use. Must be 0 or 1.
   * @param val Value to be evaluated against the diagnostic thresholds
   * @returns The index within the contingency table.
   * 
   * @internal
   */
  getCategoryIndex(col: number, val: number): number {
    let index = 0;
    // assess against each threshold
    for (let row = 0; row < this.thresholds.length; row++) {
      if (row === 0 && val < this.thresholds[row][col]) {
        //value is less than the lowest threshold
        index = 0;
        break;
      } else if (
        row > 0 &&
        val >= this.thresholds[row - 1][col] &&
        val < this.thresholds[row][col]
      ) {
        //value is between two thresholds.
        index = row;
        break;
      } else if (row === this.thresholds.length - 1 && val >= this.thresholds[row][col]) {
        //value is greater than the highest threshold
        index = row + 1;
        break;
      }
    }
    return index;
  }

  /** 
   * Create [t]x[t] matrix and fill with zeros.
   * Where t is the number of thresholds + 1
   * 
   * @param size
   * @returns
   * 
   * @internal
   */
  initializeValues(size: number): number[][] {
    const values: number[][] = new Array<number[]>(size);
    for (let i = 0; i < size; i++) {
      values[i] = new Array<number>(size).fill(0);
    }
    return values;
  }
} //ContingencyTableBuilder

/**
 * Results of the concordance analysis including Cohen's Kappa and
 * the upper and lower confidence limits
 */
interface ConcordanceResults {
  overallAgreement: number;
  expectedAgreement: number;
  kappa: number;
  seKappa: number;
  lcl: number;
  ucl: number;
}

/** 
 * Sum matrix by row or col
 * 
 * @param matrix m by n matrix
 * @param by must be 'row' or 'col'
 * @returns row or column totals as an array.
 * 
 * @internal
 */
function sum(matrix: number[][], by: string): number[] {
  const sums: number[] = new Array(matrix.length).fill(0);
  if (by === "row") {
    for (let i = 0; i < matrix.length; i++) {
      sums[i] = matrix[i].reduce((sum, val) => sum + val, 0);
    }
  } else if (by === "col") {
    for (let i = 0; i < matrix.length; i++) {
      for (let j = 0; j < matrix[i].length; j++) {
        sums[j] += matrix[i][j];
      }
    }
  }
  return sums;
}

/**
 * Create labeled and formatted contingency table for output
 * 
 * @param values the contingency table
 * @param xLabels row labels
 * @param yLabels column labels
 * @returns labeled contingency table for writing to Excel
 */
function formatContingencyTable(
  values: number[][],
  xLabels: string[],
  yLabels: string[]
): (string | number)[][] {
  const rowSums = sum(values, "row");
  const colSums = sum(values, "col");
  let table: (string | number)[][] = new Array<(string | number)[]>(values.length + 2);
  for (let i = 0; i < values.length + 1; i++) {
    table[i] = new Array<string | number>(values[0].length + 2); //create the row
    for (let j = 0; j < values[0].length + 1; j++) {
      if (i === 0 && j === 0) {
        table[i][j] = "";
      } else if (i === 0) {
        table[i][j] = yLabels[j - 1];
      } else if (j === 0) {
        table[i][j] = xLabels[i - 1];
      } else {
        table[i][j] = values[i - 1][j - 1];
      }
    }
  }
  //add last row
  table[table.length - 1] = new Array<string | number>(values[0].length + 2);
  //console.log(table);

  // Add row sums
  table[0][table[0].length - 1] = "Total";
  for (let row = 1; row < table.length - 1; row++) {
    table[row][table[row].length - 1] = rowSums[row - 1];
  }
  // Add column sums
  table[table.length - 1][0] = "Total";
  for (let col = 1; col < table[0].length - 1; col++) {
    table[table.length - 1][col] = colSums[col - 1];
  }
  // Add total
  table[table.length - 1][table[table.length - 1].length - 1] = colSums.reduce(
    (sum, val) => sum + val,
    0
  );
  return table;
}

/**
 * Calculate concordance and Cohen's Kappa. Default alpha for
 * calculating the confidence limits is 5% or 0.05
 */
class ConcordanceCalculator {
  private table: ContingencyTable;
  private alpha: number;

  /**
   * 
   * @param table the contingency table
   * @param alpha level of significance
   */
  constructor(table: ContingencyTable, alpha = 0.05) {
    this.table = table;
    this.alpha = alpha;
  }

  /**
   * Main method for calculation
   * 
   * @returns Concordance and Cohen's Kappa
   */
  calculate(): ConcordanceResults {
    // calculate the grand total. This should be equal to number of
    // rows in the original column data
    const n = this.table.values.reduce(
      (sum, row) => sum + row.reduce((rowSum, val) => rowSum + val, 0),
      0
    );
    if (n === 0) {
      return {
        overallAgreement: 0,
        expectedAgreement: 0,
        kappa: 0,
        seKappa: 0,
        lcl: 0,
        ucl: 0,
      };
    }

    // Calculate proportions
    const p = [...this.table.values]; //clone the array
    // for each row
    for (let i = 0; i < p.length; i++) {
      p[i] = [...this.table.values[i]]; //clone the array
      // for each column
      for (let j = 0; j < p[i].length; j++) {
        // proportion equals value / n
        p[i][j] /= n;
      }
    }
    // console.log(p);
    // console.log(this.table.values);

    // Calculate observed agreement by summing on the diagonal
    let observedAgreement = 0;
    for (let i = 0; i < this.table.values.length; i++) {
      observedAgreement += this.table.values[i][i];
    }
    const overallAgreement = observedAgreement / n;

    // Calculate expected agreement
    let expectedAgreement = 0;
    const rowSums: number[] = new Array(this.table.values.length).fill(0);
    const colSums: number[] = new Array(this.table.values[0].length).fill(0);
    const p_i: number[] = new Array(p.length).fill(0); //row sum for proportions p_{i.}
    const p_j: number[] = new Array(p[0].length).fill(0); //column sum for proportions p_{.i}

    for (let i = 0; i < this.table.values.length; i++) {
      for (let j = 0; j < this.table.values[i].length; j++) {
        //calculate row sums and column sums
        rowSums[i] += this.table.values[i][j];
        colSums[j] += this.table.values[i][j];
        // row and column sums of proportions is needed later
        // to calculate the standard error
        p_i[i] += p[i][j];
        p_j[j] += p[i][j];
      }
    }

    for (let i = 0; i < this.table.values.length; i++) {
      expectedAgreement += (rowSums[i] / n) * (colSums[i] / n);
    }

    const kappa = (overallAgreement - expectedAgreement) / (1 - expectedAgreement);

    // Calculate the standard error
    // Fleiss JL, Cohen J, Everitt BS. Large Sample Standard Errors of Kappa and Weighted Kappa.
    // Psychological Bulletin. 1969; 72(5): 323-327. Eqn 13.
    let sumPij = 0;
    let sumPii = 0;
    for (let i = 0; i < p.length; i++) {
      for (let j = 0; j < p[i].length; j++) {
        if (i !== j) {
          sumPij += p[i][j] * Math.pow(p_i[i] + p_j[j], 2);
        }
      }
      sumPii +=
        p[i][i] * Math.pow(1 - expectedAgreement - (p_i[i] + p_j[i]) * (1 - overallAgreement), 2);
    }

    //
    let sdKappa =
      (1 / Math.pow(1 - expectedAgreement, 2)) *
      Math.sqrt(
        Math.pow(1 - overallAgreement, 2) * sumPij +
          sumPii -
          Math.pow(
            overallAgreement * expectedAgreement - 2 * expectedAgreement + overallAgreement,
            2
          )
      );
    let seKappa = sdKappa / Math.sqrt(n);

    let zCrit = normal.inv(1 - this.alpha / 2, 0, 1); // two tail

    return {
      overallAgreement: overallAgreement,
      expectedAgreement: expectedAgreement,
      kappa: kappa,
      seKappa: seKappa,
      lcl: kappa - zCrit * seKappa,
      ucl: kappa + zCrit * seKappa,
    };
  }

  /**
   * Format concordance results for output
   * 
   * @param results 
   * @returns
   * 
   * @internal
   */
  formatResultsAsArray(results: ConcordanceResults): (string | number)[][] {
    let arr = [
      ["Overall Agreement", results.overallAgreement],
      ["Expected Agreement", results.expectedAgreement],
      ["Kappa", results.kappa],
      ["SE(Kappa)", results.seKappa],
      ["LCL", results.lcl],
      ["UCL", results.ucl],
    ];
    return arr;
  }
}

/**
 * Class to build the contingency table from data in two arrays.
 * The arrays are assumed to represent two columns of results from the
 * qualitative measurement of each sample (rows) by two different
 * methods and producing categorical data. Data may be numeric or string
 * values. If numeric, they are assumed to represent categories and
 * not have any intrinsic order.
 */
class QualitativeContengencyTableBuilder {

  /**
   * Builds the contingency table from the supplied data
   * 
   * @param x qualitative results for reference method
   * @param y qualitative results for test method
   * @returns Contingency Table.
   */
  build(x: (string | number)[], y: (string | number)[]): ContingencyTable {
    const categories = this.getCategories(x, y);
    // initialize contingency table counts to zero
    const values: number[][] = new Array<number[]>(categories.length);
    for (let row = 0; row < categories.length; row++) {
      values[row] = new Array<number>(categories.length);
      for (let col = 0; col < categories.length; col++) {
        values[row][col] = 0;
      }
    }
    // tabulate data
    for (let i = 0; i < x.length; i++) {
      const rowIndex = categories.indexOf(x[i].toString());
      const colIndex = categories.indexOf(y[i].toString());
      values[rowIndex][colIndex]++;
    }

    const table = formatContingencyTable(values, categories, categories);

    return {
      values: values,
      xLabels: categories,
      ylabels: categories,
      table: table,
    };
  }

  /**
   * Get categories from the supplied data by creating a set
   * and return this as an array.
   * 
   * @param x data in long format
   * @param y data in long format
   * @returns array of unique labels
   * 
   * @internal
   */
  getCategories(x: (string | number)[], y: (string | number)[]): string[] {
    const categories = new Set<string>();
    for (let i = 0; i < x.length; i++) {
      categories.add(x[i].toString());
    }
    for (let i = 0; i < y.length; i++) {
      categories.add(y[i].toString());
    }
    return Array.from(categories);
  }
}

// Module exports
export {
  ContingencyTableBuilder,
  QualitativeContengencyTableBuilder,
  ConcordanceCalculator,
};
