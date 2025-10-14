import { normal } from 'jstat-esm';

const X_COL = 0; // column index for x data
const Y_COL = 1; // column index for y data

interface ContingencyTable {
    values: number[][],
    xLabels: string[],
    ylabels: string[],
    table: (string | number)[][]
}

class ContingencyTableBuilder {
    private thresholds: number[][];

    constructor(thresholds: number[][]) {
        this.thresholds = thresholds;
    }

    build(x: number[], y: number[]): ContingencyTable {
        let values: number[][] = this.initializeValues(this.thresholds.length + 1);
        for (let i = 0; i < x.length; i++) {
            let rowIndex = this.getCategoryIndex(X_COL, x[i]);
            let colIndex = this.getCategoryIndex(Y_COL, y[i]);
            values[rowIndex][colIndex]++;
        }
        let xLabels: string[] = this.getLabels(X_COL);
        let yLabels: string[] = this.getLabels(Y_COL);
        // console.log(xLabels);
        // console.log(yLabels);
        // console.log(values);
        const table = formatContingencyTable(values, xLabels, yLabels);
        return {
            values: values,
            xLabels: xLabels,
            ylabels: yLabels,
            table: table,
        }
    }

    getLabels(col: number): string[] {
        let labels: string[] = new Array<string>(this.thresholds.length + 1);
        for (let row = 0; row <= this.thresholds.length; row++) {
            if (row === 0) {
                labels[row] = `<${this.thresholds[row][col].toString()}`;
            } else if (row === this.thresholds.length) {
                labels[row] = `≥${this.thresholds[row - 1][col].toString()}`;
            } else {
                labels[row] = `≥${this.thresholds[row - 1][col].toString()}-<${this.thresholds[row][col].toString()}`
            }
        }
        return labels;
    }

    getCategoryIndex(col: number, val: number): number {
        let index = 0;
        for (let row = 0; row < this.thresholds.length; row++) {
            if (row === 0 && val < this.thresholds[row][col]) {
                index = 0;
                break;
            } else if (row > 0 && val >= this.thresholds[row - 1][col] && val < this.thresholds[row][col]) {
                index = row;
                break;
            } else if (row === this.thresholds.length - 1 && val >= this.thresholds[row][col]) {
                index = row + 1;
                break;
            }
        }
        return index;
    }

    initializeValues(size: number): number[][] {
        const values: number[][] = new Array<number[]>(size);
        for (let i = 0; i < size; i++) {
            values[i] = new Array<number>(size).fill(0);
        }
        return values;
    };

} //ContingencyTableBuilder

interface ConcordanceResults {
    overallAgreement: number,
    expectedAgreement: number,
    kappa: number,
    seKappa: number,
    lcl: number,
    ucl: number,
}

// Sum matrix by row or col
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

function formatContingencyTable(values: number[][], xLabels: string[], yLabels: string[]): (string | number)[][] {
    const rowSums = sum(values, "row");
    const colSums = sum(values, "col");
    let table: (string | number)[][] = new Array<(string | number)[]>(values.length + 2);
    for (let i = 0; i < values.length + 1; i++) {
        table[i] = new Array<(string | number)>(values[0].length + 2); //create the row
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
    table[table.length - 1] = new Array<(string | number)>(values[0].length + 2)
    //console.log(table);

    // Add row sums
    table[0][table[0].length - 1] = "Total";
    for (let row = 1; row < table.length-1; row++) {
        table[row][table[row].length-1] = rowSums[row-1];
    }
    // Add column sums
    table[table.length-1][0] = "Total";
    for (let col = 1; col < table[0].length-1; col++) {
        table[table.length-1][col] = colSums[col-1];
    }
    // Add total
    table[table.length-1][table[table.length-1].length-1] = colSums.reduce((sum, val) => sum + val, 0);
    return table;
}


class ConcordanceCalculator {
    private table: ContingencyTable;
    private alpha: number;


    constructor(table: ContingencyTable, alpha = 0.05) {
        this.table = table;
        this.alpha = alpha;
    }

    calculate(): ConcordanceResults {
        const n = this.table.values.reduce((sum, row) => sum + row.reduce((rowSum, val) => rowSum + val, 0), 0);
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
        const p = [...(this.table.values)];
        for (let i = 0; i < p.length; i++) {
            p[i] = [...this.table.values[i]]
            for (let j = 0; j < p[i].length; j++) {
                p[i][j] /= n;
            }
        }
        // console.log(p);
        // console.log(this.table.values);

        let observedAgreement = 0;
        for (let i = 0; i < this.table.values.length; i++) {
            observedAgreement += this.table.values[i][i];
        }
        const overallAgreement = observedAgreement / n;

        let expectedAgreement = 0;
        const rowSums: number[] = new Array(this.table.values.length).fill(0);
        const colSums: number[] = new Array(this.table.values[0].length).fill(0);
        const p_i: number[] = new Array(p.length).fill(0); //row sum for proportions p_{i.}
        const p_j: number[] = new Array(p[0].length).fill(0) //column sum for proportions p_{.i}

        for (let i = 0; i < this.table.values.length; i++) {
            for (let j = 0; j < this.table.values[i].length; j++) {
                rowSums[i] += this.table.values[i][j];
                colSums[j] += this.table.values[i][j];
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
            sumPii += p[i][i] * Math.pow((1 - expectedAgreement) - ((p_i[i] + p_j[i]) * (1 - overallAgreement)), 2);
        }

        //
        let sdKappa = (1 / Math.pow(1 - expectedAgreement, 2)) * (Math.sqrt((Math.pow(1 - overallAgreement, 2) * sumPij + sumPii) -
            Math.pow(overallAgreement * expectedAgreement - 2 * expectedAgreement + overallAgreement, 2)));
        let seKappa = sdKappa / Math.sqrt(n);

        let zCrit = normal.inv(1 - this.alpha / 2, 0, 1);

        return {
            overallAgreement: overallAgreement,
            expectedAgreement: expectedAgreement,
            kappa: kappa,
            seKappa: seKappa,
            lcl: kappa - zCrit * seKappa,
            ucl: kappa + zCrit * seKappa,
        };
    }

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

class QualitativeContengencyTableBuilder {
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
        }
    }

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

export {
    ContingencyTableBuilder,
    QualitativeContengencyTableBuilder,
    ConcordanceCalculator,
    sum,
    formatContingencyTable,
};