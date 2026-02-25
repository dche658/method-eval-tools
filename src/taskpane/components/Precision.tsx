/*
* Analysis of imprecision panel.
*
* Data are assumed to be in a hybrid long format with a column for the day,
* a column for the run, and then tabulated with n columns for each level
* to be assessed.
* 
* Author: Douglas Chesher
*
* Created: August 2025.
*/
import * as React from "react";

import { makeStyles, Button } from "@fluentui/react-components";
import RangeInput from "./RangeInput";
import {
    OneFactorVarianceAnalysis, TwoFactorVarianceAnalysis,
    OneFactorVariance, TwoFactorVariance,
    grubbsTest, Outlier,
} from "../../precision";


interface PrecisionProps {
    notify: (intent: string, message: string) => void;
    factorARangeValue: string;
    factorBRangeValue: string;
    resultsRangeValue: string;
    precOutRangeValue: string;
    setFactorARangeValue: React.Dispatch<React.SetStateAction<string>>;
    setFactorBRangeValue: React.Dispatch<React.SetStateAction<string>>;
    setResultsRangeValue: React.Dispatch<React.SetStateAction<string>>;
    setPrecOutRangeValue: React.Dispatch<React.SetStateAction<string>>;
    uitext: {[key:string]: string};
}


const useStyles = makeStyles({
    container: {
        display: "block",
    },
    field: {
        marginLeft: "4px",
        marginBottom: "4px",
        marginTop: "4px",
    },
});

export default function Precision(props: PrecisionProps) {
    const styles = useStyles();


    interface DataObject {
        type: string;
        aLevels: (number | string)[];
        results: number[];
        aDict: {};
        bLevels: (number | string)[];
        bDict: {};
    }

    interface PrecisionData {
        numLevels: number;
        data: DataObject[];
    }

    const grubbsTestForOutliers = async () => {
        await Excel.run(async (context) => {
            const outlierList: { out: Outlier, col: number }[] = [];
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const rRange = currentWorksheet.getRange(props.resultsRangeValue);
            rRange.load(["rowCount", "columnCount", "values"]);
            await context.sync();

            //check for outliers for each column
            for (let i = 0; i < rRange.columnCount; i++) {
                // array for processing data
                let results: number[] = [];
                // array to keep track of value index on spreadsheet.
                // needed this to allow for missing values
                const arr: number[] = [];

                for (let j = 0; j < rRange.rowCount; j++) {
                    // add value to results array only if numeric
                    if (typeof rRange.values[j][i] == "number") {
                        results.push(rRange.values[j][i]);
                    }
                    // add value to arr array
                    arr.push(rRange.values[j][i]);
                }
                if (results.length > 0) {
                    // run grubbs test
                    const outlier = grubbsTest(results);
                    if (outlier.outlier !== undefined && typeof outlier.outlier === "number") {
                        //In case there a missing values before the outlier
                        const arrIdx = arr.indexOf(outlier.outlier);
                        outlier.index = arrIdx;
                        const item = { out: outlier, col: i };
                        outlierList.push(item);
                        const cell = rRange.getCell(outlier.index, i);
                        cell.format.fill.color = "yellow";
                    }
                }
            }
            // display message
            if (outlierList.length > 0) {
                let msg = "Outliers found:";
                for (let i = 0; i < outlierList.length; i++) {
                    msg += "Row: " + (outlierList[i].out.index + 1) + ", Col: " + (outlierList[i].col + 1) + ", Value: " + outlierList[i].out.outlier + "\n";
                }
                props.notify("info", msg);
            } else {
                props.notify("info", "No outliers found.");
            }
            await context.sync();
        });
    }

    const runANOVA = async () => {
        await Excel.run(async (context) => {
            try {
                const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

                let analysisType = "one-factor";

                // Load the ranges
                const aRange = currentWorksheet.getRange(props.factorARangeValue);
                const bRange = currentWorksheet.getRange(props.factorBRangeValue);
                const rRange = currentWorksheet.getRange(props.resultsRangeValue);
                let oRange = currentWorksheet.getRange(props.precOutRangeValue);
                aRange.load(["rowCount", "columnCount", "values"]);
                bRange.load(["rowCount", "columnCount", "values"]);
                rRange.load(["rowCount", "columnCount", "values"]);
                oRange.load(["rowCount", "columnCount", "values"]);
                await context.sync();

                // Create and object to hold data
                // data holds the data objects after pre-processing for each column
                // vca holds the results of the variance component analysis for each column
                let precisionData: PrecisionData = {
                    numLevels: 0,
                    data: new Array(aRange.columnCount),
                };

                // Process separately for each column in the results range.
                // Need to process first to see if columns contain data before undertaking
                // variance component analysis.
                for (let i = 0; i < rRange.columnCount; i++) {
                    let aLevels: (number | string)[] = [];
                    let bLevels: (number | string)[] = [];
                    let results: number[] = [];
                    let aDict = {};
                    let bDict = {};
                    for (let j = 0; j < rRange.rowCount; j++) {
                        if (typeof rRange.values[j][i] == "number") {
                            results.push(rRange.values[j][i]);
                            let factorA = aRange.values[j][0];
                            if (typeof factorA === "number" || typeof factorA === "string") {
                                aLevels.push(factorA);
                                if (factorA in aDict) {
                                    aDict[factorA] = aDict[factorA] + 1;
                                } else {
                                    aDict[factorA] = 1;
                                }
                            } else {
                                throw new Error(
                                    `Data exists for results column ${i + 1} but factor A has not been specified at row ${j}`
                                );
                            } // factor A exists
                            if (props.factorBRangeValue !== "") {
                                let factorB = bRange.values[j][0];
                                if (typeof factorB === "number" || typeof factorB === "string") {
                                    bLevels.push(factorB);
                                    if (factorB in bDict) {
                                        bDict[factorB] = bDict[factorB] + 1;
                                    } else {
                                        bDict[factorB] = 1;
                                    }
                                }
                            } // factor B exists
                        } // if result is number
                    } // for row j

                    // This should never happen but just in case
                    if (results.length !== aLevels.length) {
                        throw new Error(
                            `Length of results ${results.length} does not match length of factor A ${aLevels.length}.`
                        );
                    }

                    if (results.length > 1) {
                        precisionData.numLevels += 1;
                    }

                    // if no factor B or only one level then perform one factor analysis
                    // Default is one factor analysis
                    const dataObject: DataObject = {
                        type: "one-factor",
                        aLevels: aLevels,
                        results: results,
                        aDict: aDict,
                        bLevels: [],
                        bDict: {},
                    };
                    // More than one level for B, so do two-factor analysis
                    if (Object.keys(bDict).length > 1) {
                        analysisType = "two-factor";
                        dataObject["type"] = analysisType;
                        dataObject["bLevels"] = bLevels;
                        dataObject["bDict"] = bDict;
                    }
                    precisionData.data[i] = dataObject;
                } // for column i
                //Iterate over the column data
                let rangeValues: (string | number)[][];
                if (analysisType === "one-factor") {
                    rangeValues = analyseOneFactor(precisionData);
                } else {
                    rangeValues = analyseTwoFactor(precisionData);
                }
                let rowDelta = rangeValues.length - oRange.rowCount;
                let colDelta = precisionData.data.length - oRange.columnCount + 1;
                oRange = oRange.getResizedRange(rowDelta, colDelta);
                oRange.load(["rowCount", "columnCount", "values"]);
                await context.sync();
                oRange.values = rangeValues;
                await context.sync();
            } catch (err) {
                props.notify("error", err.message);
            } // try
        });
    } // runPrecisionAnalysis

    function analyseOneFactor(precisionData: PrecisionData): (string | number)[][] {
        let columnData = new Array(precisionData.data.length);
        let res: OneFactorVariance;
        for (let i = 0; i < precisionData.data.length; i++) {
            let dataObj = precisionData.data[i];

            if (dataObj.results.length > 0) {
                const oneFactorAnalysis = new OneFactorVarianceAnalysis(
                    dataObj.aLevels,
                    dataObj.results,
                    precisionData.numLevels,
                    0.05
                );
                res = oneFactorAnalysis.calculate();
                columnData[i] = Object.values(res);
            } else {
                columnData[i] = null;
            }
        } // dataObj of precisionData.data
        let labels = [
            "Mean",
            "SS Total",
            "SST",
            "SSE",
            "DF Total",
            "DF Error",
            "DF T",
            "MST",
            "MSE",
            "F",
            "N",
            "P",
            "Sum GrpN2",
            "Var Error",
            "Var B",
            "SD Error",
            "CV Error",
            "SD B",
            "CV B",
            "SD WL",
            "CV WL",
            "DF WL",
            "ChiSq E",
            "ChiSq WL",
            "F Error",
            "F WL",
        ];
        return formatPrecisionForExcel(labels, columnData);
    }

    function analyseTwoFactor(precisionData: PrecisionData): (string | number)[][] {
        let columnData = new Array(precisionData.data.length);
        let res: TwoFactorVariance;
        let maxSize = 0;
        for (let i = 0; i < precisionData.data.length; i++) {
            let dataObj = precisionData.data[i];
            if (dataObj.results.length > 0) {
                const twoFactorAnalysis = new TwoFactorVarianceAnalysis(
                    dataObj.aLevels,
                    dataObj.bLevels,
                    dataObj.results,
                    0.05
                );
                res = twoFactorAnalysis.calculate();
                let size = Object.keys(res).length;
                if (size > maxSize) maxSize = size;
                columnData[i] = Object.values(res);
            } else {
                columnData[i] = null;
            }
        } // for i
        const labels = [
            "Mean",
            "SST",
            "SSA",
            "SSB",
            "SSE",
            "DF T",
            "DF A",
            "DF B",
            "DF E",
            "MSA",
            "MSB",
            "MSE",
            "F A",
            "F B",
            "N",
            "Num A",
            "Num B",
            "Num E",
            "Var T",
            "Var A",
            "Var B",
            "Var E",
            "SD WL",
            "SD A",
            "SD B",
            "SD E",
            "CV WL",
            "CV A",
            "CV B",
            "CV E",
            "DF WL",
            "SD WL LCL",
            "SD WL UCL",
            "SD E LCL",
            "SD E UCL",
            "CV WL LCL",
            "CV WL UCL",
            "CV E LCL",
            "CV E UCL",
        ];
        return formatPrecisionForExcel(labels, columnData);
    } // analyseTwoFactor

    function formatPrecisionForExcel(labels: string[], columnData: number[][]): (string | number)[][] {
        let rangeValues: (string | number)[][] = [];
        for (let i = 0; i < labels.length; i++) {
            let numCols = columnData.length + 1; //add column for label
            let row: (string | number)[] = new Array(numCols);
            for (let j = 0; j < numCols; j++) {
                if (j === 0) {
                    row[j] = labels[i];
                } else {
                    if (columnData[j - 1] !== null) {
                        // set value if column is has data
                        row[j] = columnData[j - 1][i];
                    } else {
                        row[j] = "";
                    }
                }
            }
            rangeValues.push(row);
        }
        return rangeValues;
    } //formatPrecisionForExcel

    return (
        <div className="container">
            <div>
                {props.uitext["inf_precision"]}
            </div>
            <div>
                <RangeInput label={props.uitext["lbl_days_range"]}
                    rangeValue={props.factorARangeValue}
                    setRangeValue={props.setFactorARangeValue}
                    validationMessage={props.uitext["msg_days_range"]}
                    uitext={props.uitext} />
            </div>
            <div>
                <RangeInput label={props.uitext["lbl_runs_range"]}
                    rangeValue={props.factorBRangeValue}
                    setRangeValue={props.setFactorBRangeValue}
                    validationMessage={props.uitext["msg_runs_range"]}
                    uitext={props.uitext} />
            </div>

            <div>
                <RangeInput label={props.uitext["lbl_results_range"]}
                    rangeValue={props.resultsRangeValue}
                    setRangeValue={props.setResultsRangeValue}
                    validationMessage={props.uitext["msg_results_range"]}
                    uitext={props.uitext} />
            </div>
            <div className={styles.field}>
                <Button appearance="primary" 
                onClick={grubbsTestForOutliers}>{props.uitext["btn_grubb_test"]}</Button>
            </div>
            <div>
                <RangeInput label={props.uitext["lbl_output_range"]}
                    rangeValue={props.precOutRangeValue}
                    setRangeValue={props.setPrecOutRangeValue}
                    validationMessage={props.uitext["msg_output_range"]}
                    uitext={props.uitext} />
            </div>
            <div className={styles.field}>
                <Button appearance="primary"
                onClick={runANOVA}>{props.uitext["btn_run"]}</Button>
            </div>
        </div>
    );
}