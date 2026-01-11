/*
* Panel for analysing concordance of qualitative data
*
* Author: Douglas Chesher
*
* Created: October 2025.
*/
import * as React from "react";

import { makeStyles, Button } from "@fluentui/react-components";
import RangeInput from "./RangeInput";

import {
    QualitativeContengencyTableBuilder,
    ConcordanceCalculator
} from "../../concordance";

interface QualitativeProps {
    notify: (intent: string, message: string) => void;
    qualXRangeValue: string;
    qualYRangeValue: string;
    qualOutputRangeValue: string;
    setQualXRangeValue: React.Dispatch<React.SetStateAction<string>>;
    setQualYRangeValue: React.Dispatch<React.SetStateAction<string>>;
    setQualOutputRangeValue: React.Dispatch<React.SetStateAction<string>>;
}

const useStyles = makeStyles({
    container: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
    },
    field: {
        marginLeft: "4px",
        marginBottom: "4px",
        marginTop: "4px",
    },
});

export default function Qualitative(props: QualitativeProps) {
    const styles = useStyles();


    // Run the qualitative data comparison
    const qualComparison = async () => {
        await Excel.run(async (context) => {
            try {
                const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

                // Load the ranges
                const xRange = currentWorksheet.getRange(props.qualXRangeValue);
                const yRange = currentWorksheet.getRange(props.qualYRangeValue);
                xRange.load(["rowCount", "columnCount", "values"]);
                yRange.load(["rowCount", "columnCount", "values"]);
                await context.sync();

                if (xRange.rowCount !== yRange.rowCount) {
                    throw new Error("X and Y ranges must have the same number of rows");
                }
                let x: string[] = [];
                let y: string[] = [];
                for (let i = 0; i < xRange.rowCount; i++) {
                    if (xRange.values[i][0] !== null && xRange.values[i][0] !== "") {
                        if (typeof xRange.values[i][0] === "string") {
                            x.push(xRange.values[i][0]);
                        } else if (typeof xRange.values[i][0] === "number") {
                            x.push(xRange.values[i][0].toString());
                        }
                    }
                    if (yRange.values[i][0] !== null && yRange.values[i][0] !== "") {
                        if (typeof yRange.values[i][0] === "string") {
                            y.push(yRange.values[i][0]);
                        } else if (typeof yRange.values[i][0] === "number") {
                            y.push(yRange.values[i][0].toString());
                        }
                    }
                }
                if (x.length !== y.length) {
                    throw new Error("X and Y ranges must have the same number of rows");
                }
                const contingencyBuilder = new QualitativeContengencyTableBuilder();
                const contingencyTable = contingencyBuilder.build(x, y);
                const concordanceCalculator = new ConcordanceCalculator(contingencyTable);
                const concordanceResults = concordanceCalculator.calculate();
                //to-do: copy to Excel
                const concordanceCell = props.qualOutputRangeValue;
                if (concordanceCell === "") {
                    throw Error("Please select the concordance output range.");
                } else {
                    const concordanceRange = currentWorksheet.getRange(concordanceCell);
                    const contingencyRange = concordanceRange.getAbsoluteResizedRange(contingencyTable.table.length, contingencyTable.table[0].length);
                    contingencyRange.values = contingencyTable.table;
                    const offset = Math.max(8, contingencyTable.table.length + 1)
                    const cohenRange = concordanceRange.getOffsetRange(offset, 0).getAbsoluteResizedRange(6, 2);
                    cohenRange.values = concordanceCalculator.formatResultsAsArray(concordanceResults);
                    await context.sync();
                }
            } catch (err) {
                props.notify("error", err.message);
            } // try
        });
    }

    return (
        <div className="container">
            <div>
                Assess concordance of qualitative data and calculate Cohen's Kappa.
            </div>
            <div>
                <RangeInput label="X Range"
                    rangeValue={props.qualXRangeValue}
                    setRangeValue={props.setQualXRangeValue}
                    validationMessage="Must be a valid Excel range. e.g., A2:A26" />

            </div>
            <div>
                <RangeInput label="Y Range"
                    rangeValue={props.qualYRangeValue}
                    setRangeValue={props.setQualYRangeValue}
                    validationMessage="Must be a valid Excel range. e.g., B2:B26" />
            </div>

            <div>
                <RangeInput label="Output Range"
                    rangeValue={props.qualOutputRangeValue}
                    setRangeValue={props.setQualOutputRangeValue}
                    validationMessage="Must be a valid Excel cell reference. e.g., E2 " />
            </div>

            <div className={styles.field}>
                <Button appearance="primary" onClick={qualComparison}>Run</Button>
            </div>
        </div>
    );
}