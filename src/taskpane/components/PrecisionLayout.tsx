import * as React from "react";

import { makeStyles, Button, Field, Input } from "@fluentui/react-components";
import RangeInput from "./RangeInput";

interface PrecisionLayoutProps {
    notify: (intent: string, message: string) => void;
    factorARangeValue: string;
    factorBRangeValue: string;
    resultsRangeValue: string;
    setFactorARangeValue: React.Dispatch<React.SetStateAction<string>>;
    setFactorBRangeValue: React.Dispatch<React.SetStateAction<string>>;
    setResultsRangeValue: React.Dispatch<React.SetStateAction<string>>;
}

const useStyles = makeStyles({
    container: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
    },
    field: {
        width: "150px",
        marginLeft: "4px",
        marginBottom: "4px",
        marginTop: "4px",
    },
});

export default function PrecisionLayout(props: PrecisionLayoutProps) {
    const styles = useStyles();
    const [numDays, setNumDays] = React.useState<number>(5);
    const [numRuns, setNumRuns] = React.useState<number>(1);
    const [numReps, setNumReps] = React.useState<number>(5);
    const [numLevels, setNumLevels] = React.useState<number>(2);
    const [layoutRangeValue, setLayoutRangeValue] = React.useState<string>("");

    // Setup the layout for a precision experiment
    const setupPrecision = async () => {
        await Excel.run(async (context) => {
            try {
                if (isNaN(numDays) || isNaN(numRuns) || isNaN(numReps) || isNaN(numLevels)) {
                    throw new Error(
                        "Please enter valid numbers for the number of days, runs, replicates and levels."
                    );
                }

                const layout = setupLayout(numDays, numRuns, numReps, numLevels);
                const range = context.workbook.worksheets.getActiveWorksheet().getRange(layoutRangeValue);
                range.load(["rowCount", "columnCount", "values"]);
                await context.sync();

                // write the layout to Excel
                const layoutRange = range.getAbsoluteResizedRange(layout.length, layout[0].length);
                layoutRange.values = layout;
                await context.sync();

                // copy the range addresses to the precision analysis fields
                const daysRange = range.getCell(1, 0).getAbsoluteResizedRange(layout.length - 1, 1);
                const runsRange = range.getCell(1, 1).getAbsoluteResizedRange(layout.length - 1, 1);
                const valuesRange = range.getCell(1, 2).getAbsoluteResizedRange(layout.length - 1, numLevels);
                daysRange.load(["address"]);
                runsRange.load(["address"]);
                valuesRange.load(["address"]);
                await context.sync();
                props.setFactorARangeValue(daysRange.address);
                props.setFactorBRangeValue(runsRange.address);
                props.setResultsRangeValue(valuesRange.address);
            } catch (err) {
                props.notify("error", err.message);
            } // try
        });
    } //setupPrecision

    function setupLayout(days: number, runs: number, reps: number, levels: number) {
        let arr: string[][] = [];
        let headerRow = ["Days", "Runs"];
        for (let i = 1; i <= levels; i++) {
            headerRow.push(`Level ${i}`);
        }
        arr.push(headerRow);
        for (let i = 1; i <= days; i++) {
            for (let j = 1; j <= runs; j++) {
                for (let k = 1; k <= reps; k++) {
                    let row: string[] = [];
                    row.push(`Day ${i}`);
                    row.push(`Run ${j}`);
                    for (let l = 1; l <= levels; l++) {
                        row.push("");
                    }
                    arr.push(row);
                }
            }
        }
        return arr;
    } // setupLayout

    return (
        <div className="container">
            <div>
                Use this section to setup a layout in Excel to assess an assays imprecision. It can
                be used for both verification and validation studies.
            </div>
            <div className={styles.field}>
                <Field label="Days">
                    <Input
                        type="number"
                        value={numDays.toString()}
                        onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                            setNumDays(parseInt(event.target.value));
                        }}
                        min={2}
                        max={100}
                        title="Number of days."
                    />
                </Field>
            </div>
            <div className={styles.field}>
                <Field label="Runs">
                    <Input
                        type="number"
                        value={numRuns.toString()}
                        onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                            setNumRuns(parseInt(event.target.value));
                        }}
                        min={1}
                        max={100}
                        title="Number of runs per day."
                    />
                </Field>
            </div>
            <div className={styles.field}>
                <Field label="Replicates">
                    <Input
                        type="number"
                        value={numReps.toString()}
                        onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                            setNumReps(parseInt(event.target.value));
                        }}
                        min={2}
                        max={100}
                        title="Number of replicates per run."
                    />
                </Field>
            </div>

            <div className={styles.field}>
                <Field label="Levels">
                    <Input
                        type="number"
                        value={numLevels.toString()}
                        onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                            setNumLevels(parseInt(event.target.value));
                        }}
                        min={1}
                        max={10}
                    />
                </Field>
            </div>
            <div>
                <RangeInput label="Layout Range"
                    rangeValue={layoutRangeValue}
                    setRangeValue={setLayoutRangeValue}
                    validationMessage="Must be a valid Excel cell reference. e.g., A26" />
            </div>
            <div className={styles.field}>
                <Button appearance="primary" onClick={setupPrecision}>Setup</Button>
            </div>
        </div>
    );
}