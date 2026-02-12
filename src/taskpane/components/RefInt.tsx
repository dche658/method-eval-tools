/** Reference Intervals Panel
 * 
 * 
 */

import * as React from "react";

import {
    makeStyles,
    Button,
    Input,
    Select,
    Tooltip,
} from "@fluentui/react-components";

import {
    ref_limit,
    ROBUST,
    adjusted_fisher_pearson_coefficient,
    is_consistent_with_normal
} from "../../reference_intervals";

import RangeInput from "./RangeInput";

// Properties for this component
interface RefIntProps {
    notify: (intent: string, message: string) => void;
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
    selection: {
        maxWidth: "220px",
    },
    inputfield: {
        width: "70px",
        marginLeft: "4px",
        marginRight: "4px",
        marginTop: "4px",
    }
});

export default function RefInt(props: RefIntProps) {
    const styles = useStyles();

    const [riInputRangeValue, setRiInputRangeValue] = React.useState<string>(""); //Reference Interval Input Data
    const [riOutputRangeValue, setRiOutputRangeValue] = React.useState<string>(""); //Reference Interval Output Range - Top Left Cell
    const [riMethodValue, setRiMethodValue] = React.useState<string>(ROBUST);
    const [alphaValue, setAlphaValue] = React.useState<string>("0.05");
    const [confAlphaValue, setConfAlphaValue] = React.useState<string>("0.10");

    const selectRiMethod = (event: React.ChangeEvent<HTMLSelectElement>) => {
        setRiMethodValue(event.target.value);
    }
    const handleAlphaChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        setAlphaValue(event.target.value);
    };
    const handleConfAlphaChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        setConfAlphaValue(event.target.value);
    };

    const calcReferenceInterval = async () => {
        await Excel.run(async (context) => {
            try {
                const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

                if (riInputRangeValue === "") {
                    throw Error("Please select the input range.");
                }
                if (riOutputRangeValue === "") {
                    throw Error("Please select the top left cell for the output range.");
                }
                // Load the ranges
                const inputRange = currentWorksheet.getRange(riInputRangeValue);
                const outputRange = currentWorksheet.getRange(riOutputRangeValue);
                inputRange.load(["rowCount", "columnCount", "values"]);
                await context.sync();

                if (inputRange.rowCount < 120 && riMethodValue === "np") {
                    props.notify("warning", "This is a small data set. Recommended minimum data set size for the non-parametric method is 120 subjects.");
                }
                let data: number[] = [];
                for (let i = 0; i < inputRange.rowCount; i++) {
                    if (inputRange.values[i][0] !== null && inputRange.values[i][0] !== "") {
                        if (typeof inputRange.values[i][0] === "number") {
                            data.push(inputRange.values[i][0]);
                        }
                    }
                }

                //if robust algorith is selected, check the data for skewness and display warning
                //if data is significantly skewed.
                if (riMethodValue === "robust") {
                    const g1 = adjusted_fisher_pearson_coefficient(data);
                    if (!is_consistent_with_normal(data.length, g1, 0.1)) {
                        props.notify("warning",`The robust method should only be used for a symmetric population. The adjusted Fisher-Pearson coefficient is ${g1} indicating the data is skewed.`)
                    }
                }

                const reflimits = ref_limit(data, parseFloat(alphaValue), parseFloat(confAlphaValue), riMethodValue);
                const results: any[][] = [];
                const lowerLimit: number[] = [reflimits.lowerLimit, reflimits.lowerLimitCI[0], reflimits.lowerLimitCI[1]];
                const upperLimit: number[] = [reflimits.upperLimit, reflimits.upperLimitCI[0], reflimits.upperLimitCI[1]];
                results.push(["Ref Limit", "LCL", "UCL"]);
                results.push(lowerLimit);
                results.push(upperLimit);
                const range = outputRange.getAbsoluteResizedRange(results.length, 3);
                range.values = results

                await context.sync();
            } catch (err) {
                props.notify("error", err.message);
            } // try
        });
    }
    return (
        <div className="container">
            <div>
                <p>
                    Derive reference intervals using the non-parametric or
                    robust estimator (Horn's biweight quantile) methods.
                </p>
                <p>
                    The robust algorithm should only be used if the population
                    is symmetric. If this algorithm is selected the add-in will
                    check for skewness by calculating the adjusted Fisher-Pearson 
                    coefficient and display a warning if significant skewness is 
                    detected.
                </p>
            </div>
            <div>
                <RangeInput label="Input Range"
                    rangeValue={riInputRangeValue}
                    setRangeValue={setRiInputRangeValue}
                    validationMessage="Must be a valid Excel range. e.g., A2:A26" />

            </div>
            <div className={styles.field}>
                <label htmlFor="ri-method">Parameter Estimation Method</label>
                <Select value={riMethodValue} id="ri-method"
                    className={styles.selection} onChange={selectRiMethod} >
                    <option value="robust">Robust</option>
                    <option value="np">Non-parametric</option>
                </Select>
            </div>
            <div className={styles.field}>
                <label htmlFor="alpha">Reference Limit Alpha</label>
                <Tooltip content="Default value is 0.05 for 95% confidence interval." relationship="label">
                    <Input type="number" value={alphaValue} id="alpha"
                        className={styles.inputfield} onChange={handleAlphaChange} />
                </Tooltip>
            </div>
            <div className={styles.field}>
                <label htmlFor="conf-alpha">Confidence Limits Alpha</label>
                <Tooltip content="Default value is 0.10 for 90% confidence interval." relationship="label">
                    <Input type="number" value={confAlphaValue} id="conf-alpha"
                        className={styles.inputfield} onChange={handleConfAlphaChange} />
                </Tooltip>
            </div>
            <div>
                <RangeInput label="Output Range"
                    rangeValue={riOutputRangeValue}
                    setRangeValue={setRiOutputRangeValue}
                    validationMessage="Must be a valid Excel cell reference. e.g., E2 " />
            </div>

            <div className={styles.field}>
                <Button appearance="primary" onClick={calcReferenceInterval}>Calculate</Button>
            </div>
        </div>
    );
};
