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
    uitext: {[key:string]: string};
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
                    {props.uitext["inf_ref_int1"]}
                </p>
                <p>
                    {props.uitext["inf_ref_int2"]}
                </p>
            </div>
            <div>
                <RangeInput label={props.uitext["lbl_input_range"]}
                    rangeValue={riInputRangeValue}
                    setRangeValue={setRiInputRangeValue}
                    validationMessage={props.uitext["msg_input_range"]}
                    uitext={props.uitext} />

            </div>
            <div className={styles.field}>
                <label htmlFor="ri-method">{props.uitext["lbl_ri_method"]}</label>
                <Select value={riMethodValue} id="ri-method"
                    className={styles.selection} onChange={selectRiMethod} >
                    <option value="robust">{props.uitext["opt_robust"]}</option>
                    <option value="np">{props.uitext["opt_np"]}</option>
                </Select>
            </div>
            <div className={styles.field}>
                <label htmlFor="alpha">{props.uitext["lbl_ri_alpha"]}</label>
                <Tooltip content={props.uitext["tip_ri_alpha"]} relationship="label">
                    <Input type="number" value={alphaValue} id="alpha"
                        className={styles.inputfield} onChange={handleAlphaChange} />
                </Tooltip>
            </div>
            <div className={styles.field}>
                <label htmlFor="conf-alpha">{props.uitext["lbl_ci_alpha"]}</label>
                <Tooltip content={props.uitext["tip_ci_alpha"]} relationship="label">
                    <Input type="number" value={confAlphaValue} id="conf-alpha"
                        className={styles.inputfield} onChange={handleConfAlphaChange} />
                </Tooltip>
            </div>
            <div>
                <RangeInput label={props.uitext["lbl_output_range"]}
                    rangeValue={riOutputRangeValue}
                    setRangeValue={setRiOutputRangeValue}
                    validationMessage={props.uitext["msg_output_range"]}
                    uitext={props.uitext} />
            </div>

            <div className={styles.field}>
                <Button appearance="primary" onClick={calcReferenceInterval}>{props.uitext["btn_run"]}</Button>
            </div>
        </div>
    );
};
