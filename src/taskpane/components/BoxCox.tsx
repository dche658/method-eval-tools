/** Box-Cox Transformation Panel
 * 
 * 
 */

import * as React from "react";

import {
    makeStyles,
    Button,
    Input,
    Select,
} from "@fluentui/react-components";

import {
    boxcoxfit,
    BoxCoxTransform
} from "../../reference_intervals";

import RangeInput from "./RangeInput";

// Properties for this component
interface BoxCoxProps {
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
        marginRight: "4px",
        marginTop: "4px",
    },
    grid1x3: {
        display: "grid",
        gridTemplateColumns: " 1fr 1fr 1fr",
        gridTemplateRows: "1fr",
        gridGap: "4px",
    },
    box: {
        padding: "4px",
    }
});

export default function BoxCox(props: BoxCoxProps) {
    const styles = useStyles();

    const [bcInputRangeValue, setBcInputRangeValue] = React.useState<string>(""); //Box Cox Input Data
    const [bcOutputRangeValue, setBcOutputRangeValue] = React.useState<string>(""); //Box Cox Output Range
    const [bcMethodValue, setBcMethodValue] = React.useState<string>("sw");
    const [lambdaStartValue, setLambdaStartValue] = React.useState<string>("-2");
    const [lambdaStopValue, setLambdaStopValue] = React.useState<string>("2");
    const [lambdaStepValue, setLambdaStepValue] = React.useState<string>("0.01");


    const selectBcMethod = (event: React.ChangeEvent<HTMLSelectElement>) => {
        setBcMethodValue(event.target.value);
    }

    const handleLambdaStartChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        setLambdaStartValue(event.target.value);
    };
    const handleLambdaStopChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        setLambdaStopValue(event.target.value);
    };
    const handleLambdaStepChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        setLambdaStepValue(event.target.value);
    };

    function mapBoxCoxResults(results: BoxCoxTransform, method: string): any[][] {
        let arr: any[] = [results.transformedData.length];
        for (let i = 0; i < results.transformedData.length; i++) {
            if (method === "sw") {
                arr[i] = [4];
            } else if (method === "mle") {
                arr[i] = [3];
            }
            arr[i][0] = results.transformedData[i];
            if (i === 0) {
                arr[i][1] = results.lambda;
                if (method === "sw") {
                    arr[i][2] = results.shapiroWilkValue.w;
                    arr[i][3] = results.shapiroWilkValue.p;
                } else if (method === "mle") {
                    arr[i][2] = results.maximumLikelihoodValue;
                }
            } else {
                arr[i][1] = "";
                arr[i][2] = "";
                if (method === "sw") arr[i][3] = "";
            }
        }
        //Add column headers
        if (method === "sw") {
            arr.splice(0, 0, ["Transformed", "Lambda", "Shapiro-Wilk W", "p-value"])
        } else if (method === "mle") {
            arr.splice(0, 0, ["Transformed", "Lambda", "MLE Value"])
        }
        console.log(arr);
        return arr;
    }// mapBoxCoxResults

    const boxCoxTransform = async () => {
        await Excel.run(async (context) => {
            try {
                const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

                // Load the ranges
                const inputRange = currentWorksheet.getRange(bcInputRangeValue);
                inputRange.load(["rowCount", "columnCount", "values"]);
                await context.sync();

                if (inputRange.rowCount < 50) {
                    props.notify("warning", "This is a small data set and the estimated value of lambda may be inaccurate.");
                }
                let data: number[] = [];
                for (let i = 0; i < inputRange.rowCount; i++) {
                    if (inputRange.values[i][0] !== null && inputRange.values[i][0] !== "") {
                        if (typeof inputRange.values[i][0] === "number") {
                            data.push(inputRange.values[i][0]);
                        }
                    }
                }
                console.log(data);
                // Do Box Cox transform and format for copying to excel
                let boxCoxResults: BoxCoxTransform;
                const lambda = [parseFloat(lambdaStartValue), parseFloat(lambdaStopValue), parseFloat(lambdaStepValue)];
                if (bcMethodValue === "sw") {
                    boxCoxResults = boxcoxfit(data, lambda, "sw");
                } else if (bcMethodValue === "mle") {
                    boxCoxResults = boxcoxfit(data, lambda, "mle");
                } else {
                    throw new Error("No valid Box-Cox estimation method selected."); //This should never be required
                }
                const boxCoxResultsArray = mapBoxCoxResults(boxCoxResults, bcMethodValue);

                //Copy to Excel
                const boxCoxOutputCell = bcOutputRangeValue;
                if (boxCoxOutputCell === "") {
                    throw Error("Please select the top left cell for the output range.");
                } else {
                    const boxCoxOutputRange = currentWorksheet.getRange(boxCoxOutputCell);
                    const outputRange = boxCoxOutputRange.getAbsoluteResizedRange(boxCoxResultsArray.length, boxCoxResultsArray[0].length);
                    outputRange.values = boxCoxResultsArray;
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
                <p>
                    {props.uitext["inf_box_cox"]}
                </p>
            </div>
            <div>
                <RangeInput label={props.uitext["lbl_input_range"]}
                    rangeValue={bcInputRangeValue}
                    setRangeValue={setBcInputRangeValue}
                    validationMessage={props.uitext["msg_input_range"]}
                    uitext={props.uitext} />

            </div>
            <div>
                <RangeInput label={props.uitext["lbl_output_range"]}
                    rangeValue={bcOutputRangeValue}
                    setRangeValue={setBcOutputRangeValue}
                    validationMessage={props.uitext["msg_output_range"]}
                    uitext={props.uitext} />
            </div>
            <div className={styles.field}>
                <label htmlFor="bc-method">{props.uitext["lbl_bc_method"]}</label>
                <Select value={bcMethodValue} id="bc-method"
                    className={styles.selection} onChange={selectBcMethod} >
                    <option value="sw">{props.uitext["opt_sw"]}</option>
                    <option value="mle">{props.uitext["opt_mle"]}</option>
                </Select>
            </div>

            <div className={styles.box}>
                <span>{props.uitext["lbl_lambda_search"]}</span>
                <div className={styles.grid1x3}>

                    <div className={styles.field}>
                        <label htmlFor="lambda-start">{props.uitext["lbl_start"]}</label>
                        <Input type="number" value={lambdaStartValue} id="lambda-start"
                            className={styles.inputfield} onChange={handleLambdaStartChange} />
                    </div>

                    <div className={styles.field}>
                        <label htmlFor="lambda-stop">{props.uitext["lbl_end"]}</label>
                        <Input type="number" value={lambdaStopValue} id="lambda-stop"
                            className={styles.inputfield} onChange={handleLambdaStopChange} />
                    </div>

                    <div className={styles.field}>
                        <label htmlFor="lambda-step">{props.uitext["lbl_step"]}</label>
                        <Input type="number" value={lambdaStepValue} id="lambda-step"
                            className={styles.inputfield} onChange={handleLambdaStepChange} />
                    </div>
                </div>
            </div>
            <div className={styles.field}>
                <Button appearance="primary" onClick={boxCoxTransform}>{props.uitext["btn_run"]}</Button>
            </div>
        </div>
    );
};
