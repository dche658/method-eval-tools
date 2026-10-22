/**
 * 
 */
import * as React from "react";
import { makeStyles, Button, tokens } from "@fluentui/react-components";
import RangeInput from "./RangeInput";
import {
    lognorm,
    reflim,
    QQPlotResult,
    ConfIntResult,
    RefLimResult,
    InterpResult
} from "../../reflim";

interface RefLimProps {
    notify: (intent: string, message: string) => void;
    uitext: { [key: string]: string };
}

const useStyles = makeStyles({
    container: {
        display: "block",
    },
    field: {
        marginLeft: tokens.spacingHorizontalXS,
        marginBottom: tokens.spacingVerticalXS,
        marginTop: tokens.spacingVerticalXS,
    },
});

function isNumber(str: string) {
    return !isNaN(parseFloat(str));
}

// uitext["lbl_ri_n"]: "n",
// uitext["lbl_ri_stats"]: "Mean (sd)",
// uitext["lbl_ri_lognormal"]: "Lognormal Distribution",
// uitext["lbl_ri_reflimits"]: "Reference limit",
// uitext["lbl_ri_ltl"]: "Lower tolerance intervals (EL)",
// uitext["lbl_ri_utl"]: "Upper tolerance intervals (EL)",
// uitext["lbl_ri_lumi"]: "Lower UM intervals",
// uitext["lbl_ri_uumi"]: "Upper UM intervals",
// uitext["lbl_ri_lci"]: "Lower confidence intervals",
// uitext["lbl_ri_uci"]: "Upper confidence intervals",
// uitext["lbl_ri_targets"]: "Target Limits",
// uitext["lbl_ri_ltti"]: "Lower target tolerance intervals (EL)",
// uitext["lbl_ri_utti"]: "Upper target tolerance intervals (EL)",
// uitext["lbl_ri_lumti"]: "Lower UM target intervals",
// uitext["lbl_ri_uumti"]: "Upper UM target intervals",
// uitext["lbl_ri_intll"]: "Interpretation of the lower limit",
// uitext["lbl_ri_intul"]: "Interpretation of the upper limit",
// uitext["lbl_ri_lambda"]: "Lambda for UM",

function formatReport(res: RefLimResult, uitext: { [key: string]: string }): [string, any][] {
    const report: [string, any][] = [];

    report.push([uitext["lbl_ri_n"], res.stats.n_total]);
    report.push([uitext["lbl_ri_stats"], `${res.stats.mean} (${res.stats.sd})`]);
    report.push([uitext["lbl_ri_lognormal"], res.lognormal]);
    report.push([uitext["lbl_ri_reflimits"], `${res.limits.lower_lim} - ${res.limits.upper_lim}`]);
    report.push([uitext["lbl_ri_ltl"], `${res.limits.lower_lim_low} - ${res.limits.lower_lim_upp}`]);
    report.push([uitext["lbl_ri_utl"], `${res.limits.upper_lim_low} - ${res.limits.upper_lim_upp}`]);
    //report.push([uitext["lbl_ri_lumi"], ""]);
    //report.push([uitext["lbl_ri_uumi"], ""]);
    report.push([uitext["lbl_ri_lci"], `${res.confidence_int.lower_lim_low} - ${res.confidence_int.lower_lim_upp}`]);
    report.push([uitext["lbl_ri_uci"], `${res.confidence_int.upper_lim_low} - ${res.confidence_int.upper_lim_upp}`]);
    report.push([uitext["lbl_ri_targets"], `${res.targets.lower_lim} - ${res.targets.upper_lim}`]);
    report.push([uitext["lbl_ri_ltti"], `${res.targets.lower_lim_low} - ${res.targets.lower_lim_upp}`]);
    report.push([uitext["lbl_ri_utti"], `${res.targets.upper_lim_low} - ${res.targets.upper_lim_upp}`]);
    //report.push([uitext["lbl_ri_lumti"], ""]);
    //report.push([uitext["lbl_ri_uumti"], ""]);
    report.push([uitext["lbl_ri_intll"], `${res.dev_lim[0]}`]);
    report.push([uitext["lbl_ri_intul"], `${res.dev_lim[1]}`]);
    //report.push([uitext["lbl_ri_lambda"], ""]);
    return report;
}

export default function RefLim(props: RefLimProps) {
    const styles = useStyles();

    const [dataRangeValue, setDataRangeValue] = React.useState<string>(""); //Reference Interval Input Data
    const [targetLowRangeValue, setTargetLowRangeValue] = React.useState<string>(""); //Reference Interval Output Range - Top Left Cell
    const [targetHighRangeValue, setTargetHighRangeValue] = React.useState<string>("");
    const [resultRangeValue, setResultRangeValue] = React.useState<string>("");
    const [chartRangeValue, setChartRangeValue] = React.useState<string>("");

    const analyze = async () => {
        await Excel.run(async (context) => {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

            let targets: number[] | null = [];
            let plot_it: boolean = false;

            try {
                if (dataRangeValue == null || dataRangeValue.length < 5) {
                    throw new Error("A valid data range must be specified.");
                }
                // Read the worksheet ranges
                const dataRange = currentWorksheet.getRange(dataRangeValue);
                dataRange.load(["rowCount", "columnCount", "values"]);
                let targetLowRange = null;
                // Check if targets supplied
                if (targetLowRangeValue != null && targetLowRangeValue.length > 1) {
                    if (isNumber(targetLowRangeValue)) {
                        //Parse if numeric supplied
                        targets.push(parseFloat(targetLowRangeValue));
                    } else {
                        //Read from range if cell reference supplied
                        targetLowRange = currentWorksheet.getRange(targetLowRangeValue);
                        targetLowRange.load(["rowCount", "columnCount", "values"]);
                    }
                }
                let targetHighRange = null;
                if (targetHighRangeValue != null && targetHighRangeValue.length > 1) {
                    if (isNumber(targetHighRangeValue)) {
                        //Parse if numeric supplied
                        targets.push(parseFloat(targetHighRangeValue));
                    } else {
                        //Read from range if cell reference supplied
                        targetHighRange = currentWorksheet.getRange(targetHighRangeValue);
                        targetHighRange.load(["rowCount", "columnCount", "values"]);
                    }
                }
                if (resultRangeValue == null || resultRangeValue.length < 2) {
                    throw new Error("The top left cell for the results range must be specified");
                }
                const resultRange = currentWorksheet.getRange(resultRangeValue);
                resultRange.load(["rowCount", "columnCount", "values"]);
                let chartRange;
                if (chartRangeValue == null || chartRangeValue.length < 2) {
                    plot_it = false;
                } else {
                    chartRange = currentWorksheet.getRange(chartRangeValue);
                    chartRange.load(["rowCount", "columnCount", "values"]);
                    plot_it = true;
                }

                await context.sync();

                // Read data from range
                if (dataRange.rowCount < 40) throw new Error("Data must include at least 40 numeric values");

                //Import numeric data only
                let data: number[] = [];
                for (let i = 0; i < dataRange.rowCount; i++) {
                    if (typeof dataRange.values[i][0] === "number") {
                        data.push(dataRange.values[i][0]);
                    }
                }

                // Read targets from range
                if (targetLowRange != null && targetHighRange != null) {
                    if (isNumber(targetLowRange.values[0][0])) {
                        targets.push(parseFloat(targetLowRange.values[0][0]));
                    }
                    if (isNumber(targetHighRange.values[0][0])) {
                        targets.push(parseFloat(targetHighRange.values[0][0]));
                    }
                }

                if (targets.length < 2) targets = null;
                const lognormal = lognorm(data).lognormal;
                //Parameters
                //data[]
                //lognormal boolean
                //targets[2]
                //perc.trunc = 2.5
                //n.min = 200
                //apply_rounding boolean
                const res: RefLimResult = reflim(data, lognormal, targets, 2.5, 200, true);

                //Calculate report range rows and columns
                const rpt = formatReport(res, props.uitext);
                let deltaRows = rpt.length - resultRange.rowCount;
                let deltaCols = 2 - resultRange.columnCount;
                const reportRange = resultRange.getResizedRange(deltaRows, deltaCols);
                reportRange.values = rpt;
                await context.sync();

            } catch (error) {
                console.log(error);
                if (error instanceof Error) {
                    props.notify("error", error.message);
                }
            }
        });
    }

    return (
        <div className="container">
            <div>{props.uitext["inf_reflim"]}</div>
            <div>
                <RangeInput
                    label={props.uitext["lbl_data_range"]}
                    rangeValue={dataRangeValue}
                    setRangeValue={setDataRangeValue}
                    validationMessage={props.uitext["msg_ridata_range"]}
                    uitext={props.uitext}
                />
            </div>
            <div>
                <RangeInput
                    label={props.uitext["lbl_tar_low_range"]}
                    rangeValue={targetLowRangeValue}
                    setRangeValue={setTargetLowRangeValue}
                    validationMessage={props.uitext["msg_tar_range"]}
                    uitext={props.uitext}
                    allowNumbers={true}
                    tooltipContent={props.uitext["tip_ri_tar"]}
                />
            </div>

            <div>
                <RangeInput
                    label={props.uitext["lbl_tar_high_range"]}
                    rangeValue={targetHighRangeValue}
                    setRangeValue={setTargetHighRangeValue}
                    validationMessage={props.uitext["msg_tar_range"]}
                    uitext={props.uitext}
                    allowNumbers={true}
                    tooltipContent={props.uitext["tip_ri_tar"]}
                />
            </div>
            <div>
                <RangeInput
                    label={props.uitext["lbl_results_range"]}
                    rangeValue={resultRangeValue}
                    setRangeValue={setResultRangeValue}
                    validationMessage={props.uitext["msg_ri_results_range"]}
                    uitext={props.uitext}
                    tooltipContent={props.uitext["tip_ri_resrange"]}
                />
            </div>
            <div>
                <RangeInput
                    label={props.uitext["lbl_reflim_chart_range"]}
                    rangeValue={chartRangeValue}
                    setRangeValue={setChartRangeValue}
                    validationMessage={props.uitext["msg_ri_chart_range"]}
                    uitext={props.uitext}
                    tooltipContent={props.uitext["tip_ri_chtrange"]}
                />
            </div>
            <div className={styles.field}>
                <Button appearance="primary" onClick={analyze}>
                    {props.uitext["btn_run"]}
                </Button>
            </div>
        </div>
    );
}