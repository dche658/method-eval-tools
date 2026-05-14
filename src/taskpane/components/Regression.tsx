/**
 * Regression.tsx
 * Regression analysis component for the task pane.
 * This component allows users to perform regression analysis on their data
 * using Deming, Weighted Deming, and Passing-Bablok methods.
 * 
 * @author Douglas Chesher
 */
import * as React from "react";

import {
  makeStyles,
  useId,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Select,
  Button,
  Checkbox,
  Field,
  Input,
  Card,
} from "@fluentui/react-components";

import RangeInput from "./RangeInput";
import ConcordanceThreshold from "./ConcordanceThreshold";

import {
  JackknifeConfidenceInterval,
  BootstrapConfidenceInterval,
  PassingBablokRegression,
  DemingRegression,
  WeightedDemingRegression,
  DEFAULT_ERROR_RATIO,
} from "../../regression";

import { ExcelBlandAltmanChart, ExcelRegressionChart } from "../../charts";

import {
  ContingencyTableBuilder,
  QualitativeContengencyTableBuilder,
  ConcordanceCalculator,
} from "../../concordance";

// Properties for this component
interface RegressionProps {
  apsAbsValue: string;
  setApsAbsValue: React.Dispatch<React.SetStateAction<string>>;
  apsRelValue: string;
  setApsRelValue: React.Dispatch<React.SetStateAction<string>>;
  xRangeValue: string;
  setXRangeValue: React.Dispatch<React.SetStateAction<string>>;
  yRangeValue: string;
  setYRangeValue: React.Dispatch<React.SetStateAction<string>>;
  compOutRangeValue: string;
  setCompOutRangeValue: React.Dispatch<React.SetStateAction<string>>;
  baRangeValue: string;
  setBaRangeValue: React.Dispatch<React.SetStateAction<string>>; // Bland-Altman output range
  scRangeValue: string;
  setScRangeValue: React.Dispatch<React.SetStateAction<string>>; // Scatter chart output range
  cdRangeValue: string;
  setCdRangeValue: React.Dispatch<React.SetStateAction<string>>; // Chart data output range
  errorRatio: string;
  setErrorRatio: React.Dispatch<React.SetStateAction<string>>;
  regressionType: string;
  setRegressionType: React.Dispatch<React.SetStateAction<string>>; // Default to "paba" (Passing-Bablok)
  useCalcErrorRatio: boolean;
  setUseCalcErrorRatio: React.Dispatch<React.SetStateAction<boolean>>; // Default to false
  importApsSettings: boolean;
  setImportApsSettings: React.Dispatch<React.SetStateAction<boolean>>; // Default to false
  labelOutput: boolean;
  setLabelOutput: React.Dispatch<React.SetStateAction<boolean>>; // Default to true
  differencePlotType: string;
  setDifferencePlotType: React.Dispatch<React.SetStateAction<string>>; // Default to "rel"
  concordanceOutputRange: string;
  setConcordanceOutputRange: React.Dispatch<React.SetStateAction<string>>;
  ciMethod: string;
  setCiMethod: React.Dispatch<React.SetStateAction<string>>; // Default to "bootstrap"
  xThreshold0: string;
  setXThreshold0: React.Dispatch<React.SetStateAction<string>>;
  xThreshold1: string;
  setXThreshold1: React.Dispatch<React.SetStateAction<string>>;
  xThreshold2: string;
  setXThreshold2: React.Dispatch<React.SetStateAction<string>>;
  xThreshold3: string;
  setXThreshold3: React.Dispatch<React.SetStateAction<string>>;
  yThreshold0: string;
  setYThreshold0: React.Dispatch<React.SetStateAction<string>>;
  yThreshold1: string;
  setYThreshold1: React.Dispatch<React.SetStateAction<string>>;
  yThreshold2: string;
  setYThreshold2: React.Dispatch<React.SetStateAction<string>>;
  yThreshold3: string;
  setYThreshold3: React.Dispatch<React.SetStateAction<string>>;
  notify: (intent: string, message: string) => void;
  uitext: { [key: string]: string };
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  container: {
    diplay: "flex",
    flexDirection: "column",
    justifyContent: "center",
  },
  field: {
    marginLeft: "4px",
    marginBottom: "4px",
    marginTop: "4px",
  },
  mvwhidden: {
    marginLeft: "4px",
    marginBottom: "4px",
    display: "none",
  },
  panels: {
    padding: "0 10px",
    "& th": {
      textAlign: "left",
      padding: "0 30px 0 0",
    },
  },
  selection: {
    maxWidth: "220px",
  },
  label: {
    fontWeight: "bold",
  },
  specificationsPanel: {   
    display: "flex",
    flexDirection: "column",
    marginBottom: "10px",
  },
});

/* Interface for the input data
 */
interface InputData {
  means: number[];
  x1: number[];
  x2: number[];
  devsq: number[];
  sd: number;
  cv: number;
  size: number;
  mean: number;
}

/*
 * Read data from the excel range and copy to an array. If there is more than one column
 * then calculate the mean.
 */
function processRangeData(range: Excel.Range): InputData {
  let means: number[] = [];
  let size = 0;

  // Validate the range data
  if (range === null || range === undefined) {
    throw new Error("Range is null or undefined");
  }
  if (range.rowCount === 0 || range.columnCount === 0) {
    throw new Error("Range is empty");
  }
  if (range.columnCount > 2) {
    throw new Error("Range has more than two columns");
  }
  let x1: number[] = [];
  let x2: number[] = [];
  let devsq: number[] = [];
  let sd = 0;
  let cv = 0;
  let mean = 0;

  // If the range has two columns then we calculate assuming analysis was done in replicates
  if (range.columnCount === 2) {
    // Process two columns
    for (let i = 0; i < range.rowCount; i++) {
      //console.log(i, typeof range.values[i][0], typeof range.values[i][1]);
      if (typeof range.values[i][0] === "number" && typeof range.values[i][1] === "number") {
        x1.push(range.values[i][0]);
        x2.push(range.values[i][1]);
        // Calculate sum of squares of deviations
        devsq.push(
          Math.pow(range.values[i][0] - (range.values[i][1] + range.values[i][0]) / 2, 2) +
            Math.pow(range.values[i][1] - (range.values[i][0] + range.values[i][1]) / 2, 2)
        );
        // Calculate means for each pair
        means.push((range.values[i][0] + range.values[i][1]) / 2);
        size++;
      }
    }
    // Calculate standard deviation and coefficient of variation of the replicates.
    sd = Math.sqrt(devsq.reduce((a, b) => a + b, 0) / (size - 1));
    mean = means.reduce((a, b) => a + b, 0) / size;
    cv = sd / mean;
  } else if (range.columnCount === 1) {
    // Process single column
    for (let i = 0; i < range.rowCount; i++) {
      if (typeof range.values[i][0] === "number") {
        x1.push(range.values[i][0]);
        means.push(range.values[i][0]);
        size++;
      }
    }
    mean = means.reduce((a, b) => a + b, 0) / size;
  }
  return { means: means, x1: x1, x2: x2, devsq: devsq, sd: sd, cv: cv, size: size, mean: mean };
}

interface RegressionResults {
  slope: number;
  intercept: number;
  slopeLCL: number;
  slopeUCL: number;
  interceptLCL: number;
  interceptUCL: number;
  slopeSE: number;
  interceptSE: number;
}

/* Main function to perform regression analysis on the supplied data
 */
function doRegression(
  x: number[],
  y: number[],
  regressionMethod: string,
  ciMethod: string,
  errorRatio: number = DEFAULT_ERROR_RATIO
): RegressionResults {
  //console.log(`Error ratio is ${errorRatio}`);
  let res: RegressionResults = {
    slope: NaN,
    intercept: NaN,
    slopeSE: NaN,
    interceptSE: NaN,
    slopeLCL: NaN,
    slopeUCL: NaN,
    interceptLCL: NaN,
    interceptUCL: NaN,
  };
  if (regressionMethod === "paba") {
    let regression = new PassingBablokRegression();
    if (ciMethod === "bootstrap") {
      let regCI = new BootstrapConfidenceInterval(x, y, regression, 1000);
      res = regCI.calculate();
    } else {
      let reg = regression.calculate(x, y);
      res.slope = reg.slope;
      res.intercept = reg.intercept;
      res.slopeLCL = reg.slopeLCL;
      res.slopeUCL = reg.slopeUCL;
      res.interceptLCL = reg.interceptLCL;
      res.interceptUCL = reg.interceptUCL;
    }
  } else if (regressionMethod === "deming") {
    let regression = new DemingRegression(errorRatio);
    if (ciMethod === "bootstrap") {
      let regCI = new BootstrapConfidenceInterval(x, y, regression, 1000);
      res = regCI.calculate();
    } else if (ciMethod === "default") {
      let regCI = new JackknifeConfidenceInterval(x, y, regression);
      res = regCI.calculate();
    }
  } else if (regressionMethod === "wdeming") {
    let regression = new WeightedDemingRegression(errorRatio);
    if (ciMethod === "bootstrap") {
      let regCI = new BootstrapConfidenceInterval(x, y, regression, 1000);
      res = regCI.calculate();
    } else if (ciMethod === "default") {
      let regCI = new JackknifeConfidenceInterval(x, y, regression);
      res = regCI.calculate();
    }
  }
  return res;
} //doRegression

export default function Precision(props: RegressionProps) {
  const styles = useStyles();
  const regressionTypeId = useId("regression-type");

  function createBlandAltmanChart(
    xData: InputData,
    yData: InputData,
    apsAbs: number,
    apsRel: number
  ) {
    if (props.cdRangeValue === "") {
      throw new Error("Please specify the chart data output range.");
    }
    const blandAltmanChart = new ExcelBlandAltmanChart(
      xData.means,
      yData.means,
      props.differencePlotType,
      apsAbs,
      apsRel,
      props.cdRangeValue,
      props.baRangeValue
    );
    blandAltmanChart.createChart();
  } // createBlandAltmanChart

  function createRegressionChart(
    xData: InputData,
    yData: InputData,
    regressionResults: RegressionResults,
    apsAbs: number,
    apsRel: number
  ) {
    if (props.cdRangeValue === "") {
      throw new Error("Please specify the chart data output range.");
    }
    const regressionChart = new ExcelRegressionChart(
      xData.means,
      yData.means,
      regressionResults.slope,
      regressionResults.intercept,
      apsAbs,
      apsRel,
      props.cdRangeValue,
      props.scRangeValue
    );
    regressionChart.createChart();
  } // createRegressionChart

  const loadThresholds = () => {
    let thresholds: number[][] = [];
    if (props.xThreshold0 !== "" && props.yThreshold0 !== "") {
      thresholds.push([Number(props.xThreshold0), Number(props.yThreshold0)]);
      if (props.xThreshold1 !== "" && props.yThreshold1 !== "") {
        thresholds.push([Number(props.xThreshold1), Number(props.yThreshold1)]);
        if (props.xThreshold2 !== "" && props.yThreshold2 !== "") {
          thresholds.push([Number(props.xThreshold2), Number(props.yThreshold2)]);
          if (props.xThreshold3 !== "" && props.yThreshold3 !== "") {
            thresholds.push([Number(props.xThreshold3), Number(props.yThreshold3)]);
          }
        }
      }
    }
    return thresholds;
  };

  const runRegression = async () => {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      let apsAbs = -1;
      let apsRel = -1;

      try {
        // Load APS data if importAps is checked
        if (props.importApsSettings) {
          if (props.apsAbsValue === "" || props.apsRelValue === "") {
            throw new Error("Please select the APS absolute and relative ranges.");
          }
          // Load the APS data from the ranges
          const apsAbsRange = currentWorksheet.getRange(props.apsAbsValue);
          const apsRelRange = currentWorksheet.getRange(props.apsRelValue);
          apsAbsRange.load(["values"]);
          apsRelRange.load(["values"]);
          await context.sync();
          apsAbs = apsAbsRange.values[0][0];
          apsRel = apsRelRange.values[0][0];
          if (typeof apsAbs !== "number" || typeof apsRel !== "number") {
            throw new TypeError(
              `APS absolute and relative values at ${props.apsAbsValue} and ${props.apsRelValue} must be numbers.`
            );
          }
        } else {
          // if importAps is not checked, attempt to get the APS values from the input fields
          // If the input fields are empty, we will use -1 as the default value
          if (props.apsAbsValue !== "") {
            apsAbs = Number(props.apsAbsValue);
            if (isNaN(apsAbs)) {
              throw new TypeError("APS absolute value must be a number.");
            }
          } else {
            apsAbs = -1; // Default value if not provided
          }
          if (props.apsRelValue !== "") {
            apsRel = Number(props.apsRelValue);
            if (isNaN(apsRel)) {
              throw new TypeError("APS relative value must be a number.");
            }
          } else {
            apsRel = -1; // Default value if not provided
          }
        }

        //console.log(`APS absolute: ${apsAbs}`);
        //console.log(`APS relative: ${apsRel}`);
        // Load the ranges
        const xRange = currentWorksheet.getRange(props.xRangeValue);
        const yRange = currentWorksheet.getRange(props.yRangeValue);
        let outputRng = currentWorksheet.getRange(props.compOutRangeValue);
        xRange.load(["rowCount", "columnCount", "values"]);
        yRange.load(["rowCount", "columnCount", "values"]);

        await context.sync();

        // Read data and process
        let xData = processRangeData(xRange);
        let yData = processRangeData(yRange);
        let xArr = xData.means;
        let yArr = yData.means;
        if (xData.size !== yData.size) {
          throw new RangeError("X and Y ranges must have the same number of rows");
        }
        if (xData.size > 0) {
          let errRatio = DEFAULT_ERROR_RATIO;
          if (!props.useCalcErrorRatio) {
            // If we are not using the calculated error ratio, we will use the value from the input field
            // If the input field is empty, we will use the default value.
            if (props.errorRatio === "") {
              errRatio = DEFAULT_ERROR_RATIO;
            } else {
              errRatio = Number(props.errorRatio);
            }
          } else if (props.useCalcErrorRatio && xData.devsq.length > 0) {
            // If we are using the calculated error ratio, we will calculate it from the data
            // The error ratio is the coefficient of variation of the means
            // We will use the sd_y / sd_x as the error ratio for deming regression
            // and cv_y / cv_x for weighted deming regression.
            if (props.regressionType === "deming" || props.regressionType === "paba") {
              // For Deming regression, we will use the standard deviation of the y data divided by the standard deviation of the x data
              let sd_x = xData.sd;
              let sd_y = yData.sd;
              if (sd_x === 0 || sd_y === 0) {
                throw new RangeError("Standard deviation cannot be zero");
              }
              errRatio = sd_y / sd_x;
            } else if (props.regressionType === "wdeming") {
              // For Weighted Deming regression, we will use the coefficient of variation of the y data divided by the coefficient of variation of the x data
              let cv_x = xData.cv;
              let cv_y = yData.cv;
              if (cv_x === 0 || cv_y === 0) {
                throw new RangeError("Coefficient of variation cannot be zero");
              }
              errRatio = cv_y / cv_x;
            }
            // Set the error ratio input field to the calculated value
            props.setErrorRatio(errRatio.toFixed(4));
          }

          //Ensure output has two rows and four columns

          outputRng.load([`rowCount`, `columnCount`]);
          await context.sync();
          let deltaRows = 2 - outputRng.rowCount;
          let deltaCols = 4 - outputRng.columnCount;
          if (props.labelOutput) {
            deltaRows += 2; // Add an extra row for labels
            deltaCols += 1; // Add an extra column for labels
          }
          outputRng = outputRng.getResizedRange(deltaRows, deltaCols);

          // Run the regression
          let res: RegressionResults = doRegression(xArr, yArr, props.regressionType, props.ciMethod, errRatio);
          let data: any[][] = [];
          if (props.labelOutput) {
            let method = "Passing-Bablock Regression";
            let ciType = "Non-parametric CI";
            let seLabel = "SE";
            if (props.regressionType === "deming") {
              method = "Deming Regression";
              ciType = props.ciMethod === "default" ? "Jackknife CI" : "Bootstrap CI";
            } else if (props.regressionType === "wdeming") {
              method = "Weighted Deming Regression";
              ciType = props.ciMethod === "default" ? "Jackknife CI" : "Bootstrap CI";
            } else {
              //Passing-Bablock
              ciType = props.ciMethod === "default" ? "Non-parametric CI" : "Bootstrap CI";
              seLabel = "";
            }
            data = [
              [method, "", "", ciType, ""],
              ["", "Coefficents", "LCL", "UCL", seLabel],
              ["Slope", res.slope, res.slopeLCL, res.slopeUCL, res.slopeSE],
              ["Intercept", res.intercept, res.interceptLCL, res.interceptUCL, res.interceptSE],
            ];
          } else {
            data = [
              [res.slope, res.slopeLCL, res.slopeUCL, res.slopeSE],
              [res.intercept, res.interceptLCL, res.interceptUCL, res.interceptSE],
            ];
          }
          outputRng.values = data;

          // Process concordance assessment if thresholds have been specified
          const concordanceThresholds = loadThresholds();
          if (concordanceThresholds.length > 0) {
            //console.log(concordanceThresholds)
            const contingencyBuilder = new ContingencyTableBuilder(concordanceThresholds);
            const contingencyTable = contingencyBuilder.build(xData.means, yData.means);
            const concordanceCalculator = new ConcordanceCalculator(contingencyTable);
            const concordanceResults = concordanceCalculator.calculate();
            //Copy to Excel
            if (props.concordanceOutputRange === "") {
              throw Error("Please select the concordance output range.");
            } else {
              const concordanceRange = currentWorksheet.getRange(props.concordanceOutputRange);
              const contingencyRange = concordanceRange.getAbsoluteResizedRange(
                contingencyTable.table.length,
                contingencyTable.table[0].length
              );
              contingencyRange.values = contingencyTable.table;
              const offset = Math.max(8, contingencyTable.table.length + 1);
              const cohenRange = concordanceRange
                .getOffsetRange(offset, 0)
                .getAbsoluteResizedRange(6, 2);
              cohenRange.values = concordanceCalculator.formatResultsAsArray(concordanceResults);
              await context.sync();
            }
          }

          // Check if the worksheet is protected
          currentWorksheet.load("protection/protected");
          await context.sync();
          //console.log(currentWorksheet.protection.protected)
          if (currentWorksheet.protection.protected) {
            throw new Error(
              "The worksheet is protected. The charts will not display if the worksheet is protected."
            );
          }
          // Create Bland-Altman chart if requested
          if (props.baRangeValue !== "") {
            createBlandAltmanChart(xData, yData, apsAbs, apsRel);
          }
          // Create regression chart if requested
          if (props.scRangeValue !== "") {
            createRegressionChart(xData, yData, res, apsAbs, apsRel);
          }
        } else {
          throw new RangeError("Insufficient data");
        } // if size > 0
        await context.sync();
      } catch (error) {
        console.log(error);
        if (error instanceof Error) {
          props.notify("error", error.message);
        }
      }
    });
  };

  const selectRegressionType = (event: React.ChangeEvent<HTMLSelectElement>) => {
    props.setRegressionType(event.target.value);
    const errorRatioCard = document.getElementById("errorRatioCard");
    if (errorRatioCard) {
      if (event.target.value === "wdeming" || event.target.value === "deming") {
        errorRatioCard.className = styles.field;
        //console.log("Showing error ratio card");
      } else {
        //console.log("Hiding error ratio card");
        errorRatioCard.className = styles.mvwhidden;
      }
    }
  };

  const selectCiMethod = (event: React.ChangeEvent<HTMLSelectElement>) => {
    props.setCiMethod(event.target.value);
  };

  const selectDiffferencePlotType = (event: React.ChangeEvent<HTMLSelectElement>) => {
    props.setDifferencePlotType(event.target.value);
  };

  const toggleLabelOutput = (event: React.ChangeEvent<HTMLInputElement>) => {
    props.setLabelOutput(event.target.checked);
  };

  const toggleUseCalcErrorRatio = (event: React.ChangeEvent<HTMLInputElement>) => {
    props.setUseCalcErrorRatio(event.target.checked);
  };

  const handleErrorRatioChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    props.setErrorRatio(event.target.value);
  };

  const toggleImportApsSettings = (event: React.ChangeEvent<HTMLInputElement>) => {
    props.setImportApsSettings(event.target.checked);
  };

  return (
    <div className="container">
      <Accordion collapsible={true}>
        <AccordionItem value="10">
          <AccordionHeader>{props.uitext["h_performance_specs"]}</AccordionHeader>
          <AccordionPanel>
            <div className={styles.specificationsPanel}>
              <Checkbox
                label={props.uitext["lbl_import_aps"]}
                checked={props.importApsSettings}
                onChange={toggleImportApsSettings}
                className={styles.field}
              />
              <RangeInput
                label={props.uitext["lbl_aps_abs"]}
                rangeValue={props.apsAbsValue}
                setRangeValue={props.setApsAbsValue}
                validationMessage={props.uitext["msg_aps_abs"]}
                tooltipContent={props.uitext["tip_aps_abs"]}
                allowNumbers={true}
                uitext={props.uitext}
              />
              <RangeInput
                label={props.uitext["lbl_aps_rel"]}
                rangeValue={props.apsRelValue}
                setRangeValue={props.setApsRelValue}
                validationMessage={props.uitext["msg_aps_rel"]}
                tooltipContent={props.uitext["tip_aps_rel"]}
                allowNumbers={true}
                uitext={props.uitext}
              />
            </div>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
      <RangeInput
        label={props.uitext["lbl_x_range"]}
        rangeValue={props.xRangeValue}
        setRangeValue={props.setXRangeValue}
        validationMessage={props.uitext["msg_x_range"]}
        uitext={props.uitext}
      />
      <RangeInput
        label={props.uitext["lbl_y_range"]}
        rangeValue={props.yRangeValue}
        setRangeValue={props.setYRangeValue}
        validationMessage={props.uitext["msg_y_range"]}
        uitext={props.uitext}
      />
      <RangeInput
        label={props.uitext["lbl_output_range"]}
        rangeValue={props.compOutRangeValue}
        setRangeValue={props.setCompOutRangeValue}
        validationMessage={props.uitext["msg_output_range"]}
        tooltipContent={props.uitext["tip_output_range"]}
        uitext={props.uitext}
      />
      <div className={styles.field}>
        <label htmlFor={regressionTypeId}>{props.uitext["lbl_reg_method"]}</label>
        <Select
          value={props.regressionType}
          id={regressionTypeId}
          className={styles.selection}
          onChange={selectRegressionType}
        >
          <option value="paba">{props.uitext["opt_paba"]}</option>
          <option value="deming">{props.uitext["opt_dem"]}</option>
          <option value="wdeming">{props.uitext["opt_wdem"]}</option>
        </Select>
      </div>
      <div className={styles.field}>
        <label htmlFor="ci-method">{props.uitext["lbl_ci_method"]}</label>
        <Select
          value={props.ciMethod}
          id="ci-method"
          className={styles.selection}
          onChange={selectCiMethod}
        >
          <option value="default">{props.uitext["opt_default"]}</option>
          <option value="bootstrap">{props.uitext["opt_bootstrap"]}</option>
        </Select>
      </div>
      <div id="errorRatioCard" className={styles.mvwhidden}>
        <Checkbox
          label={props.uitext["lbl_use_calc_ratio"]}
          checked={props.useCalcErrorRatio}
          onChange={toggleUseCalcErrorRatio}
        />
        <Field label={props.uitext["lbl_error_ratio"]} className={styles.field}>
          <Input value={props.errorRatio} type="number" onChange={handleErrorRatioChange} />
        </Field>
      </div>
      <div className={styles.field}>
        <label htmlFor="lbl-output">{props.uitext["lbl_output_labels"]}</label>
        <Checkbox id="lbl-output" checked={props.labelOutput} onChange={toggleLabelOutput} />
      </div>
      <RangeInput
        label={props.uitext["lbl_ba_range"]}
        rangeValue={props.baRangeValue}
        setRangeValue={props.setBaRangeValue}
        validationMessage={props.uitext["msg_ba_range"]}
        tooltipContent={props.uitext["tip_ba_range"]}
        uitext={props.uitext}
      />
      <div className={styles.field}>
        <label htmlFor="difference-type">{props.uitext["lbl_diff_type"]}</label>
        <Select
          value={props.differencePlotType}
          id="difference-type"
          className={styles.selection}
          onChange={selectDiffferencePlotType}
        >
          <option value="abs">{props.uitext["opt_abs"]}</option>
          <option value="rel">{props.uitext["opt_rel"]}</option>
        </Select>
      </div>
      <RangeInput
        label={props.uitext["lbl_sc_range"]}
        rangeValue={props.scRangeValue}
        setRangeValue={props.setScRangeValue}
        validationMessage={props.uitext["msg_sc_range"]}
        tooltipContent={props.uitext["tip_sc_range"]}
        uitext={props.uitext}
      />
      <RangeInput
        label="Chart Data Output Range"
        rangeValue={props.cdRangeValue}
        setRangeValue={props.setCdRangeValue}
        validationMessage="Must be a valid Excel cell reference. e.g., H1"
        tooltipContent="Top left cell where data used to construct the charts is to be saved."
        uitext={props.uitext}
      />
      <div className={styles.field}>
        <Button appearance="primary" onClick={runRegression}>
          {props.uitext["btn_run"]}
        </Button>
      </div>
      <Accordion collapsible={true}>
        <AccordionItem value="20">
          <AccordionHeader>{props.uitext["h_cohen_kappa"]}</AccordionHeader>
          <AccordionPanel>
            <div className={styles.field}>
              <div>{props.uitext["inf_cohen_kappa"]}</div>
              <ConcordanceThreshold
                xThreshold0={props.xThreshold0}
                setXThreshold0={props.setXThreshold0}
                xThreshold1={props.xThreshold1}
                setXThreshold1={props.setXThreshold1}
                xThreshold2={props.xThreshold2}
                setXThreshold2={props.setXThreshold2}
                xThreshold3={props.xThreshold3}
                setXThreshold3={props.setXThreshold3}
                yThreshold0={props.yThreshold0}
                setYThreshold0={props.setYThreshold0}
                yThreshold1={props.yThreshold1}
                setYThreshold1={props.setYThreshold1}
                yThreshold2={props.yThreshold2}
                setYThreshold2={props.setYThreshold2}
                yThreshold3={props.yThreshold3}
                setYThreshold3={props.setYThreshold3}
              />
            </div>

            <RangeInput
              label={props.uitext["lbl_concordance_range"]}
              rangeValue={props.concordanceOutputRange}
              setRangeValue={props.setConcordanceOutputRange}
              tooltipContent={props.uitext["tip_concordance_range"]}
              uitext={props.uitext}
            />
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </div>
  );
}