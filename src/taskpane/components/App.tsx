/*
* Main component for the add-in. 
*
* The Bland-Altman and Regression analysis was developed first so those 
* components appear hear. To reduce the complexity and size of this file 
* the other functions were moved to separate subcomponents but the properties
* for the complete add-in are all defined here as they are in many cases
* required to be used in more than one subcomponent.
* 
* Author: Douglas Chesher
*
* Created: October 2025.
*/
import * as React from "react";
import { useState } from "react";
import RangeInput from "./RangeInput";
import ConcordanceThreshold from "./ConcordanceThreshold";
import PrecisionLayout from "./PrecisionLayout";
import Qualitative from "./Qualitative";
import BoxCox from "./BoxCox";
import RefInt from "./RefInt";

import {
  makeStyles, useId,
  Accordion, AccordionItem, AccordionHeader, AccordionPanel,
  Select, Button, Checkbox, CheckboxProps,
  Field, Input,
  Link,
  Card,
  Tab, TabList,
  Toaster,
  useToastController,
  ToastTitle,
  Toast,
  ToastTrigger,
  ToastBody,
  Tooltip,
} from "@fluentui/react-components";

import type {
  SelectTabData,
  SelectTabEvent,
  TabValue,
} from "@fluentui/react-components";

import {
  JackknifeConfidenceInterval,
  BootstrapConfidenceInterval,
  PassingBablokRegression,
  DemingRegression,
  WeightedDemingRegression,
  DEFAULT_ERROR_RATIO,
} from "../../regression";

import {
  ExcelBlandAltmanChart, ExcelRegressionChart
} from "../../charts";
import {
  ContingencyTableBuilder, QualitativeContengencyTableBuilder,
  ConcordanceCalculator
} from "../../concordance";
import Precision from "./Precision";

interface AppProps {
  title: string;
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
  }
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
    interceptUCL: NaN
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

/* Main application entry
*/
const App: React.FC<AppProps> = () => {
  const [apsAbsValue, setApsAbsValue] = useState<string>("");
  const [apsRelValue, setApsRelValue] = useState<string>("");
  const [xRangeValue, setXRangeValue] = useState<string>("");
  const [yRangeValue, setYRangeValue] = useState<string>("");
  const [compOutRangeValue, setCompOutRangeValue] = useState<string>("");
  const [baRangeValue, setBaRangeValue] = useState<string>(""); // Bland-Altman output range
  const [scRangeValue, setScRangeValue] = useState<string>(""); // Scatter chart output range
  const [cdRangeValue, setCdRangeValue] = useState<string>(""); // Chart data output range
  const [errorRatio, setErrorRatio] = useState<string>("1.0");
  const [regressionType, setRegressionType] = useState<string>("paba") // Default to "paba" (Passing-Bablok)
  const [useCalcErrorRatio, setUseCalcErrorRatio] = useState<boolean>(false); // Default to false
  const [importApsSettings, setImportApsSettings] = useState<boolean>(false); // Default to false
  const [labelOutput, setLabelOutput] = useState<boolean>(true); // Default to true
  const [differencePlotType, setDifferencePlotType] = useState<string>("rel"); // Default to "rel"
  const [concordanceOutputRange, setConcordanceOutputRange] = useState<string>("");
  const [ciMethod, setCiMethod] = useState<string>("default") // Default to "bootstrap"
  const [xThreshold0, setXThreshold0] = React.useState<string>("");
  const [xThreshold1, setXThreshold1] = React.useState<string>("");
  const [xThreshold2, setXThreshold2] = React.useState<string>("");
  const [xThreshold3, setXThreshold3] = React.useState<string>("");
  const [yThreshold0, setYThreshold0] = React.useState<string>("");
  const [yThreshold1, setYThreshold1] = React.useState<string>("");
  const [yThreshold2, setYThreshold2] = React.useState<string>("");
  const [yThreshold3, setYThreshold3] = React.useState<string>("");

  const [factorARangeValue, setFactorARangeValue] = React.useState<string>("");
  const [factorBRangeValue, setFactorBRangeValue] = React.useState<string>("");
  const [resultsRangeValue, setResultsRangeValue] = React.useState<string>("");
  const [precOutRangeValue, setPrecOutRangeValue] = React.useState<string>("");

  const [qualXRangeValue, setQualXRangeValue] = React.useState<string>("");
  const [qualYRangeValue, setQualYRangeValue] = React.useState<string>("");
  const [qualOutputRangeValue, setQualOutputRangeValue] = React.useState<string>("");

  const [layoutSheet, setLayoutSheet] = React.useState<string>("Layout");
  const [importLayoutChecked, setImportLayoutChecked] = React.useState<CheckboxProps["checked"]>(false);

  const [selectedTabValue, setSelectedTabValue] = React.useState<TabValue>("evaluation");

  const regressionTypeId = useId("regression-type");
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  const styles = useStyles();

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    if (event.isTrusted) {
      setSelectedTabValue(data.value);
    }
  };

  /* Configure default values for the input fields */
  const setDefaultValues = () => {
    if (importLayoutChecked) {
      loadWorkbookDefaults();
    } else {
      // Set explicitly for now but in the future we may load these from a workbook or a settings file.
      try {
        // APS
        setImportApsSettings(true);
        setApsAbsValue("J4");
        setApsRelValue("J6");

        // Method Comparison
        setXRangeValue("O61:P160");
        setYRangeValue("R61:S160");
        setCompOutRangeValue("C167");
        setRegressionType("paba");
        setCiMethod("default");
        setUseCalcErrorRatio(false);
        setLabelOutput(true);
        setErrorRatio("1.0");
        setDifferencePlotType("rel");
        setBaRangeValue("N162:U175");
        setScRangeValue("N177:U190");
        setCdRangeValue("AP61");
        setConcordanceOutputRange("C176");

        // Imprecision

        setFactorARangeValue("B300:B324");
        setFactorBRangeValue("C300:C324");
        setResultsRangeValue("D300:G324");
        setPrecOutRangeValue("W299");

        // concordance of qualitative data
        setQualXRangeValue("E61:E160");
        setQualYRangeValue("H61:H160");
        setQualOutputRangeValue("C176");
        notify("success", "Ranges have been successfully updated");
      } catch (error) {
        notify("error", error.message);
      }
    }
  } // setDefaultValues

  // Read the default cell addresses from and excel workbook.
  const loadWorkbookDefaults = async () => {
    await Excel.run(async (context) => {
      try {
        if (layoutSheet === "") throw new Error("Please specify the name of the layout sheet.");
        const worksheet = context.workbook.worksheets.getItem(layoutSheet);
        const address = "A2:B23";
        const range = worksheet.getRange(address);
        range.load(["values", "rowCount", "columnCount"]);
        await context.sync();

        for (let row = 0; row < range.rowCount; row++) {
          const attribute = range.values[row][0];
          const value = range.values[row][1];
          switch (attribute) {
            case "import-aps":
              setImportApsSettings(Boolean(value));
              break;
            case "aps-abs":
              setApsAbsValue(value);
              break;
            case "aps-rel":
              setApsRelValue(value);
              break;
            case "x-range":
              setXRangeValue(value);
              break;
            case "y-range":
              setYRangeValue(value);
              break;
            case "output-range":
              setCompOutRangeValue(value);
              break;
            case "regression-method":
              setRegressionType(value);
              break;
            case "confidence-interval-method":
              setCiMethod(value);
              break;
            case "use-calculated-error-ratio":
              setUseCalcErrorRatio(Boolean(value));
              break;
            case "error-ratio":
              setErrorRatio(value);
              break;
            case "bland-altman-range":
              setBaRangeValue(value);
              break;
            case "scatter-plot-range":
              setScRangeValue(value);
              break;
            case "output-labels":
              setLabelOutput(Boolean(value));
              break;
            case "chart-data-range":
              setCdRangeValue(value);
              break;
            case "p-days-range":
              setFactorARangeValue(value);
              break;
            case "p-runs-range":
              setFactorBRangeValue(value);
              break;
            case "p-results-range":
              setResultsRangeValue(value);
              break;
            case "p-output-range":
              setPrecOutRangeValue(value);
              break;
            case "concordance-output":
              setConcordanceOutputRange(value);
              break;
            case "qx-range":
              setQualXRangeValue(value);
              break;
            case "qy-range":
              setQualYRangeValue(value);
              break;
            case "q-output-range":
              setQualOutputRangeValue(value);
              break;
          } //switch
        }
        notify("success", "Ranges have been successfully updated");
      } catch (error) {
        notify("error", error.message)
      }
    });
  } //readLayout

  /* Display messages
  */
  const notify = (intent: string, message: string) => {
    switch (intent) {
      case "error":
        dispatchToast(
          <Toast appearance="inverted">
            <ToastTitle
              action={
                <ToastTrigger>
                  <Link>Dismiss</Link>
                </ToastTrigger>
              }
            >Error</ToastTitle>
            <ToastBody>{message}</ToastBody>
          </Toast>,
          { intent: "error" }
        );
        break;
      case "warning":
        dispatchToast(
          <Toast appearance="inverted">
            <ToastTitle
              action={
                <ToastTrigger>
                  <Link>Dismiss</Link>
                </ToastTrigger>
              }
            >Warning</ToastTitle>
            <ToastBody>{message}</ToastBody>
          </Toast>,
          { intent: "warning" }
        );
        break;
      default:
        dispatchToast(
          <Toast appearance="inverted">
            <ToastTitle
              action={
                <ToastTrigger>
                  <Link>Dismiss</Link>
                </ToastTrigger>
              }
            >Info</ToastTitle>
            <ToastBody>{message}</ToastBody>
          </Toast>,
          { intent: "info" }
        );
    }
  }

  function createBlandAltmanChart(xData: InputData, yData: InputData, apsAbs: number, apsRel: number) {
    if (cdRangeValue === "") {
      throw new Error("Please specify the chart data output range.");
    }
    const blandAltmanChart = new ExcelBlandAltmanChart(
      xData.means,
      yData.means,
      differencePlotType,
      apsAbs,
      apsRel,
      cdRangeValue,
      baRangeValue
    );
    blandAltmanChart.createChart();
  } // createBlandAltmanChart

  function createRegressionChart(xData: InputData, yData: InputData, regressionResults: RegressionResults, apsAbs: number, apsRel: number) {
    if (cdRangeValue === "") {
      throw new Error("Please specify the chart data output range.");
    }
    const regressionChart = new ExcelRegressionChart(
      xData.means,
      yData.means,
      regressionResults.slope,
      regressionResults.intercept,
      apsAbs,
      apsRel,
      cdRangeValue,
      scRangeValue
    );
    regressionChart.createChart();
  } // createRegressionChart

  const loadThresholds = () => {
    let thresholds: number[][] = [];
    if (xThreshold0 !== "" && yThreshold0 !== "") {
      thresholds.push([Number(xThreshold0), Number(yThreshold0)]);
      if (xThreshold1 !== "" && yThreshold1 !== "") {
        thresholds.push([Number(xThreshold1), Number(yThreshold1)]);
        if (xThreshold2 !== "" && yThreshold2 !== "") {
          thresholds.push([Number(xThreshold2), Number(yThreshold2)]);
          if (xThreshold3 !== "" && yThreshold3 !== "") {
            thresholds.push([Number(xThreshold3), Number(yThreshold3)]);
          }
        }
      }
    }
    return thresholds;
  }

  const runRegression = async () => {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      let apsAbs = -1;
      let apsRel = -1;

      try {
        // Load APS data if importAps is checked
        if (importApsSettings) {
          if (apsAbsValue === "" || apsRelValue === "") {
            throw new Error("Please select the APS absolute and relative ranges.");
          }
          // Load the APS data from the ranges
          const apsAbsRange = currentWorksheet.getRange(apsAbsValue);
          const apsRelRange = currentWorksheet.getRange(apsRelValue);
          apsAbsRange.load(["values"]);
          apsRelRange.load(["values"]);
          await context.sync();
          apsAbs = apsAbsRange.values[0][0];
          apsRel = apsRelRange.values[0][0];
          if (typeof apsAbs !== "number" || typeof apsRel !== "number") {
            throw new TypeError(
              `APS absolute and relative values at ${apsAbsValue} and ${apsRelValue} must be numbers.`
            );
          }
        } else {
          // if importAps is not checked, attempt to get the APS values from the input fields
          // If the input fields are empty, we will use -1 as the default value
          if (apsAbsValue !== "") {
            apsAbs = Number(apsAbsValue);
            if (isNaN(apsAbs)) {
              throw new TypeError("APS absolute value must be a number.");
            }
          } else {
            apsAbs = -1; // Default value if not provided
          }
          if (apsRelValue !== "") {
            apsRel = Number(apsRelValue);
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
        const xRange = currentWorksheet.getRange(xRangeValue);
        const yRange = currentWorksheet.getRange(yRangeValue);
        let outputRng = currentWorksheet.getRange(compOutRangeValue);
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
          if (!useCalcErrorRatio) {
            // If we are not using the calculated error ratio, we will use the value from the input field
            // If the input field is empty, we will use the default value.
            if (errorRatio === "") {
              errRatio = DEFAULT_ERROR_RATIO;
            } else {
              errRatio = Number(errorRatio);
            }
          } else if (useCalcErrorRatio && xData.devsq.length > 0) {
            // If we are using the calculated error ratio, we will calculate it from the data
            // The error ratio is the coefficient of variation of the means
            // We will use the sd_y / sd_x as the error ratio for deming regression
            // and cv_y / cv_x for weighted deming regression.
            if (regressionType === "deming" || regressionType === "paba") {
              // For Deming regression, we will use the standard deviation of the y data divided by the standard deviation of the x data
              let sd_x = xData.sd;
              let sd_y = yData.sd;
              if (sd_x === 0 || sd_y === 0) {
                throw new RangeError("Standard deviation cannot be zero");
              }
              errRatio = sd_y / sd_x;
            } else if (regressionType === "wdeming") {
              // For Weighted Deming regression, we will use the coefficient of variation of the y data divided by the coefficient of variation of the x data
              let cv_x = xData.cv;
              let cv_y = yData.cv;
              if (cv_x === 0 || cv_y === 0) {
                throw new RangeError("Coefficient of variation cannot be zero");
              }
              errRatio = cv_y / cv_x;
            }
            // Set the error ratio input field to the calculated value
            setErrorRatio(errRatio.toFixed(4));
          }

          //Ensure output has two rows and four columns

          outputRng.load([`rowCount`, `columnCount`]);
          await context.sync();
          let deltaRows = 2 - outputRng.rowCount;
          let deltaCols = 4 - outputRng.columnCount;
          if (labelOutput) {
            deltaRows += 2; // Add an extra row for labels
            deltaCols += 1; // Add an extra column for labels
          }
          outputRng = outputRng.getResizedRange(deltaRows, deltaCols);

          // Run the regression
          let res: RegressionResults = doRegression(
            xArr,
            yArr,
            regressionType,
            ciMethod,
            errRatio
          );
          let data: any[][] = [];
          if (labelOutput) {
            let method = "Passing-Bablock Regression";
            let ciType = "Non-parametric CI";
            let seLabel = "SE";
            if (regressionType === "deming") {
              method = "Deming Regression";
              ciType = ciMethod === "default" ? "Jackknife CI" : "Bootstrap CI";
            } else if (regressionType === "wdeming") {
              method = "Weighted Deming Regression";
              ciType = ciMethod === "default" ? "Jackknife CI" : "Bootstrap CI";
            } else {
              //Passing-Bablock
              ciType = ciMethod === "default" ? "Non-parametric CI" : "Bootstrap CI";
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
            if (concordanceOutputRange === "") {
              throw Error("Please select the concordance output range.");
            } else {
              const concordanceRange = currentWorksheet.getRange(concordanceOutputRange);
              const contingencyRange = concordanceRange.getAbsoluteResizedRange(contingencyTable.table.length, contingencyTable.table[0].length);
              contingencyRange.values = contingencyTable.table;
              const offset = Math.max(8, contingencyTable.table.length + 1)
              const cohenRange = concordanceRange.getOffsetRange(offset, 0).getAbsoluteResizedRange(6, 2);
              cohenRange.values = concordanceCalculator.formatResultsAsArray(concordanceResults);
              await context.sync();
            }
          }

          // Check if the worksheet is protected
          currentWorksheet.load("protection/protected")
          await context.sync();
          //console.log(currentWorksheet.protection.protected)
          if (currentWorksheet.protection.protected) {
            throw new Error("The worksheet is protected. The charts will not display if the worksheet is protected.")
          }
          // Create Bland-Altman chart if requested
          if (baRangeValue !== "") {
            createBlandAltmanChart(xData, yData, apsAbs, apsRel);
          }
          // Create regression chart if requested
          if (scRangeValue !== "") {
            createRegressionChart(xData, yData, res, apsAbs, apsRel);
          }

        } else {
          throw new RangeError("Insufficient data");
        } // if size > 0
        await context.sync();
      } catch (error) {
        console.log(error);
        notify("error", error.message);
      }
    });
  }

  const selectRegressionType = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setRegressionType(event.target.value);
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
  }

  const selectCiMethod = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setCiMethod(event.target.value);
  }

  const selectDiffferencePlotType = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setDifferencePlotType(event.target.value);
  }

  const toggleLabelOutput = (event: React.ChangeEvent<HTMLInputElement>) => {
    setLabelOutput(event.target.checked);
  }

  const toggleUseCalcErrorRatio = (event: React.ChangeEvent<HTMLInputElement>) => {
    setUseCalcErrorRatio(event.target.checked);
  }

  const toggleImportApsSettings = (event: React.ChangeEvent<HTMLInputElement>) => {
    setImportApsSettings(event.target.checked);
  }
  const handleErrorRatioChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setErrorRatio(event.target.value);
  }

  const Evaluation = React.memo(() => {
    return (
      <div role="tabpanel" aria-labelledby="Evaluation">
        <Accordion defaultOpenItems="2">
          <AccordionItem value="1">
            <AccordionHeader>Performance Specifications</AccordionHeader>
            <AccordionPanel>
              <Card>
                <Checkbox
                  label="Import APS settings from worksheet"
                  checked={importApsSettings}
                  onChange={toggleImportApsSettings}
                  className={styles.field}
                />
                <RangeInput label="APS Absolute"
                  rangeValue={apsAbsValue}
                  setRangeValue={setApsAbsValue}
                  validationMessage="Must be a number or valid Excel cell reference. e.g., D1"
                  tooltipContent="Excel cell reference or APS absolute value"
                  allowNumbers={true} />
                <RangeInput label="APS Relative"
                  rangeValue={apsRelValue}
                  setRangeValue={setApsRelValue}
                  validationMessage="Must be a number or valid Excel cell reference. e.g., D2"
                  tooltipContent="Excel cell reference or APS relative value as a fraction. e.g., enter 0.1 for 10%"
                  allowNumbers={true} />
              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="2">
            <AccordionHeader>Comparison</AccordionHeader>
            <AccordionPanel>

              <Card>
                <RangeInput label="X Range"
                  rangeValue={xRangeValue}
                  setRangeValue={setXRangeValue}
                  validationMessage="Must be a valid Excel range. e.g., A2:A21" />
                <RangeInput label="Y Range"
                  rangeValue={yRangeValue}
                  setRangeValue={setYRangeValue}
                  validationMessage="Must be a valid Excel range. e.g., B2:B21" />
                <RangeInput label="Output Range"
                  rangeValue={compOutRangeValue}
                  setRangeValue={setCompOutRangeValue}
                  validationMessage="Must be a valid Excel cell reference. e.g., E1"
                  tooltipContent="Top left cell where the regression statistics are to be saved." />
                <div className={styles.field}>
                  <label htmlFor={regressionTypeId}>Regression Method:</label>
                  <Select value={regressionType} id={regressionTypeId}
                    className={styles.selection} onChange={selectRegressionType} >
                    <option value="paba">Passing-Bablok</option>
                    <option value="deming">Deming</option>
                    <option value="wdeming">Weighted Deming</option>
                  </Select>
                </div>
                <div className={styles.field}>
                  <label htmlFor="ci-method">Confidence Interval Method:</label>
                  <Select value={ciMethod} id="ci-method"
                    className={styles.selection} onChange={selectCiMethod} >
                    <option value="default">Default</option>
                    <option value="bootstrap">Bootstrap</option>
                  </Select>
                </div>
                <div id="errorRatioCard" className={styles.mvwhidden}>
                  <Checkbox
                    label="Use Calculated Error Ratio"
                    checked={useCalcErrorRatio}
                    onChange={toggleUseCalcErrorRatio}
                  />
                  <Field label="Error Ratio" className={styles.field}>
                    <Input value={errorRatio} type="number" onChange={handleErrorRatioChange} />
                  </Field>
                </div>
                <div className={styles.field}>
                  <label htmlFor="lbl-output">Output Labels:</label>
                  <Checkbox
                    id="lbl-output"
                    checked={labelOutput}
                    onChange={toggleLabelOutput}
                  />
                </div>
                <RangeInput label="Bland-Altman Output Range"
                  rangeValue={baRangeValue}
                  setRangeValue={setBaRangeValue}
                  validationMessage="Must be a valid Excel range. e.g., C2:I15"
                  tooltipContent="Cell range over which the Bland-Altman chart will be displayed" />
                <div className={styles.field}>
                  <label htmlFor="difference-type">Difference Plot Type:</label>
                  <Select value={differencePlotType} id="difference-type"
                    className={styles.selection} onChange={selectDiffferencePlotType} >
                    <option value="abs">Absolute</option>
                    <option value="rel">Relative</option>
                  </Select>
                </div>
                <RangeInput label="Scatter Chart Output Range"
                  rangeValue={scRangeValue}
                  setRangeValue={setScRangeValue}
                  validationMessage="Must be a valid Excel range. e.g., C16:I19"
                  tooltipContent="Cell range over which the scatter chart will be displayed" />
                <RangeInput label="Chart Data Output Range"
                  rangeValue={cdRangeValue}
                  setRangeValue={setCdRangeValue}
                  validationMessage="Must be a valid Excel cell reference. e.g., H1"
                  tooltipContent="Top left cell where data used to construct the charts is to be saved." />
                <div className={styles.field}>
                  <Button appearance="primary"
                    onClick={runRegression}>Run Regression</Button>
                </div>
                <Accordion collapsible={true}>
                  <AccordionItem value="3">
                    <AccordionHeader>Cohen's Kappa/</AccordionHeader>
                    <AccordionPanel>
                      <div className={styles.field}>
                        <div>
                          Specify cutoffs for each method to assess concordance. Each method must have
                          the same number of cutoffs.
                        </div>
                        <ConcordanceThreshold
                          xThreshold0={xThreshold0}
                          setXThreshold0={setXThreshold0}
                          xThreshold1={xThreshold1}
                          setXThreshold1={setXThreshold1}
                          xThreshold2={xThreshold2}
                          setXThreshold2={setXThreshold2}
                          xThreshold3={xThreshold3}
                          setXThreshold3={setXThreshold3}
                          yThreshold0={yThreshold0}
                          setYThreshold0={setYThreshold0}
                          yThreshold1={yThreshold1}
                          setYThreshold1={setYThreshold1}
                          yThreshold2={yThreshold2}
                          setYThreshold2={setYThreshold2}
                          yThreshold3={yThreshold3}
                          setYThreshold3={setYThreshold3} />
                      </div>

                      <RangeInput label="Concordance Output Range"
                        rangeValue={concordanceOutputRange}
                        setRangeValue={setConcordanceOutputRange}
                        tooltipContent="Top left cell where the concordance metrics will be saved." />
                    </AccordionPanel>
                  </AccordionItem>
                </Accordion>
              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="4">
            <AccordionHeader>Precision</AccordionHeader>
            <AccordionPanel>
              <Card>
                <Precision
                  notify={notify}
                  factorARangeValue={factorARangeValue}
                  factorBRangeValue={factorBRangeValue}
                  resultsRangeValue={resultsRangeValue}
                  precOutRangeValue={precOutRangeValue}
                  setFactorARangeValue={setFactorARangeValue}
                  setFactorBRangeValue={setFactorBRangeValue}
                  setResultsRangeValue={setResultsRangeValue}
                  setPrecOutRangeValue={setPrecOutRangeValue}
                />
              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="5">
            <AccordionHeader>Precision Layout</AccordionHeader>
            <AccordionPanel>
              <Card>
                <PrecisionLayout
                  notify={notify}
                  factorARangeValue={factorARangeValue} setFactorARangeValue={setFactorARangeValue}
                  factorBRangeValue={factorBRangeValue} setFactorBRangeValue={setFactorBRangeValue}
                  resultsRangeValue={resultsRangeValue} setResultsRangeValue={setResultsRangeValue}
                />
              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="6">
            <AccordionHeader>Qualitative Comparison</AccordionHeader>
            <AccordionPanel>
              <Card>
                <Qualitative notify={notify}
                  qualXRangeValue={qualXRangeValue}
                  qualYRangeValue={qualYRangeValue}
                  qualOutputRangeValue={qualOutputRangeValue}
                  setQualXRangeValue={setQualXRangeValue}
                  setQualYRangeValue={setQualYRangeValue}
                  setQualOutputRangeValue={setQualOutputRangeValue} />
              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="7">
            <AccordionHeader>Load Range Defaults</AccordionHeader>
            <AccordionPanel>
              <Card>
                <div className={styles.container}>
                  <div className={styles.field}>
                    <p>
                      This add in was developed for use with the associated Method Verification Workbook.
                      The default cell ranges used by the workbook can be populated by clicking the button
                      below.
                    </p>
                    <Button appearance="primary"
                      onClick={setDefaultValues}>Load Defaults</Button>

                    <p>
                      Alternatively, the defaults for a user defined template can defined in a worksheet
                      and loaded from that page. The definitions need to be supplied in two columns from
                      cells A2:B20. The first column should give the name of the attribute as it appears
                      below. The second column must contain the values to be copied. A1 and B1 can be
                      column headers but they are not read.
                    </p>
                    <Checkbox label={"Import from worksheet"} checked={importLayoutChecked} onChange={(_ev, data) => setImportLayoutChecked(data.checked)} />
                    <Field label="Layout Sheet" className={styles.field}>
                      <Tooltip relationship="description" content="Name of the worksheet specifying the default values.">
                        <Input value={layoutSheet} onChange={(ev) => setLayoutSheet(ev.target.value)} type="text" />
                      </Tooltip>
                    </Field>
                    <div className={styles.field}>
                      <table>
                        <thead>
                          <tr>
                            <th>Attribute</th>
                            <th>Value</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td>import-aps</td>
                            <td>TRUE</td>
                          </tr>
                          <tr>
                            <td>aps-abs</td>
                            <td>J4</td>
                          </tr>
                          <tr>
                            <td>aps-rel</td>
                            <td>J6</td>
                          </tr>
                          <tr>
                            <td>x-range</td>
                            <td>O61:P160</td>
                          </tr>
                          <tr>
                            <td>y-range</td>
                            <td>R61:S160</td>
                          </tr>
                          <tr>
                            <td>output-range</td>
                            <td>C167</td>
                          </tr>
                          <tr>
                            <td>regression-method</td>
                            <td>paba, deming, or wdeming</td>
                          </tr>
                          <tr>
                            <td>confidence-interval-method</td>
                            <td>default or bootstrap</td>
                          </tr>
                          <tr>
                            <td>use-calculated-error-ratio</td>
                            <td>FALSE</td>
                          </tr>
                          <tr>
                            <td>error-ratio</td>
                            <td>1.0</td>
                          </tr>
                          <tr>
                            <td>bland-altman-range</td>
                            <td>N162:U175</td>
                          </tr>
                          <tr>
                            <td>scatter-plot-range</td>
                            <td>N177:U190</td>
                          </tr>
                          <tr>
                            <td>output-labels</td>
                            <td>TRUE</td>
                          </tr>
                          <tr>
                            <td>chart-data-range</td>
                            <td>AP61</td>
                          </tr>
                          <tr>
                            <td>p-days-range</td>
                            <td>B300:B324</td>
                          </tr>
                          <tr>
                            <td>p-runs-range</td>
                            <td>C300:C324</td>
                          </tr>
                          <tr>
                            <td>p-results-range</td>
                            <td>D300:G324</td>
                          </tr>
                          <tr>
                            <td>p-output-range</td>
                            <td>W299</td>
                          </tr>
                          <tr>
                            <td>concordance-output</td>
                            <td>C176</td>
                          </tr>
                          <tr>
                            <td>qx-range</td>
                            <td>E61:E160</td>
                          </tr>
                          <tr>
                            <td>qy-range</td>
                            <td>H61:H160</td>
                          </tr>
                          <tr>
                            <td>q-output-range</td>
                            <td>C176</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </Card>
            </AccordionPanel>
          </AccordionItem>
        </Accordion>
      </div>
    );
  });

  const Utilities = React.memo(() => {
    return (
      <div role="tabpanel" aria-labelledby="Utilities">
        <Accordion defaultOpenItems="1">
          <AccordionItem value="1">
            <AccordionHeader>Reference Interval</AccordionHeader>
            <AccordionPanel>
              <Card>
                <RefInt notify={notify} />

              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="2">
            <AccordionHeader>Box Cox</AccordionHeader>
            <AccordionPanel>
              <Card>
                <BoxCox notify={notify} />
              </Card>
            </AccordionPanel>
          </AccordionItem>
        </Accordion>
      </div>
    );
  });

  const Help = React.memo(() => {
    return (
      <div role="tabpanel" aria-labelledby="Help">
        <Card>
          <h3>About</h3><p>The purpose of this add-in is to provide some statistical procedures that are
            necessary for assessing clinical laboratory methods but are not easy to implement
            using builtin Excel spreadsheet functions. It is not meant to be a comprehensive
            statistical analysis tool. The current interation allows the user to perform linear
            regression techniques including Passing-Bablok, Deming, and Weighted Deming.
            Procedures are also provided to allow users to analyse variance components as described
            in CLSI EP15 and EP05. <span className={styles.label}>Instructions for use</span> can be found <a
              href="https://metools.chesher.id.au/help/index.html" target="help">here</a>.
          </p>
          <Accordion>
            <AccordionItem value="1">
              <AccordionHeader>Source</AccordionHeader>
              <AccordionPanel>
                <p>
                  Source code is available on GitHub at <a
                    href="https://github.com/dche658/method-eval-tools">
                    https://github.com/dche658/method-eval-tools </a
                  >.
                </p>
                <p>
                  THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR
                  IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE,
                  MERCHANTABILITY, OR NON-INFRINGEMENT.
                </p>
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
        </Card>
      </div>
    );
  });

  return (
    <div className={styles.root}>
      <Toaster toasterId={toasterId} position="top" timeout={5000} pauseOnHover />
      <TabList selectedValue={selectedTabValue} onTabSelect={onTabSelect}>
        <Tab id="Evaluation" value="evaluation">
          Evaluation
        </Tab>
        <Tab id="Utilities" value="utilities">
          Utilities
        </Tab>
        <Tab id="Help" value="help">
          Help
        </Tab>
      </TabList>
      <div className={styles.panels}>
        {selectedTabValue === "evaluation" && <Evaluation />}
        {selectedTabValue === "utilities" && <Utilities />}
        {selectedTabValue === "help" && <Help />}
      </div>

    </div>
  );
};

export default App;
