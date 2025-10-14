/*
 * See LICENSE in the project root for license information.
 */

import {
  provideFluentDesignSystem,
  fluentTextField,
  fluentNumberField,
  fluentSelect,
  fluentOption,
  fluentButton,
  fluentTextArea,
  fluentAccordion,
  fluentAccordionItem,
  fluentCheckbox,
} from "@fluentui/web-components";

provideFluentDesignSystem()
  .register(
    fluentTextField(),
    fluentNumberField(),
    fluentSelect(),
    fluentOption(),
    fluentButton(),
    fluentTextArea(),
    fluentAccordion(),
    fluentAccordionItem(),
    fluentCheckbox()
  );

import {
  JackknifeConfidenceInterval,
  BootstrapConfidenceInterval,
  PassingBablokRegression,
  DemingRegression,
  WeightedDemingRegression,
  DEFAULT_ERROR_RATIO,
} from "../regression";
import {
  OneFactorVarianceAnalysis, TwoFactorVarianceAnalysis,
  OneFactorVariance, TwoFactorVariance,
} from "../precision";
import {
  ExcelBlandAltmanChart, ExcelRegressionChart
} from "../charts";
import {
  ContingencyTableBuilder, QualitativeContengencyTableBuilder,
  ConcordanceCalculator
} from "../concordance";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign event handlers and other initialization logic.

    // Method comparison and regression
    document.getElementById("run-regression").onclick = () => tryCatch(runRegression);
    document.getElementById("select-x-range").onclick = () => tryCatch(selectXRange);
    document.getElementById("select-y-range").onclick = () => tryCatch(selectYRange);
    document.getElementById("select-output-range").onclick = () => tryCatch(selectOutputRange);

    document.getElementById("select-bland-altman-range").onclick = () =>
      tryCatch(selectBlandAltmanRange);
    document.getElementById("select-scatter-plot-range").onclick = () =>
      tryCatch(selectScatterPlotRange);
    document.getElementById("select-chart-data-range").onclick = () =>
      tryCatch(selectChartDataRange);
    document.getElementById("select-concordance-output-range").onclick = () =>
      tryCatch(selectConcordanceOutputRange);


    // APS
    document.getElementById("select-abs-aps-address").onclick = () => tryCatch(selectApsAbsRange);
    document.getElementById("select-rel-aps-address").onclick = () => tryCatch(selectApsRelRange);

    // Precision and variance components
    document.getElementById("select-p-days-range").onclick = () => tryCatch(selectFactorARange);
    document.getElementById("select-p-runs-range").onclick = () => tryCatch(selectFactorBRange);
    document.getElementById("select-p-results-range").onclick = () =>
      tryCatch(selectPrecisionResultsRange);
    document.getElementById("select-p-output-range").onclick = () =>
      tryCatch(selectPrecisionOutputRange);
    document.getElementById("run-precision").onclick = () => tryCatch(runANOVA);

    // Qualitative data comparison
    document.getElementById("select-qx-range").onclick = () => tryCatch(selectQualXRange);
    document.getElementById("select-qy-range").onclick = () => tryCatch(selectQualYRange);
    document.getElementById("select-q-output-range").onclick = () => tryCatch(selectQualOutputRange);
    document.getElementById("qual-comparison").onclick = () => tryCatch(qualComparison);


    // Load MVW default ranges
    document.getElementById("load-defaults-btn").onclick = () => tryCatch(setDefaultValues);
    document.getElementById("load-wb-def-btn").onclick = () => tryCatch(loadWorkbookDefaults);

    // Hide the message box by default
    document.getElementById("msgbox").style.display = "none";

    // Set the onclick handler for the message box OK button
    document.getElementById("msgbox").onclick = () => tryCatch(alertOK);

    // Setup a precision experiment layout
    document.getElementById("setup-precision").onclick = () => tryCatch(setupPrecision);
    document.getElementById("select-c-layout-range").onclick = () =>
      tryCatch(selectPrecisionLayoutRange);

    // Display if initialisation completes
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Add value change listener to regression method select
    document.getElementById("regression-method").onchange = regressionMethodChanged;
  }
});

async function runRegression() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const xRngInput = document.getElementById("x-range") as HTMLInputElement;
    const yRngInput = document.getElementById("y-range") as HTMLInputElement;
    const outputRangeInput = document.getElementById("output-range") as HTMLInputElement;
    const regressionMethodInput = document.getElementById("regression-method") as HTMLSelectElement;
    const ciMethodInput = document.getElementById("confidence-interval-method") as HTMLSelectElement;
    const useCalculatedErrorRatioInput = document.getElementById(
      "use-calculated-error-ratio"
    ) as HTMLInputElement;
    const outputLabelsInput = document.getElementById("output-labels") as HTMLInputElement;
    const importApsInput = document.getElementById("import-aps") as HTMLInputElement;
    const apsAbsHtmlInput = document.getElementById("aps-abs") as HTMLInputElement;
    const apsRelHtmlInput = document.getElementById("aps-rel") as HTMLInputElement;
    const xRng: string = xRngInput.value;
    const yRng: string = yRngInput.value;
    const outputRange: string = outputRangeInput.value;
    const regressionMethod: string = regressionMethodInput.value;
    const ciMethod: string = ciMethodInput.value;
    const useCalculatedErrorRatio: boolean = Boolean(useCalculatedErrorRatioInput.checked);
    const outputLabels: boolean = Boolean(outputLabelsInput.checked);
    const importAps: boolean = Boolean(importApsInput.checked);
    const apsAbsInput: string = apsAbsHtmlInput.value;
    const apsRelInput: string = apsRelHtmlInput.value;

    // The Bland-Altman and scatter
    const blandAltmanRangeInput = document.getElementById("bland-altman-range") as HTMLInputElement;
    const scatterPlotRangeInput = document.getElementById("scatter-plot-range") as HTMLInputElement;
    const blandAltmanRange: string = blandAltmanRangeInput.value;
    const scatterPlotRange: string = scatterPlotRangeInput.value;

    // If useCalculatedErrorRatio is true, we will not use the error ratio from the input field
    // but calculate it from the data.
    // If it is false, we will use the value from the input field.
    // If the input field is empty, we will use the default value.
    const errorRatioInput = document.getElementById("error-ratio") as HTMLInputElement;
    const errorRatioGiven = errorRatioInput.value;

    let apsAbs = -1;
    let apsRel = -1;
    // Load APS data if importAps is checked
    if (importAps) {
      if (apsAbsInput === "" || apsRelInput === "") {
        throw new Error("Please select the APS absolute and relative ranges.");
      }
      // Load the APS data from the ranges
      const apsAbsRange = currentWorksheet.getRange(apsAbsInput);
      const apsRelRange = currentWorksheet.getRange(apsRelInput);
      apsAbsRange.load(["values"]);
      apsRelRange.load(["values"]);
      await context.sync();
      apsAbs = apsAbsRange.values[0][0];
      apsRel = apsRelRange.values[0][0];
      if (typeof apsAbs !== "number" || typeof apsRel !== "number") {
        throw new TypeError(
          `APS absolute and relative values at ${apsAbsInput} and ${apsRelInput} must be numbers.`
        );
      }
    } else {
      // if importAps is not checked, attempt to get the APS values from the input fields
      // If the input fields are empty, we will use -1 as the default value
      if (apsAbsInput !== "") {
        apsAbs = Number(apsAbsInput);
        if (isNaN(apsAbs)) {
          throw new TypeError("APS absolute value must be a number.");
        }
      } else {
        apsAbs = -1; // Default value if not provided
      }
      if (apsRelInput !== "") {
        apsRel = Number(apsRelInput);
        if (isNaN(apsRel)) {
          throw new TypeError("APS relative value must be a number.");
        }
      } else {
        apsRel = -1; // Default value if not provided
      }
    }

    // Load the ranges
    const xRange = currentWorksheet.getRange(xRng);
    const yRange = currentWorksheet.getRange(yRng);
    xRange.load(["rowCount", "columnCount", "values"]);
    yRange.load(["rowCount", "columnCount", "values"]);
    await context.sync();

    //let size = 0;

    let xData = processRangeData(xRange);
    let yData = processRangeData(yRange);
    let xArr = xData.means;
    let yArr = yData.means;
    if (xData.size !== yData.size) {
      throw new RangeError("X and Y ranges must have the same number of rows");
    }
    // for (let i = 0; i < xRange.rowCount; i++) {
    //   if (typeof xRange.values[i][0] === "number" && typeof yRange.values[i][0] === "number") {
    //     xArr.push(xRange.values[i][0]);
    //     yArr.push(yRange.values[i][0]);
    //     size++;
    //   }
    // }
    if (xData.size > 0) {
      let errorRatio = DEFAULT_ERROR_RATIO;
      if (!useCalculatedErrorRatio) {
        // If we are not using the calculated error ratio, we will use the value from the input field
        // If the input field is empty, we will use the default value.
        if (errorRatioGiven === "") {
          errorRatio = DEFAULT_ERROR_RATIO;
        } else {
          errorRatio = Number(errorRatioGiven);
        }
      } else if (useCalculatedErrorRatio && xData.devsq.length > 0) {
        // If we are using the calculated error ratio, we will calculate it from the data
        // The error ratio is the coefficient of variation of the means
        // We will use the sd_y / sd_x as the error ratio for deming regression
        // and cv_y / cv_x for weighted deming regression.
        if (regressionMethod === "deming" || regressionMethod === "paba") {
          // For Deming regression, we will use the standard deviation of the y data divided by the standard deviation of the x data
          let sd_x = xData.sd;
          let sd_y = yData.sd;
          if (sd_x === 0 || sd_y === 0) {
            throw new RangeError("Standard deviation cannot be zero");
          }
          errorRatio = sd_y / sd_x;
        } else if (regressionMethod === "wdeming") {
          // For Weighted Deming regression, we will use the coefficient of variation of the y data divided by the coefficient of variation of the x data
          let cv_x = xData.cv;
          let cv_y = yData.cv;
          if (cv_x === 0 || cv_y === 0) {
            throw new RangeError("Coefficient of variation cannot be zero");
          }
          errorRatio = cv_y / cv_x;
        }
        // Set the error ratio input field to the calculated value
        errorRatioInput.value = errorRatio.toFixed(4);
      }

      //Ensure output has two rows and four columns
      let outputRng = currentWorksheet.getRange(outputRange);
      outputRng.load([`rowCount`, `columnCount`]);
      await context.sync();
      let deltaRows = 2 - outputRng.rowCount;
      let deltaCols = 4 - outputRng.columnCount;
      if (outputLabels) {
        deltaRows += 2; // Add an extra row for labels
        deltaCols += 1; // Add an extra column for labels
      }
      outputRng = outputRng.getResizedRange(deltaRows, deltaCols);

      // Run the regression
      let res = doRegression(
        xArr,
        yArr,
        regressionMethod,
        ciMethod,
        useCalculatedErrorRatio,
        errorRatio
      );
      let data: any[][] = [];
      if (outputLabels) {
        let method = "Passing-Bablock Regression";
        let ciType = "Non-parametric CI";
        let seLabel = "SE";
        if (regressionMethod === "deming") {
          method = "Deming Regression";
          ciType = ciMethod === "default" ? "Jackknife CI" : "Bootstrap CI";
        } else if (regressionMethod === "wdeming") {
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
      // Create Bland-Altman chart if requested
      if (blandAltmanRange !== "") {
        createBlandAltmanChart(xData, yData, apsAbs, apsRel);
      }
      // Create regression chart if requested
      if (scatterPlotRange !== "") {
        createRegressionChart(xData, yData, res, apsAbs, apsRel);
      }

      // Process concordance assessment if thresholds have been specified
      const concordanceThresholds: number[][] = loadThresholds();
      if (concordanceThresholds.length > 0) {
        const contingencyBuilder = new ContingencyTableBuilder(concordanceThresholds);
        const contingencyTable = contingencyBuilder.build(xData.means, yData.means);
        const concordanceCalculator = new ConcordanceCalculator(contingencyTable);
        const concordanceResults = concordanceCalculator.calculate();
        //to-do: copy to Excel
        const concordanceCell = document.getElementById("concordance-output") as HTMLInputElement;
        if (concordanceCell.value === "") {
          throw Error("Please select the concordance output range.");
        } else {
          const concordanceRange = currentWorksheet.getRange(concordanceCell.value);
          const contingencyRange = concordanceRange.getAbsoluteResizedRange(contingencyTable.table.length, contingencyTable.table[0].length);
          contingencyRange.values = contingencyTable.table;
          const offset = Math.max(8, contingencyTable.table.length + 1)
          const cohenRange = concordanceRange.getOffsetRange(offset, 0).getAbsoluteResizedRange(6, 2);
          cohenRange.values = concordanceCalculator.formatResultsAsArray(concordanceResults);
          await context.sync();
        }
      }
    } else {
      throw new RangeError("Insufficient data");
    } // if size > 0
    await context.sync();
  });
}

function doRegression(
  x: number[],
  y: number[],
  regressionMethod: string,
  ciMethod: string,
  useCalculatedErrorRatio: boolean = false,
  errorRatio: number = DEFAULT_ERROR_RATIO
): {
  slope: number,
  intercept: number,
  slopeSE: number,
  interceptSE: number,
  slopeLCL: number,
  slopeUCL: number,
  interceptLCL: number,
  interceptUCL: number
} {
  //console.log(`Error ratio is ${errorRatio}`);
  let res = {
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

// Select the range containing the X data
async function selectXRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const xRangeInput = document.getElementById("x-range") as HTMLInputElement;
    xRangeInput.value = range.address;
  });
}

// Select the range containing the Y data
async function selectYRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const yRangeInput = document.getElementById("y-range") as HTMLInputElement;
    yRangeInput.value = range.address;
  });
}

// Select the range containing the APS absolute value
async function selectApsAbsRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const apsAbsInput = document.getElementById("aps-abs") as HTMLInputElement;
    apsAbsInput.value = range.address;
  });
}

// Select the range for APS relative
async function selectApsRelRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const apsRelInput = document.getElementById("aps-rel") as HTMLInputElement;
    apsRelInput.value = range.address;
  });
}

// Select the output range for the Bland-Altman chart
async function selectBlandAltmanRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const blandAltmanInput = document.getElementById("bland-altman-range") as HTMLInputElement;
    blandAltmanInput.value = range.address;
  });
}

// Select the output range for the Bland-Altman chart
async function selectScatterPlotRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const scatterPlotInput = document.getElementById("scatter-plot-range") as HTMLInputElement;
    scatterPlotInput.value = range.address;
  });
}

// Select the output range for the regression results
async function selectOutputRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const outputRangeInput = document.getElementById("output-range") as HTMLInputElement;
    outputRangeInput.value = range.address;
  });
}

// Select the range containing the data for the charts
async function selectChartDataRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const chartDataRangeInput = document.getElementById("chart-data-range") as HTMLInputElement;
    chartDataRangeInput.value = range.address;
  });
}

// Select the range containing the qualitative X data
async function selectQualXRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const r = document.getElementById("qx-range") as HTMLInputElement;
    r.value = range.address;
  });
}

// Select the range containing the qualitative X data
async function selectQualYRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const r = document.getElementById("qy-range") as HTMLInputElement;
    r.value = range.address;
  });
}

// Select the output range for the qualitative comparison results
async function selectQualOutputRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const r = document.getElementById("q-output-range") as HTMLInputElement;
    r.value = range.address;
  });
}

// Run the qualitative data comparison
async function qualComparison() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const xRngInput = document.getElementById("qx-range") as HTMLInputElement;
    const yRngInput = document.getElementById("qy-range") as HTMLInputElement;
    const outputRangeInput = document.getElementById("q-output-range") as HTMLInputElement;
    const xRng: string = xRngInput.value;
    const yRng: string = yRngInput.value;
    const outputRange: string = outputRangeInput.value;

    // Load the ranges
    const xRange = currentWorksheet.getRange(xRng);
    const yRange = currentWorksheet.getRange(yRng);
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
    const concordanceCell = document.getElementById("q-output-range") as HTMLInputElement;
    if (concordanceCell.value === "") {
      throw Error("Please select the concordance output range.");
    } else {
      const concordanceRange = currentWorksheet.getRange(concordanceCell.value);
      const contingencyRange = concordanceRange.getAbsoluteResizedRange(contingencyTable.table.length, contingencyTable.table[0].length);
      contingencyRange.values = contingencyTable.table;
      const offset = Math.max(8, contingencyTable.table.length + 1)
      const cohenRange = concordanceRange.getOffsetRange(offset, 0).getAbsoluteResizedRange(6, 2);
      cohenRange.values = concordanceCalculator.formatResultsAsArray(concordanceResults);
      await context.sync();
    }
  });
}

// Precision analysis related functions
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

async function runANOVA() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const factorARngInput = document.getElementById("p-days-range") as HTMLInputElement;
    const factorARng = factorARngInput.value;
    const factorBRngInput = document.getElementById("p-runs-range") as HTMLInputElement;
    const factorBRng = factorBRngInput.value;
    const resultsRngInput = document.getElementById("p-results-range") as HTMLInputElement;
    const resultsRng = resultsRngInput.value;
    const outputRangeInput = document.getElementById("p-output-range") as HTMLInputElement;
    const outputRange = outputRangeInput.value;
    let analysisType = "one-factor";

    // Load the ranges
    const aRange = currentWorksheet.getRange(factorARng);
    const bRange = currentWorksheet.getRange(factorBRng);
    const rRange = currentWorksheet.getRange(resultsRng);
    let oRange = currentWorksheet.getRange(outputRange);
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
          if (factorBRng !== "") {
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
          `Number of length of results ${results.length} does not match length of factor A ${aLevels.length}.`
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
    let rangeValues;
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
  });
} // runPrecisionAnalysis

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

function analyseOneFactor(precisionData: PrecisionData) {
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

function analyseTwoFactor(precisionData: PrecisionData) {
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

// Setup the layout for a precision experiment
async function setupPrecision() {
  await Excel.run(async (context) => {
    const daysInput = document.getElementById("c-num-days") as HTMLInputElement;
    const runsInput = document.getElementById("c-num-runs") as HTMLInputElement;
    const repsInput = document.getElementById("c-num-reps") as HTMLInputElement;
    const levelsInput = document.getElementById("c-num-levels") as HTMLInputElement;
    const layoutRngInput = document.getElementById("c-layout-range") as HTMLInputElement;

    const daysVal = daysInput.value;
    const runsVal = runsInput.value;
    const repsVal = repsInput.value;
    const levelsVal = levelsInput.value;
    const layoutRng = layoutRngInput.value;
    if (daysVal === "") {
      throw new Error("Please enter the number of days.");
    }
    if (runsVal === "") {
      throw new Error("Please enter the number of runs.");
    }
    if (repsVal === "") {
      throw new Error("Please enter the number of replicates.");
    }
    if (levelsVal === "") {
      throw new Error("Please enter the number of levels.");
    }
    if (layoutRng === "") {
      throw new Error("Please select the top left cell of the layout range.");
    }
    const days = Number(daysVal);
    const runs = Number(runsVal);
    const reps = Number(repsVal);
    const levels = Number(levelsVal);

    if (isNaN(days) || isNaN(runs) || isNaN(reps) || isNaN(levels)) {
      throw new Error(
        "Please enter valid numbers for the number of days, runs, replicates and levels."
      );
    }

    const layout = setupLayout(days, runs, reps, levels);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(layoutRng);
    range.load(["rowCount", "columnCount", "values"]);
    await context.sync();

    // write the layout to Excel
    const layoutRange = range.getAbsoluteResizedRange(layout.length, layout[0].length);
    layoutRange.values = layout;
    await context.sync();

    // copy the range addresses to the precision analysis fields
    const daysRange = range.getCell(1, 0).getAbsoluteResizedRange(layout.length - 1, 1);
    const runsRange = range.getCell(1, 1).getAbsoluteResizedRange(layout.length - 1, 1);
    const valuesRange = range.getCell(1, 2).getAbsoluteResizedRange(layout.length - 1, levels);
    daysRange.load(["address"]);
    runsRange.load(["address"]);
    valuesRange.load(["address"]);
    await context.sync();
    const daysRangeInput = document.getElementById("p-days-range") as HTMLInputElement;
    const runsRangeInput = document.getElementById("p-runs-range") as HTMLInputElement;
    const resultsRngInput = document.getElementById("p-results-range") as HTMLInputElement;
    daysRangeInput.value = daysRange.address;
    runsRangeInput.value = runsRange.address;
    resultsRngInput.value = valuesRange.address;
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

// Setup the layout for a precision experiment
async function selectPrecisionLayoutRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const layoutRngInput = document.getElementById("c-layout-range") as HTMLInputElement;
    layoutRngInput.value = range.address;
  });
}

// The following functions are for selecting ranges for the precision analysis
// Select the range containing the factor A (e.g. days)
async function selectFactorARange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const factorARngInput = document.getElementById("p-days-range") as HTMLInputElement;
    factorARngInput.value = range.address;
  });
}

// Select the range containing the factor B (e.g. runs)
async function selectFactorBRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const factorARngInput = document.getElementById("p-runs-range") as HTMLInputElement;
    factorARngInput.value = range.address;
  });
}

// Select the range containing the values (e.g. replicates)
// May have multiple columns representing different levels
async function selectPrecisionResultsRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const resultsRngInput = document.getElementById("p-results-range") as HTMLInputElement;
    resultsRngInput.value = range.address;
  });
}

// Select the output range for the precision results
async function selectPrecisionOutputRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const pOutputRangeInput = document.getElementById("p-output-range") as HTMLInputElement;
    pOutputRangeInput.value = range.address;
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Log the error to the console
    console.error(error);
    // Display the error in a message box
    const msgbox = document.getElementById("msgbox");
    const msgboxContent = document.getElementById("msgbox-content");
    msgbox.style.display = "block";
    msgboxContent.innerHTML = error.message;
  }
}

/** Create Bland Altman Chart */
function createBlandAltmanChart(xData, yData, apsAbs, apsRel) {
  // Type of difference plot
  const diffTypeInput = document.getElementById("difference-type") as HTMLSelectElement
  const diffType = diffTypeInput.value;

  // Load the range for plotting the chart
  const blandAltmanInput = document.getElementById("bland-altman-range") as HTMLInputElement;
  const outputRange = blandAltmanInput.value;

  // Range for chart data
  const chartDataRangeInput = document.getElementById("chart-data-range") as HTMLInputElement;
  const chartDataRange = chartDataRangeInput.value;
  if (chartDataRange === "") {
    throw new Error("Please select the chart data range.");
  }

  const blandAltmanChart = new ExcelBlandAltmanChart(
    xData.means,
    yData.means,
    diffType,
    apsAbs,
    apsRel,
    chartDataRange,
    outputRange
  );
  blandAltmanChart.createChart();
} // createBlandAltmanChart

function createRegressionChart(xData, yData, regressionResults, apsAbs, apsRel) {
  // Load the range for plotting the chart
  const scatterPlotInput = document.getElementById("scatter-plot-range") as HTMLInputElement;
  const outputRange = scatterPlotInput.value;

  // Range for chart data
  const chartDataRangeInput = document.getElementById("chart-data-range") as HTMLInputElement;
  const chartDataRange = chartDataRangeInput.value;
  if (chartDataRange === "") {
    throw new Error("Please select the chart data range.");
  }
  const regressionChart = new ExcelRegressionChart(
    xData.means,
    yData.means,
    regressionResults.slope,
    regressionResults.intercept,
    apsAbs,
    apsRel,
    chartDataRange,
    outputRange
  );
  regressionChart.createChart();
}

function alertOK() {
  const msgbox = document.getElementById("msgbox");
  msgbox.style.display = "none";
}

function processRangeData(range: Excel.Range): {
  means: number[];
  x1: number[];
  x2: number[];
  devsq: number[];
  sd: number;
  cv: number;
  size: number;
  mean: number;
} {
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

function loadThresholds(): number[][] {
  const X_COL = 0;
  const Y_COL = 1;
  let thresholds: number[][] = []
  for (let i = 0; i < 4; i++) {
    let x_id = `x-threshold-${i}`;
    let y_id = `y-threshold-${i}`;
    let x_input = document.getElementById(x_id) as HTMLInputElement;
    let y_input = document.getElementById(y_id) as HTMLInputElement;
    if (x_input.value !== "" && y_input.value !== "") {
      thresholds.push([Number(x_input.value), Number(y_input.value)]);
    } else if (x_input.value !== "" || y_input.value !== "") {
      throw Error("Concordance thresholds for x and y methods must be equal in number and aligned.");
    }
  }
  return thresholds;
}

async function selectConcordanceOutputRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    const outputRangeInput = document.getElementById("concordance-output") as HTMLInputElement;
    outputRangeInput.value = range.address;
  });
}

// Read the default cell addresses from and excel workbook.
async function loadWorkbookDefaults() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const address = "A2:B23";
    const range = worksheet.getRange(address);
    range.load(["values", "rowCount", "columnCount"]);

    await context.sync();
    for (let row = 0; row < range.rowCount; row++) {
      const attribute = range.values[row][0];
      const value = range.values[row][1];
      let element: HTMLInputElement | HTMLSelectElement;
      if (attribute === "import-aps" ||
        attribute === "use-calculated-error-ratio" ||
        attribute === "output-labels") {
        element = document.getElementById(attribute) as HTMLInputElement;
        element.checked = value;
      } else if (attribute === "regression-method" ||
        attribute === "confidence-interval-method") {
        element = document.getElementById(attribute) as HTMLSelectElement;
        element.value = value;
      } else {
        element = document.getElementById(attribute) as HTMLInputElement;
        element.value = value;
      }
    }
  });
} //readLayout

/* Configure default values for the input fields */
function setDefaultValues() {
  // Set explicitly for now but in the future we may load these from a workbook or a settings file.
  // APS
  const importApsInput = document.getElementById("import-aps") as HTMLInputElement;
  importApsInput.checked = true;
  const apsAbsInput = document.getElementById("aps-abs") as HTMLInputElement;
  apsAbsInput.value = "J4";
  const apsRelInput = document.getElementById("aps-rel") as HTMLInputElement;
  apsRelInput.value = "J6";

  // Method Comparison
  const xRangeInput = document.getElementById("x-range") as HTMLInputElement;
  xRangeInput.value = "O61:P160";
  const yRangeInput = document.getElementById("y-range") as HTMLInputElement;
  yRangeInput.value = "R61:S160";
  const outputRangeInput = document.getElementById("output-range") as HTMLInputElement;
  outputRangeInput.value = "C167";
  const regressionMethodInput = document.getElementById(
    "regression-method"
  ) as HTMLSelectElement;
  regressionMethodInput.value = "paba";
  const confIntMethodInput = document.getElementById(
    "confidence-interval-method"
  ) as HTMLSelectElement;
  confIntMethodInput.value = "default";
  const useCalculatedErrorRatioInput = document.getElementById("use-calculated-error-ratio") as HTMLInputElement;
  useCalculatedErrorRatioInput.checked = false;
  const errorRatioInput = document.getElementById("error-ratio") as HTMLInputElement;
  errorRatioInput.value = DEFAULT_ERROR_RATIO.toFixed(4);
  const blandAltmanInput = document.getElementById("bland-altman-range") as HTMLInputElement;
  blandAltmanInput.value = "N162:U175";
  const scatterPlotInput = document.getElementById("scatter-plot-range") as HTMLInputElement;
  scatterPlotInput.value = "N177:U190";
  const outputLabelsInput = document.getElementById("output-labels") as HTMLInputElement;
  outputLabelsInput.checked = true;
  const chartDataRangeInput = document.getElementById("chart-data-range") as HTMLInputElement;
  chartDataRangeInput.value = "AP61";
  const concordanceOutputCell = document.getElementById("concordance-output") as HTMLInputElement;
  concordanceOutputCell.value = "C176";

  // Imprecision
  const factorARngInput = document.getElementById("p-days-range") as HTMLInputElement;
  factorARngInput.value = "B300:B324";
  const factorBRngInput = document.getElementById("p-runs-range") as HTMLInputElement;
  factorBRngInput.value = "C300:C324";
  const resultsRngInput = document.getElementById("p-results-range") as HTMLInputElement;
  resultsRngInput.value = "D300:G324";
  const pOutputRangeInput = document.getElementById("p-output-range") as HTMLInputElement;
  pOutputRangeInput.value = "W299";

  // concordance of qualitative data
  const qxRngInput = document.getElementById("qx-range") as HTMLInputElement;
  const qyRngInput = document.getElementById("qy-range") as HTMLInputElement;
  const qOutInput = document.getElementById("q-output-range") as HTMLInputElement;
  qxRngInput.value = "E61:E160";
  qyRngInput.value = "H61:H160";
  qOutInput.value = "C176";
}

function regressionMethodChanged() {
  const regressionMethodInput = document.getElementById(
    "regression-method"
  ) as HTMLSelectElement;

  if (regressionMethodInput.value === "deming" || regressionMethodInput.value === "wdeming") {
    document.getElementById("error-ratio-container").style.display = "block";
  } else {
    document.getElementById("error-ratio-container").style.display = "none";
  }
}