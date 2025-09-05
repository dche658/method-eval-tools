/*
 * See LICENSE in the project root for license information.
 */

import {
  JackknifeConfidenceInterval,
  BootstrapConfidenceInterval,
  PassingBablokRegression,
  DemingRegression,
  WeightedDemingRegression,
  DEFAULT_ERROR_RATIO,
} from "/src/js/regression.js";
import { OneFactorVarianceAnalysis, TwoFactorVarianceAnalysis } from "/src/js/precision.js";
import { ExcelBlandAltmanChart, ExcelRegressionChart } from "/src/js/charts.js";

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

    // Load MVW default ranges
    document.getElementById("load-defaults-btn").onclick = () => tryCatch(setDefaultValues);

    // Hide the message box by default
    document.getElementById("msgbox").style.display = "none";

    // Set the onclick handler for the message box OK button
    document.getElementById("msgbox").onclick = () => tryCatch(alertOK);

    // Display if initialisation completes
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function runRegression() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const xRng = document.getElementById("x-range").value;
    const yRng = document.getElementById("y-range").value;
    const outputRange = document.getElementById("output-range").value;
    const regressionMethod = document.getElementById("regression-method").value;
    const ciMethod = document.getElementById("confidence-interval-method").value;
    const useCalculatedErrorRatio = document.getElementById("use-calculated-error-ratio").checked;
    const outputLabels = document.getElementById("output-labels").checked;
    const importAps = document.getElementById("import-aps").checked;
    const apsAbsInput = document.getElementById("aps-abs").value;
    const apsRelInput = document.getElementById("aps-rel").value;

    // The Bland-Altman and scatter
    const blandAltmanRange = document.getElementById("bland-altman-range").value;
    const scatterPlotRange = document.getElementById("scatter-plot-range").value;

    // If useCalculatedErrorRatio is true, we will not use the error ratio from the input field
    // but calculate it from the data.
    // If it is false, we will use the value from the input field.
    // If the input field is empty, we will use the default value.
    const errorRatioGiven = document.getElementById("error-ratio").value;

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
        document.getElementById("error-ratio").value = errorRatio.toFixed(4);
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
      let data = [];
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
    } else {
      throw new RangeError("Insufficient data");
    } // if size > 0
    await context.sync();
  });
}

function doRegression(
  x,
  y,
  regressionMethod,
  ciMethod,
  useCalculatedErrorRatio = false,
  errorRatio = DEFAULT_ERROR_RATIO
) {
  //console.log(`Error ratio is ${errorRatio}`);
  let res = {};
  if (regressionMethod === "paba") {
    let regression = new PassingBablokRegression();
    if (ciMethod === "bootstrap") {
      let regCI = new BootstrapConfidenceInterval(x, y, regression, 1000);
      res = regCI.calculate();
    } else {
      res = regression.calculate(x, y);
    }
    res["slopeSE"] = "";
    res["interceptSE"] = "";
  } else if (regressionMethod === "deming") {
    let regression = new DemingRegression(errorRatio);
    if (ciMethod === "bootstrap") {
      let regCI = new BootstrapConfidenceInterval(x, y, regression, 1000);
      res = regCI.calculate();
      res["slopeSE"] = "";
      res["interceptSE"] = "";
    } else if (ciMethod === "default") {
      let regCI = new JackknifeConfidenceInterval(x, y, regression);
      res = regCI.calculate();
    }
  } else if (regressionMethod === "wdeming") {
    let regression = new WeightedDemingRegression(errorRatio);
    if (ciMethod === "bootstrap") {
      let regCI = new BootstrapConfidenceInterval(x, y, regression, 1000);
      res = regCI.calculate();
      res["slopeSE"] = "";
      res["interceptSE"] = "";
    } else if (ciMethod === "default") {
      let regCI = new JackknifeConfidenceInterval(x, y, regression);
      res = regCI.calculate();
    }
  }
  return res;
}

// Select the range containing the X data
async function selectXRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("x-range").value = range.address;
  });
}

// Select the range containing the Y data
async function selectYRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("y-range").value = range.address;
  });
}

// Select the range containing the APS absolute value
async function selectApsAbsRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("aps-abs").value = range.address;
  });
}

// Select the range for APS relative
async function selectApsRelRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("aps-rel").value = range.address;
  });
}

// Select the output range for the Bland-Altman chart
async function selectBlandAltmanRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("bland-altman-range").value = range.address;
  });
}

// Select the output range for the Bland-Altman chart
async function selectScatterPlotRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("scatter-plot-range").value = range.address;
  });
}

// Select the output range for the regression results
async function selectOutputRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("output-range").value = range.address;
  });
}

// Select the range containing the data for the charts
async function selectChartDataRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("chart-data-range").value = range.address;
  });
}

// Precision analysis related functions

async function runANOVA() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const factorARng = document.getElementById("p-days-range").value;
    const factorBRng = document.getElementById("p-runs-range").value;
    const resultsRng = document.getElementById("p-results-range").value;
    const outputRange = document.getElementById("p-output-range").value;
    const analysisType = "one-factor";

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
    let precisionData = {
      numLevels: 0,
      data: new Array(rRange.columnCount),
    };
    // Process separately for each column in the results range.
    // Need to process first to see if columns contain data before undertaking
    // variance component analysis.
    for (let i = 0; i < rRange.columnCount; i++) {
      let aLevels = [];
      let bLevels = [];
      let results = [];
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
      if (results.length !== aLevels.length)
        throw new Error(
          `Number of length of results ${results.length} does not match length of factor A ${aLevels.length}.`
        );

      if (results.length > 1) {
        precisionData.numLevels += 1;
      }

      // if no factor B or only one level then perform one factor analysis
      // Default is one factor analysis
      const dataObject = {
        type: "one-factor",
        aLevels: aLevels,
        results: results,
        aDict: aDict,
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

function formatPrecisionForExcel(labels, columnData) {
  let rangeValues = [];
  for (let i = 0; i < labels.length; i++) {
    let numCols = columnData.length + 1; //add column for label
    let row = new Array(numCols);
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

function analyseOneFactor(precisionData) {
  let columnData = new Array(precisionData.data.length);
  let res;
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

function analyseTwoFactor(precisionData) {
  let columnData = new Array(precisionData.data.length);
  let res;
  let maxSize = 0;
  for (let i = 0; i < precisionData.data.length; i++) {
    let dataObj = precisionData.data[i];
    if (dataObj.results.length > 0) {
      const twoFactorAnalysis = new TwoFactorVarianceAnalysis(
        dataObj.aLevels,
        dataObj.bLevels,
        dataObj.results,
        precisionData.numLevels,
        0.05
      );
      res = twoFactorAnalysis.calculate();
      let size = res.length;
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
    "SSAB",
    "DF T",
    "DF A",
    "DF B",
    "DF AB",
    "DF E",
    "MSA",
    "MSB",
    "MSAB",
    "MSE",
    "F AB",
    "F A",
    "F B",
    "N",
    "Num A",
    "Num B",
    "Num E",
    "Var A",
    "Var AB",
    "Var E",
    "SD A",
    "SD AB",
    "SD E",
    "SD WL",
    "CV A",
    "CV AB",
    "CV E",
    "CV T",
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
}

// The following functions are for selecting ranges for the precision analysis
// Select the range containing the factor A (e.g. days)
async function selectFactorARange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("p-days-range").value = range.address;
  });
}

// Select the range containing the factor B (e.g. runs)
async function selectFactorBRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("p-runs-range").value = range.address;
  });
}

// Select the range containing the values (e.g. replicates)
// May have multiple columns representing different levels
async function selectPrecisionResultsRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("p-results-range").value = range.address;
  });
}

// Select the output range for the precision results
async function selectPrecisionOutputRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("p-output-range").value = range.address;
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
  const diffType = document.getElementById("difference-type").value;

  // Load the range for plotting the chart
  const outputRange = document.getElementById("bland-altman-range").value;

  const chartDataRange = document.getElementById("chart-data-range").value;
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
  const outputRange = document.getElementById("scatter-plot-range").value;

  // Range for chart data
  const chartDataRange = document.getElementById("chart-data-range").value;
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

function processRangeData(range) {
  let means = [];
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
  // If the range has two columns then we calculate assuming analysis was done in replicates
  if (range.columnCount === 2) {
    let x1 = [];
    let x2 = [];
    let devsq = [];
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
    let sd = Math.sqrt(devsq.reduce((a, b) => a + b, 0) / (size - 1));
    let mean = means.reduce((a, b) => a + b, 0) / size;
    let cv = sd / mean;
    // Return the mean and coefficient of variation
    return { means: means, x1: x1, x2: x2, devsq: devsq, sd: sd, cv: cv, size: size, mean: mean };
  } else if (range.columnCount === 1) {
    // Process single column
    for (let i = 0; i < range.rowCount; i++) {
      if (typeof range.values[i][0] === "number") {
        means.push(range.values[i][0]);
        size++;
      }
    }
    let mean = means.reduce((a, b) => a + b, 0) / size;
    return { means: means, x1: [], x2: [], devsq: [], sd: 0, cv: 0, size: size, mean: mean };
  }
}

/* Configure default values for the input fields */
function setDefaultValues() {
  // Set explicitly for now but in the future we may load these from a workbook or a settings file.
  // APS
  document.getElementById("aps-abs").value = "J4";
  document.getElementById("aps-rel").value = "J6";

  // Method Comparison
  document.getElementById("x-range").value = "O61:P160";
  document.getElementById("y-range").value = "R61:S160";
  document.getElementById("output-range").value = "C167";
  document.getElementById("regression-method").value = "paba";
  document.getElementById("confidence-interval-method").value = "default";
  document.getElementById("use-calculated-error-ratio").checked = false;
  document.getElementById("error-ratio").value = DEFAULT_ERROR_RATIO.toFixed(4);
  document.getElementById("bland-altman-range").value = "N162:U175";
  document.getElementById("scatter-plot-range").value = "N177:U190";
  document.getElementById("output-labels").checked = true;
  document.getElementById("chart-data-range").value = "AP61";

  // Imprecision
  document.getElementById("p-days-range").value = "B295:B319";
  document.getElementById("p-runs-range").value = "C295:C319";
  document.getElementById("p-results-range").value = "D295:G319";
  document.getElementById("p-output-range").value = "AD294";
}
