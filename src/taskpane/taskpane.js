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
import { ExcelBlandAltmanChart, ExcelRegressionChart } from "/src/js/charts.js";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign event handlers and other initialization logic.
    document.getElementById("run-regression").onclick = () => tryCatch(runRegression);
    document.getElementById("select-x-range").onclick = () => tryCatch(selectXRange);
    document.getElementById("select-y-range").onclick = () => tryCatch(selectYRange);
    document.getElementById("select-output-range").onclick = () => tryCatch(selectOutputRange);
    document.getElementById("select-abs-aps-address").onclick = () => tryCatch(selectApsAbsRange);
    document.getElementById("select-rel-aps-address").onclick = () => tryCatch(selectApsRelRange);
    document.getElementById("select-bland-altman-range").onclick = () =>
      tryCatch(selectBlandAltmanRange);
    document.getElementById("select-scatter-plot-range").onclick = () =>
      tryCatch(selectScatterPlotRange);
    document.getElementById("select-chart-data-range").onclick = () =>
      tryCatch(selectChartDataRange);
    document.getElementById("load-defaults-btn").onclick = () => tryCatch(setDefaultValues);
    // Hide the message box by default
    document.getElementById("msgbox").style.display = "none";
    // Set the onclick handler for the message box OK button
    document.getElementById("msgbox").onclick = () => tryCatch(alertOK);
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
        deltaRows += 1; // Add an extra row for labels
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
        data = [
          ["", "Coefficents", "Lower Confidence Limit", "Upper Confidence Limit", "Standard Error"],
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

async function selectXRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("x-range").value = range.address;
  });
}

async function selectYRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("y-range").value = range.address;
  });
}

async function selectApsAbsRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("aps-abs").value = range.address;
  });
}

async function selectApsRelRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("aps-rel").value = range.address;
  });
}

async function selectBlandAltmanRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("bland-altman-range").value = range.address;
  });
}

async function selectScatterPlotRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("scatter-plot-range").value = range.address;
  });
}

async function selectOutputRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("output-range").value = range.address;
  });
}

async function selectChartDataRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("chart-data-range").value = range.address;
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
  document.getElementById("x-range").value = "B3:C100";
  document.getElementById("y-range").value = "D3:E100";
  document.getElementById("output-range").value = "H6";
  document.getElementById("regression-method").value = "paba";
  document.getElementById("confidence-interval-method").value = "default";
  document.getElementById("use-calculated-error-ratio").checked = false;
  document.getElementById("error-ratio").value = DEFAULT_ERROR_RATIO.toFixed(4);
  document.getElementById("aps-abs").value = "I1";
  document.getElementById("aps-rel").value = "I2";
  document.getElementById("bland-altman-range").value = "H15:P30";
  document.getElementById("scatter-plot-range").value = "H31:P46";
  document.getElementById("output-labels").checked = true;
  document.getElementById("chart-data-range").value = "R1";
}
