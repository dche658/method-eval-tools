/**
 * ExcelBlandAltmanChart and ExcelRegressionChart classes for creating
 * charts in Excel using Office.js
 * Requires jStat library for statistical functions
 *
 * @author Douglas Chesher
 *
 */

import { mean, stdev } from "jstat-esm";

class ExcelBlandAltmanChart {
  private x: number[]; // Array of numbers
  private y: number[]; // Array of numbers
  private diffType: string; // 'abs' or 'rel'
  private apsAbs: number; // Absolute total error specification
  private apsRel: number; // Relative total error specification
  private chartDataRange: string; // Cell range for chart data
  private outputRange: string; // Cell range for output

  constructor(
    x: number[],
    y: number[],
    diffType: string,
    apsAbs: number,
    apsRel: number,
    chartDataRange: string,
    outputRange: string = ""
  ) {
    //console.log("ExcelBlandAltmanChart initialized");
    this.x = x;
    this.y = y;
    this.diffType = diffType;
    this.apsAbs = apsAbs;
    this.apsRel = apsRel;
    this.chartDataRange = chartDataRange;
    this.outputRange = outputRange;
    if (diffType !== "rel" && diffType !== "abs") {
      throw new Error("Bland-Altman Type must be 'rel' or 'abs'.");
    }
  } // constructor

  initializeModel(): {
    means: number[];
    diffs: number[];
    meanDiff: number;
    sdDiff: number;
    minX: number;
    maxX: number;
    meanDiffData: {
      meanDiffXValues: number[][];
      meanDiffYValues: number[][];
    };
    apsData: {
      apsUpperX: number[][];
      apsUpperY: number[][];
      apsLowerX: number[][];
      apsLowerY: number[][];
    } | null;
    loaData: {
      upperLimitXValues: number[][];
      upperLimitYValues: number[][];
      lowerLimitXValues: number[][];
      lowerLimitYValues: number[][];
    };
  } {
    // The mean of x and y
    let means: number[] = [];

    // The difference or relative between y and x
    let diffs: number[] = [];
    if (this.x.length !== this.y.length || this.x.length === 0) {
      throw new Error("Input arrays must have the same non-zero length.");
    }
    for (let i = 0; i < this.x.length; i++) {
      means[i] = (this.x[i] + this.y[i]) / 2;
      diffs[i] =
        this.diffType === "rel" ? (this.y[i] - this.x[i]) / means[i] : this.y[i] - this.x[i];
    }
    const meanDiff = mean(diffs);
    const sdDiff = stdev(diffs, true); // Sample standard deviation
    const minX = Math.min(...means);
    const maxX = Math.max(...means);

    let apsData = null;
    if (!(this.apsAbs <= 0 && this.apsRel <= 0)) {
      // Calculate the APS lines only if APS values are provided
      apsData = this.getBlandAltmanApsData(minX, maxX, this.apsAbs, this.apsRel, this.diffType);
    }

    const loaData = this.getLimitOfAgreementData(minX, maxX, meanDiff, sdDiff);

    const meanDiffData = this.getMeanDifferenceData(minX, maxX, meanDiff);

    return {
      means: means,
      diffs: diffs,
      meanDiff: meanDiff,
      sdDiff: sdDiff,
      minX: minX,
      maxX: maxX,
      meanDiffData: meanDiffData,
      apsData: apsData,
      loaData: loaData,
    };
  } // initializeModel

  getMeanDifferenceData(
    minX: number,
    maxX: number,
    meanDiff: number
  ): { meanDiffXValues: number[][]; meanDiffYValues: number[][] } {
    // This function returns the data for the Mean Difference Line Series
    const meanDiffXValues = [[minX], [maxX]]; //2 rows, 1 column
    const meanDiffYValues = [[meanDiff], [meanDiff]]; //2 rows, 1 column
    return {
      meanDiffXValues: meanDiffXValues,
      meanDiffYValues: meanDiffYValues,
    };
  } // getMeanDifferenceData

  getLimitOfAgreementData(
    minX: number,
    maxX: number,
    meanDiff: number,
    sdDiff: number
  ): {
    upperLimitXValues: number[][];
    upperLimitYValues: number[][];
    lowerLimitXValues: number[][];
    lowerLimitYValues: number[][];
  } {
    // This function returns the data for the 95% Limits of Agreemen Line Series
    const upperLimitXValues = [[minX], [maxX]]; //2 rows, 1 column
    const upperLimitYValues = [[meanDiff + sdDiff * 1.96], [meanDiff + sdDiff * 1.96]]; //2 rows, 1 column
    const lowerLimitXValues = [[minX], [maxX]]; //2 rows, 1 column
    const lowerLimitYValues = [[meanDiff - sdDiff * 1.96], [meanDiff - sdDiff * 1.96]]; //2 rows, 1 column
    return {
      upperLimitXValues: upperLimitXValues,
      upperLimitYValues: upperLimitYValues,
      lowerLimitXValues: lowerLimitXValues,
      lowerLimitYValues: lowerLimitYValues,
    };
  } // getLimitOfAgreementData

  getBlandAltmanApsData(
    minX: number,
    maxX: number,
    apsAbs: number,
    apsRel: number,
    diffType: string
  ): {
    apsUpperX: number[][];
    apsUpperY: number[][];
    apsLowerX: number[][];
    apsLowerY: number[][];
  } {
    // This function returns the data for the Bland-Altman Line Series
    // Calculate the APS cut point
    let apsCut = apsAbs / apsRel;

    // Upper APS line
    let apsUpperX: number[][] = [];
    let apsUpperY: number[][] = [];
    // Lower APS line
    let apsLowerX: number[][] = [];
    let apsLowerY: number[][] = [];

    // Calculate the APS lines based on the difference type
    if (diffType === "rel") {
      // Relative difference
      if (minX < apsCut) {
        apsUpperX.push([minX]);
        apsUpperY.push([0 + apsAbs / apsRel]); // difference from zero, not meanDiff
        apsLowerX.push([minX]);
        apsLowerY.push([0 - apsAbs / apsRel]);
      }
      if (maxX < apsCut) {
        apsUpperX.push([maxX]);
        apsUpperY.push([0 + apsAbs / apsRel]);
        apsLowerX.push([maxX]);
        apsLowerY.push([0 - apsAbs / apsRel]);
      } else {
        apsUpperX.push([apsCut]);
        apsUpperY.push([0 + apsRel]);
        apsLowerX.push([apsCut]);
        apsLowerY.push([0 - apsRel]);
        apsUpperX.push([maxX]);
        apsUpperY.push([0 + apsRel]);
        apsLowerX.push([maxX]);
        apsLowerY.push([0 - apsRel]);
      }
    } else {
      // Absolute difference
      if (minX < apsCut) {
        apsUpperX.push([minX]);
        apsUpperY.push([0 + apsAbs]);
        apsLowerX.push([minX]);
        apsLowerY.push([0 - apsAbs]);
      }
      if (maxX < apsCut) {
        apsUpperX.push([maxX]);
        apsUpperY.push([0 + apsAbs]);
        apsLowerX.push([maxX]);
        apsLowerY.push([0 - apsAbs]);
      } else {
        apsUpperX.push([apsCut]);
        apsUpperY.push([0 + apsAbs]);
        apsLowerX.push([apsCut]);
        apsLowerY.push([0 - apsAbs]);
        apsUpperX.push([maxX]);
        apsUpperY.push([0 + apsRel * maxX]);
        apsLowerX.push([maxX]);
        apsLowerY.push([0 - apsRel * maxX]);
      }
    }
    return {
      apsUpperX: apsUpperX,
      apsUpperY: apsUpperY,
      apsLowerX: apsLowerX,
      apsLowerY: apsLowerY,
    };
  } // getBlandAltmanApsData

  /** Create Bland Altman Chart */
  async createChart() {
    await Excel.run(async (context) => {
      // initialise the chart model
      const model = this.initializeModel();

      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      // Range to hold the chart data
      let chartData = currentWorksheet.getRange(this.chartDataRange);
      chartData.load(["address", "rowCount", "columnCount"]);
      await context.sync();

      // Resize the chart data range to fit the Bland-Altman model data
      // The chart data will have four columns: mean X, mean Y, means and differences
      // The number of rows will be the number of means in the Bland-Altman model
      // We will resize the range to fit the data
      chartData = chartData.getResizedRange(
        model.means.length - chartData.rowCount + 1,
        2 - chartData.columnCount
      );
      chartData.load(["values"]);
      await context.sync();

      // Set the chart data values. First row will be the headers
      const headers =
        this.diffType === "rel" ? ["Mean", "Relative Difference"] : ["Mean", "Difference"];

      // Initialize the chart data values
      let data = [];

      // Set the first row as headers
      data.push(headers);

      // Fill the chart data with the means and differences
      for (let i = 0; i < model.means.length; i++) {
        const row = [model.means[i], model.diffs[i]];
        data.push(row);
      }
      chartData.values = data;
      await context.sync();

      let offset = 4; // Offset for additional series data

      // x and y data values
      const xyDataRange = chartData
        .getOffsetRange(0, 2)
        .getAbsoluteResizedRange(this.x.length + 1, 2);
      xyDataRange.load(["address", "values"]);
      await context.sync();
      // Set the x and y values
      let xyData = [];
      xyData.push(["X", "Y"]); // Headers
      for (let i = 0; i < this.x.length; i++) {
        xyData.push([this.x[i], this.y[i]]);
      }
      xyDataRange.values = xyData;
      await context.sync();

      // Range for mean difference chart data
      const meanDiffXRange = chartData.getOffsetRange(1, offset).getAbsoluteResizedRange(2, 1);
      const meanDiffYRange = chartData.getOffsetRange(1, offset + 1).getAbsoluteResizedRange(2, 1);
      meanDiffXRange.load(["address", "values"]);
      meanDiffYRange.load(["address", "values"]);
      await context.sync();

      // Set the mean difference values
      meanDiffXRange.values = model.meanDiffData.meanDiffXValues;
      meanDiffYRange.values = model.meanDiffData.meanDiffYValues;
      await context.sync();

      //Upper limits of agreement
      const upperLimitXRange = chartData
        .getOffsetRange(1, offset + 2)
        .getAbsoluteResizedRange(2, 1);
      const upperLimitYRange = chartData
        .getOffsetRange(1, offset + 3)
        .getAbsoluteResizedRange(2, 1);
      upperLimitXRange.load(["address", "values"]);
      upperLimitYRange.load(["address", "values"]);
      await context.sync();

      // Set the upper limit values
      upperLimitXRange.values = model.loaData.upperLimitXValues;
      upperLimitYRange.values = model.loaData.upperLimitYValues;
      await context.sync();

      // Lower limits of agreement
      const lowerLimitXRange = chartData
        .getOffsetRange(1, offset + 4)
        .getAbsoluteResizedRange(2, 1);
      const lowerLimitYRange = chartData
        .getOffsetRange(1, offset + 5)
        .getAbsoluteResizedRange(2, 1);
      lowerLimitXRange.load(["address", "values"]);
      lowerLimitYRange.load(["address", "values"]);
      await context.sync();

      // Set the lower limit values
      lowerLimitXRange.values = model.loaData.lowerLimitXValues;
      lowerLimitYRange.values = model.loaData.lowerLimitYValues;
      await context.sync();

      // Create the Bland-Altman chart
      // The chart will be a scatter plot with the means on the x-axis and the differences on the y-axis
      // We will use the chartData range as the data source for the chart
      const chart = currentWorksheet.charts.add(
        Excel.ChartType.xyscatter,
        chartData,
        Excel.ChartSeriesBy.columns
      );
      chart.title.text = "Difference Chart";
      chart.legend.position = Excel.ChartLegendPosition.bottom;
      chart.legend.visible = true;
      chart.axes.categoryAxis.title.text = headers[0]; //Mean
      chart.axes.valueAxis.title.text = headers[1]; //Difference or Relative Difference

      // Add the mean difference line
      const meanDiffSeries = chart.series.add("Mean Difference");
      meanDiffSeries.setValues(meanDiffYRange);
      meanDiffSeries.setXAxisValues(meanDiffXRange);
      meanDiffSeries.format.line.color = "blue"; // Set the line color to red
      meanDiffSeries.format.line.lineStyle = Excel.ChartLineStyle.continuous; // Set the line style to solid
      meanDiffSeries.format.line.weight = 2; // Set the line weight
      meanDiffSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers

      // Add the upper limit of agreement line
      const upperLimitSeries = chart.series.add("Upper Limit of Agreement");
      upperLimitSeries.setValues(upperLimitYRange);
      upperLimitSeries.setXAxisValues(upperLimitXRange);
      upperLimitSeries.format.line.color = "green"; // Set the line color to green
      upperLimitSeries.format.line.lineStyle = Excel.ChartLineStyle.dot; // Set the line style to dashed
      upperLimitSeries.format.line.weight = 1; // Set the line weight
      upperLimitSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers

      // Add the lower limit of agreement line
      const lowerLimitSeries = chart.series.add("Lower Limit of Agreement");
      lowerLimitSeries.setValues(lowerLimitYRange);
      lowerLimitSeries.setXAxisValues(lowerLimitXRange);
      lowerLimitSeries.format.line.color = "green"; // Set the line color to green
      lowerLimitSeries.format.line.lineStyle = Excel.ChartLineStyle.dot; // Set the line style to dot
      lowerLimitSeries.format.line.weight = 1; // Set the line weight
      lowerLimitSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers

      if (model.apsData !== null) {
        // Upper APS line
        const apsUpperXRange = chartData
          .getOffsetRange(1, offset + 6)
          .getAbsoluteResizedRange(model.apsData.apsUpperX.length, 1);
        const apsUpperYRange = chartData
          .getOffsetRange(1, offset + 7)
          .getAbsoluteResizedRange(model.apsData.apsUpperY.length, 1);
        apsUpperXRange.load(["address", "values"]);
        apsUpperYRange.load(["address", "values"]);
        await context.sync();

        apsUpperXRange.values = model.apsData.apsUpperX;
        apsUpperYRange.values = model.apsData.apsUpperY;
        await context.sync();

        const apsUpperSeries = chart.series.add("Upper APS");
        apsUpperSeries.setValues(apsUpperYRange);
        apsUpperSeries.setXAxisValues(apsUpperXRange);
        apsUpperSeries.format.line.color = "red"; // Set the line color to purple
        apsUpperSeries.format.line.lineStyle = Excel.ChartLineStyle.dash; // Set the line style to dashed
        apsUpperSeries.format.line.weight = 1; // Set the line weight
        apsUpperSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers

        // Lower APS line
        const apsLowerXRange = chartData
          .getOffsetRange(1, offset + 8)
          .getAbsoluteResizedRange(model.apsData.apsLowerX.length, 1);
        const apsLowerYRange = chartData
          .getOffsetRange(1, offset + 9)
          .getAbsoluteResizedRange(model.apsData.apsLowerY.length, 1);
        apsLowerXRange.load(["address", "values"]);
        apsLowerYRange.load(["address", "values"]);
        await context.sync();

        apsLowerXRange.values = model.apsData.apsLowerX;
        apsLowerYRange.values = model.apsData.apsLowerY;
        await context.sync();

        const apsLowerSeries = chart.series.add("Lower APS");
        apsLowerSeries.setValues(apsLowerYRange);
        apsLowerSeries.setXAxisValues(apsLowerXRange);
        apsLowerSeries.format.line.color = "red"; // Set the line color to purple
        apsLowerSeries.format.line.lineStyle = Excel.ChartLineStyle.dash; // Set the line style to dashed
        apsLowerSeries.format.line.weight = 1; // Set the line weight
        apsLowerSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers
      } // if apsData

      // Scale value axis from -0.5 to 0.5 if the difference type is relative
      if (this.diffType === "rel") {
        chart.axes.valueAxis.minimum = -0.5;
        chart.axes.valueAxis.maximum = 0.5;
      }

      // Set the chart position
      if (this.outputRange !== "") {
        //console.log(`Setting chart position to ${this.outputRange}`);
        const chartCell = this.outputRange.split(":");
        chart.setPosition(chartCell[0], chartCell[1]);
        await context.sync();
      }
    });
  }
} // ExcelBlandAltmanChart

class ExcelRegressionChart {
  private x: number[];
  private y: number[];
  private slope: number;
  private intercept: number;
  private apsAbs: number;
  private apsRel: number;
  private chartDataRange: string;
  private outputRange: string;

  constructor(
    x: number[],
    y: number[],
    slope: number,
    intercept: number,
    apsAbs: number,
    apsRel: number,
    chartDataRange: string,
    outputRange = ""
  ) {
    //console.log("RegressionChart initialized");
    this.x = x; // Array of numbers
    this.y = y; // Array of numbers
    this.slope = slope; // Slope of the regression line
    this.intercept = intercept; // Intercept of the regression line
    this.apsAbs = apsAbs; // Absolute total error specification
    this.apsRel = apsRel; // Relative total error specification
    this.chartDataRange = chartDataRange; // Cell range for chart data
    this.outputRange = outputRange; // Cell range for output
  }

  initializeModel(): {
    minX: number;
    maxX: number;
    identityX: number[][];
    identityY: number[][];
    regressionX: number[][];
    regressionY: number[][];
    apsData: {
      apsUpperX: number[][];
      apsUpperY: number[][];
      apsLowerX: number[][];
      apsLowerY: number[][];
    } | null;
  } {
    const minX = Math.min(...this.x);
    const maxX = Math.max(...this.x);
    const identityX = [[minX], [maxX]]; //2 rows, 1 column
    const identityY = [[minX], [maxX]]; //2 rows, 1 column
    const regressionX = [[minX], [maxX]]; //2 rows, 1 column
    const regressionY = [
      [this.intercept + this.slope * minX],
      [this.intercept + this.slope * maxX],
    ]; //2 rows, 1 column
    let apsData = null;
    if (!(this.apsAbs <= 0 && this.apsRel <= 0)) {
      // Calculate the APS lines only if APS values are provided
      apsData = this.getApsData(minX, maxX, this.apsAbs, this.apsRel);
    }
    return {
      minX: minX,
      maxX: maxX,
      identityX: identityX,
      identityY: identityY,
      regressionX: regressionX,
      regressionY: regressionY,
      apsData: apsData,
    };
  }

  getApsData(
    minX: number,
    maxX: number,
    apsAbs: number,
    apsRel: number
  ): {
    apsUpperX: number[][];
    apsUpperY: number[][];
    apsLowerX: number[][];
    apsLowerY: number[][];
  } {
    // This function returns the data for the APS Line Series
    // Calculate the APS cut point
    let apsCut = apsAbs / apsRel;
    let apsUpperX: number[][] = [];
    let apsUpperY: number[][] = [];
    let apsLowerX: number[][] = [];
    let apsLowerY: number[][] = [];
    if (minX < apsCut) {
      apsUpperX.push([minX]);
      apsUpperY.push([minX + apsAbs]);
      apsLowerX.push([minX]);
      apsLowerY.push([minX - apsAbs]);
    }
    if (maxX < apsCut) {
      apsUpperX.push([maxX]);
      apsUpperY.push([maxX + apsAbs]);
      apsLowerX.push([maxX]);
      apsLowerY.push([maxX - apsAbs]);
    } else {
      apsUpperX.push([apsCut]);
      apsUpperY.push([apsCut + apsAbs]);
      apsLowerX.push([apsCut]);
      apsLowerY.push([apsCut - apsAbs]);
      apsUpperX.push([maxX]);
      apsUpperY.push([maxX + apsRel * maxX]);
      apsLowerX.push([maxX]);
      apsLowerY.push([maxX - apsRel * maxX]);
    }
    return {
      apsUpperX: apsUpperX,
      apsUpperY: apsUpperY,
      apsLowerX: apsLowerX,
      apsLowerY: apsLowerY,
    };
  } // getApsData

  /** Create Regression Chart */
  async createChart() {
    await Excel.run(async (context) => {
      // initialise the chart model
      const model = this.initializeModel();

      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      // Range to hold the chart data
      let chartData = currentWorksheet.getRange(this.chartDataRange);
      chartData.load(["address", "rowCount", "columnCount"]);
      await context.sync();

      // Resize the chart data range to fit the Bland-Altman model data
      // The chart data will have four columns: mean X, mean Y, means and differences
      // The number of rows will be the number of means in the Bland-Altman model
      // We will resize the range to fit the data
      chartData = chartData.getOffsetRange(0, 2).getAbsoluteResizedRange(this.x.length + 1, 2);
      chartData.load(["values"]);
      await context.sync();

      // Set the chart data values. First row will be the headers
      const headers = ["X", "Y"];

      // Initialize the chart data values
      let data = [];

      // Set the first row as headers
      data.push(headers);

      // Fill the chart data with the means and differences
      for (let i = 0; i < this.x.length; i++) {
        const row = [this.x[i], this.y[i]];
        data.push(row);
      }
      chartData.values = data;
      await context.sync();

      let offset = 4; // Offset for additional series data

      // Range for identity line chart data
      const identityXRange = chartData.getOffsetRange(10, offset).getAbsoluteResizedRange(2, 1);
      const identityYRange = chartData.getOffsetRange(10, offset + 1).getAbsoluteResizedRange(2, 1);
      identityXRange.load(["address", "values"]);
      identityYRange.load(["address", "values"]);
      await context.sync();
      // Set the identity line values
      identityXRange.values = model.identityX;
      identityYRange.values = model.identityY;
      await context.sync();
      // Range for regression line chart data
      const regressionXRange = chartData
        .getOffsetRange(10, offset + 2)
        .getAbsoluteResizedRange(2, 1);
      const regressionYRange = chartData
        .getOffsetRange(10, offset + 3)
        .getAbsoluteResizedRange(2, 1);
      regressionXRange.load(["address", "values"]);
      regressionYRange.load(["address", "values"]);
      await context.sync();
      // Set the regression line values
      regressionXRange.values = model.regressionX;
      regressionYRange.values = model.regressionY;
      await context.sync();

      // Create a scatter chart
      const chart = currentWorksheet.charts.add(
        Excel.ChartType.xyscatter,
        chartData,
        Excel.ChartSeriesBy.columns
      );
      chart.title.text = "Scatter Chart";
      chart.legend.position = Excel.ChartLegendPosition.bottom;
      chart.legend.visible = true;
      chart.axes.categoryAxis.title.text = "Reference Method";
      chart.axes.valueAxis.title.text = "Comparison Method";

      // Add the identity line
      const identitySeries = chart.series.add("Identity Line");
      identitySeries.setValues(identityYRange);
      identitySeries.setXAxisValues(identityXRange);
      identitySeries.format.line.color = "black"; // Set the line color to black
      identitySeries.format.line.lineStyle = Excel.ChartLineStyle.continuous; // Set the line style to solid
      identitySeries.format.line.weight = 1; // Set the line weight
      identitySeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers
      // Add the regression line
      const regressionSeries = chart.series.add("Regression Line");
      regressionSeries.setValues(regressionYRange);
      regressionSeries.setXAxisValues(regressionXRange);
      regressionSeries.format.line.color = "blue"; // Set the line color to blue
      regressionSeries.format.line.lineStyle = Excel.ChartLineStyle.continuous; // Set the line style to solid
      regressionSeries.format.line.weight = 2; // Set the line weight
      regressionSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers

      // Add the APS lines if provided
      if (model.apsData !== null) {
        // Upper APS line
        const apsUpperXRange = chartData
          .getOffsetRange(10, offset + 4)
          .getAbsoluteResizedRange(model.apsData.apsUpperX.length, 1);
        const apsUpperYRange = chartData
          .getOffsetRange(10, offset + 5)
          .getAbsoluteResizedRange(model.apsData.apsUpperY.length, 1);
        apsUpperXRange.load(["address", "values"]);
        apsUpperYRange.load(["address", "values"]);
        await context.sync();

        apsUpperXRange.values = model.apsData.apsUpperX;
        apsUpperYRange.values = model.apsData.apsUpperY;
        await context.sync();

        const apsUpperSeries = chart.series.add("Upper APS");
        apsUpperSeries.setValues(apsUpperYRange);
        apsUpperSeries.setXAxisValues(apsUpperXRange);
        apsUpperSeries.format.line.color = "red"; // Set the line color to purple
        apsUpperSeries.format.line.lineStyle = Excel.ChartLineStyle.dash; // Set the line style to dashed
        apsUpperSeries.format.line.weight = 1; // Set the line weight
        apsUpperSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers

        // Lower APS line
        const apsLowerXRange = chartData
          .getOffsetRange(10, offset + 6)
          .getAbsoluteResizedRange(model.apsData.apsLowerX.length, 1);
        const apsLowerYRange = chartData
          .getOffsetRange(10, offset + 7)
          .getAbsoluteResizedRange(model.apsData.apsLowerY.length, 1);
        apsLowerXRange.load(["address", "values"]);
        apsLowerYRange.load(["address", "values"]);
        await context.sync();

        apsLowerXRange.values = model.apsData.apsLowerX;
        apsLowerYRange.values = model.apsData.apsLowerY;
        await context.sync();

        const apsLowerSeries = chart.series.add("Lower APS");
        apsLowerSeries.setValues(apsLowerYRange);
        apsLowerSeries.setXAxisValues(apsLowerXRange);
        apsLowerSeries.format.line.color = "red"; // Set the line color to purple
        apsLowerSeries.format.line.weight = 1; // Set the line weight
        apsLowerSeries.format.line.lineStyle = Excel.ChartLineStyle.dash; // Set the line style to dashed
        apsLowerSeries.markerStyle = Excel.ChartMarkerStyle.none; // Hide the markers
      } // if apsData

      // Set the chart position
      if (this.outputRange !== "") {
        //console.log(`Setting chart position to ${this.outputRange}`);
        const chartCell = this.outputRange.split(":");
        chart.setPosition(chartCell[0], chartCell[1]);
        await context.sync();
      }
    });
  }
} //ExcelRegressionChart

// module exports
export { ExcelBlandAltmanChart, ExcelRegressionChart };
