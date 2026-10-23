import { RefLimResult, iqr } from "./reflim";
import { max, median, min, arange, seq, histogram, normal } from "jstat-esm"
/**
 * Adds a combined histogram/XY chart and a Q-Q plot below it.
 *
 * Chart data is written to the right of the chart ranges.
 */
export async function addRefLimCharts(
    startCellAddress: string,
    binValues: readonly number[],
    binCounts: readonly number[],
    xValues: readonly number[],
    yValues: readonly number[],
    res: RefLimResult
): Promise<void> {
    if (binValues.length === 0 || binValues.length !== binCounts.length) {
        throw new Error("Histogram bin values and counts must have the same non-zero length.");
    }
    if (xValues.length === 0 || xValues.length !== yValues.length) {
        throw new Error("XY x-values and y-values must have the same non-zero length.");
    }

    if (![binValues, binCounts, xValues, yValues].every((values) => values.every(Number.isFinite))) {
        throw new Error("Chart data must contain only finite numbers.");
    }
    // Validate Q-Q plot data
    const qqData = res.res_qq?.chart_data;
    if (
        !qqData ||
        qqData.data_x.length === 0 ||
        qqData.data_x.length !== qqData.data_y.length ||
        ![qqData.data_x, qqData.data_y].every((values) => values.every(Number.isFinite))
    ) {
        throw new Error("Q-Q chart data must contain matching, non-empty finite x and y values.");
    }
    if (
        res.stats.mean === undefined ||
        res.stats.sd === undefined ||
        !Number.isFinite(res.stats.mean) ||
        !Number.isFinite(res.stats.sd)
    ) {
        throw new Error("Q-Q trendline requires finite reference mean and standard deviation values.");
    }
    const trendlineIntercept = res.stats.mean;
    const trendlineSlope = res.stats.sd;

    // Scale the y-values of the XY data to match the histogram's y-axis range
    const minCount = binCounts.reduce((minimum, count) => Math.min(minimum, count), binCounts[0]);
    const maxCount = binCounts.reduce((maximum, count) => Math.max(maximum, count), binCounts[0]);
    const minY = yValues.reduce((minimum, y) => Math.min(minimum, y), yValues[0]);
    const maxY = yValues.reduce((maximum, y) => Math.max(maximum, y), yValues[0]);
    const scaledYValues =
        minY === maxY
            ? yValues.map(() => (minCount + maxCount) / 2)
            : yValues.map((y) => minCount + ((y - minY) / (maxY - minY)) * (maxCount - minCount));
    const limits = [
        { name: "Lower limit", value: res.limits.lower_lim, columnOffset: 12 },
        { name: "Upper limit", value: res.limits.upper_lim, columnOffset: 14 },
    ].filter(
        (limit): limit is { name: string; value: number; columnOffset: number } =>
            limit.value !== undefined && Number.isFinite(limit.value)
    );

    // Run the Excel operations in a single batch
    await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const chartRange = worksheet
            .getRange(startCellAddress)
            .getCell(0, 0)
            .getResizedRange(11, 7);
        const histogramRange = worksheet
            .getRange(startCellAddress)
            .getCell(0, 0)
            .getOffsetRange(0, 8)
            .getAbsoluteResizedRange(binValues.length + 1, 2);
        const xyRange = worksheet
            .getRange(startCellAddress)
            .getCell(0, 0)
            .getOffsetRange(0, 10)
            .getAbsoluteResizedRange(xValues.length + 1, 2);
        const limitRanges = limits.map((limit) => ({
            ...limit,
            range: worksheet
                .getRange(startCellAddress)
                .getCell(0, 0)
                .getOffsetRange(0, limit.columnOffset)
                .getAbsoluteResizedRange(3, 2),
        }));
        const qqRange = worksheet
            .getRange(startCellAddress)
            .getCell(0, 0)
            .getOffsetRange(0, 16)
            .getAbsoluteResizedRange(qqData.data_x.length + 1, 2);
        const qqTrendlineRange = worksheet
            .getRange(startCellAddress)
            .getCell(0, 0)
            .getOffsetRange(0, 18)
            .getAbsoluteResizedRange(3, 2);
        const qqChartRange = worksheet
            .getRange(startCellAddress)
            .getCell(0, 0)
            .getOffsetRange(12, 0)
            .getResizedRange(11, 7);

        histogramRange.values = [
            ["Bin", "Count"],
            ...binValues.map((bin, index) => [bin, binCounts[index]]),
        ];
        xyRange.values = [
            ["X", "Y (scaled)"],
            ...xValues.map((x, index) => [x, scaledYValues[index]]),
        ];
        limitRanges.forEach(({ name, value, range }) => {
            range.values = [
                [`${name} X`, `${name} Y`],
                [value, 0],
                [value, maxCount],
            ];
        });
        qqRange.values = [
            [qqData.xlab, qqData.ylab],
            ...qqData.data_x.map((x, index) => [x, qqData.data_y[index]]),
        ];
        const trendlineX =
            qqData.xlim.length === 2 && qqData.xlim.every(Number.isFinite)
                ? qqData.xlim
                : [
                    qqData.data_x.reduce((minimum, x) => Math.min(minimum, x), qqData.data_x[0]),
                    qqData.data_x.reduce((maximum, x) => Math.max(maximum, x), qqData.data_x[0]),
                ];
        qqTrendlineRange.values = [
            ["Trendline X", "Trendline Y"],
            ...trendlineX.map((x) => [x, trendlineIntercept + trendlineSlope * x]),
        ];
        await context.sync();

        // Create the combined histogram/XY chart
        const chart = worksheet.charts.add(
            Excel.ChartType.xyscatter,
            histogramRange.getOffsetRange(0, 1).getAbsoluteResizedRange(binValues.length + 1, 1),
            Excel.ChartSeriesBy.columns
        );
        chart.title.text = "Histogram";
        chart.legend.visible = true;
        chart.legend.position = Excel.ChartLegendPosition.bottom;
        chart.axes.categoryAxis.title.text = "Bin / X";
        chart.axes.valueAxis.title.text = "Count";

        const histogramSeries = chart.series.getItemAt(0);
        histogramSeries.name = "Count";
        histogramSeries.setValues(
            histogramRange.getOffsetRange(1, 1).getAbsoluteResizedRange(binCounts.length, 1)
        );
        histogramSeries.setXAxisValues(
            histogramRange.getOffsetRange(1, 0).getAbsoluteResizedRange(binValues.length, 1)
        );
        histogramSeries.markerStyle = Excel.ChartMarkerStyle.dash;
        histogramSeries.markerSize = 5;

        const xySeries = chart.series.add("XY");
        xySeries.setValues(xyRange.getOffsetRange(1, 1).getAbsoluteResizedRange(yValues.length, 1));
        xySeries.setXAxisValues(xyRange.getOffsetRange(1, 0).getAbsoluteResizedRange(xValues.length, 1));
        xySeries.chartType = Excel.ChartType.xyscatter;
        xySeries.markerStyle = Excel.ChartMarkerStyle.none;
        xySeries.format.line.lineStyle = Excel.ChartLineStyle.continuous;
        xySeries.format.line.weight = 1;

        limitRanges.forEach(({ name, range }) => {
            const series = chart.series.add(name);
            series.setValues(range.getOffsetRange(1, 1).getAbsoluteResizedRange(2, 1));
            series.setXAxisValues(range.getOffsetRange(1, 0).getAbsoluteResizedRange(2, 1));
            series.chartType = Excel.ChartType.xyscatter;
            series.markerStyle = Excel.ChartMarkerStyle.none;
            series.format.line.lineStyle = Excel.ChartLineStyle.dot;
            series.format.line.weight = 1;
        });

        chart.setPosition(chartRange);

        // Create the Q-Q plot chart
        const qqChart = worksheet.charts.add(
            Excel.ChartType.xyscatter,
            qqRange.getOffsetRange(0, 1),
            Excel.ChartSeriesBy.columns
        );
        qqChart.title.text = qqData.main;
        qqChart.legend.visible = false;
        qqChart.axes.categoryAxis.title.text = qqData.xlab;
        qqChart.axes.valueAxis.title.text = qqData.ylab;
        const qqSeries = qqChart.series.getItemAt(0);
        qqSeries.name = qqData.ylab;
        qqSeries.setValues(qqRange.getOffsetRange(1, 1).getAbsoluteResizedRange(qqData.data_y.length, 1));
        qqSeries.setXAxisValues(qqRange.getOffsetRange(1, 0).getAbsoluteResizedRange(qqData.data_x.length, 1));
        qqSeries.chartType = Excel.ChartType.xyscatter;
        qqSeries.markerStyle = Excel.ChartMarkerStyle.circle;

        const qqTrendlineSeries = qqChart.series.add("Reference line");
        qqTrendlineSeries.setValues(qqTrendlineRange.getOffsetRange(1, 1).getAbsoluteResizedRange(2, 1));
        qqTrendlineSeries.setXAxisValues(qqTrendlineRange.getOffsetRange(1, 0).getAbsoluteResizedRange(2, 1));
        qqTrendlineSeries.chartType = Excel.ChartType.xyscatter;
        qqTrendlineSeries.markerStyle = Excel.ChartMarkerStyle.none;
        qqTrendlineSeries.format.line.lineStyle = Excel.ChartLineStyle.continuous;
        qqTrendlineSeries.format.line.weight = 1;
        if (qqData.xlim.length === 2 && qqData.xlim.every(Number.isFinite)) {
            qqChart.axes.categoryAxis.minimum = qqData.xlim[0];
            qqChart.axes.categoryAxis.maximum = qqData.xlim[1];
        }
        if (qqData.ylim.length === 2 && qqData.ylim.every(Number.isFinite)) {
            qqChart.axes.valueAxis.minimum = qqData.ylim[0];
            qqChart.axes.valueAxis.maximum = qqData.ylim[1];
        }
        qqChart.setPosition(qqChartRange);
        await context.sync();
    });
}

/**
 * Gets the histogram data for the reference limits chart
 * @param x 
 * @param res 
 * @param remove_extremes 
 * @returns 
 */
export function getHistogramData(x: number[], res: RefLimResult, remove_extremes: boolean): { breaks: number[], hist: number[], from: number, to: number } {
    let xx: number[] = [];
    const m = median(x);
    const r = iqr(x);
    if (remove_extremes) {
        x.forEach(x => {
            if (x < m + 8 * r) xx.push(x);
        });
    } else {
        xx = x;
    }
    const diff = max(xx) - min(xx);
    let from = min(xx) - 0.1 * diff;
    if (from < 0) from = 0;
    let to = max(xx) + 0.1 * diff;
    if (res.limits.upper_lim != undefined && to < res.limits.upper_lim) {
        to < - res.limits.upper_lim + 0.1 * diff;
    }
    let breaks: number[] = [];
    let bin_width = 0;
    if (res.limits.upper_lim != null && res.limits.lower_lim != null) {
        bin_width = (res.limits.upper_lim - res.limits.lower_lim) / 10;
        breaks = arange(min(x), max(x), bin_width);
    }
    let hist = histogram(x, breaks.length);
    const bins = breaks.map(b => b + bin_width / 2);
    // Remove breaks and hist values that are outside the range of from and to
    const filteredBreaks = bins.filter(b => b >= from && b <= to);
    const filteredHist = hist.slice(bins.indexOf(filteredBreaks[0]), bins.indexOf(filteredBreaks[filteredBreaks.length - 1]) + 1);
    return { breaks: filteredBreaks, hist: filteredHist, from, to };
}

/**
 * Gets the X and Y data for the reference limits chart
 * @param res 
 * @param from 
 * @param to 
 * @returns 
 */
export function getXYData(res: RefLimResult, from: number, to: number): { x: number[], y: number[] } {
    //x values are from lower limit to upper limit,
    //y values are the density values from a normal or log normal distribution with the mean and sd from the results
    const x: number[] = seq(from, to, 300);
    let y: number[] = [];
    if (res.lognormal) {
        y = x.map(x => normal.pdf(Math.log(x), res.stats.mean, res.stats.sd));
    } else {
        y = x.map(x => normal.pdf(x, res.stats.mean, res.stats.sd));
    }
    // Remove x and y values that are outside the range of from and to
    const filteredX = x.filter(x => x >= from && x <= to);
    const filteredY = y.slice(x.indexOf(filteredX[0]), x.indexOf(filteredX[filteredX.length - 1]) + 1);
    return { x: filteredX, y: filteredY };
}
