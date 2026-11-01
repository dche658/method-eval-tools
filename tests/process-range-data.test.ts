import { expect, test } from "vitest";
import { processRangeData } from "../src/taskpane/components/Regression";

test("processes numeric observations from a single-column range", () => {
  const result = processRangeData({
    rowCount: 3,
    columnCount: 1,
    values: [[1], [""], [3]],
  });

  expect(result.means).toEqual([1, 3]);
  expect(result.size).toBe(2);
  expect(result.mean).toBe(2);
});

test("calculates replicate statistics from a two-column range", () => {
  const result = processRangeData({
    rowCount: 2,
    columnCount: 2,
    values: [
      [1, 3],
      [2, 4],
    ],
  });

  expect(result.means).toEqual([2, 3]);
  expect(result.size).toBe(2);
  expect(result.sd).toBeCloseTo(2);
  expect(result.cv).toBeCloseTo(0.8);
});

test.each([
  {
    description: "no numeric observations",
    values: [[""], [""]],
  },
  {
    description: "only one numeric observation",
    values: [[1], [""]],
  },
])("rejects a range with $description", ({ values }) => {
  expect(() =>
    processRangeData({
      rowCount: values.length,
      columnCount: 1,
      values,
    })
  ).toThrow("At least two numeric observations are required.");
});

test.each([Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY])(
  "rejects a non-finite numeric cell (%s)",
  (value) => {
    expect(() =>
      processRangeData({
        rowCount: 2,
        columnCount: 1,
        values: [[1], [value]],
      })
    ).toThrow("Range contains a non-finite number at row 2, column 1.");
  }
);

test("rejects a non-finite value in either replicate column", () => {
  expect(() =>
    processRangeData({
      rowCount: 2,
      columnCount: 2,
      values: [
        [1, 2],
        [3, Number.NaN],
      ],
    })
  ).toThrow("Range contains a non-finite number at row 2, column 2.");
});
