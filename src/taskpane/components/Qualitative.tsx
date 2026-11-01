/*
 * Panel for analysing concordance of qualitative data
 *
 * Author: Douglas Chesher
 *
 * Created: October 2025.
 */
import * as React from "react";

import { makeStyles, Button, tokens } from "@fluentui/react-components";
import RangeInput from "./RangeInput";

import { QualitativeContengencyTableBuilder, ConcordanceCalculator } from "../../concordance";

interface QualitativeProps {
  notify: (intent: string, message: string) => void;
  qualXRangeValue: string;
  qualYRangeValue: string;
  qualOutputRangeValue: string;
  setQualXRangeValue: React.Dispatch<React.SetStateAction<string>>;
  setQualYRangeValue: React.Dispatch<React.SetStateAction<string>>;
  setQualOutputRangeValue: React.Dispatch<React.SetStateAction<string>>;
  uitext: { [key: string]: string };
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  field: {
    marginLeft: tokens.spacingHorizontalXS,
    marginBottom: tokens.spacingVerticalXS,
    marginTop: tokens.spacingVerticalXS,
  },
});

export default function Qualitative(props: QualitativeProps) {
  const styles = useStyles();

  // Run the qualitative data comparison
  const qualComparison = async () => {
    await Excel.run(async (context) => {
      //console.log("Running qualitative data comparison...");
      try {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

        // Load the ranges
        const xRange = currentWorksheet.getRange(props.qualXRangeValue);
        const yRange = currentWorksheet.getRange(props.qualYRangeValue);
        xRange.load(["rowCount", "columnCount", "values"]);
        yRange.load(["rowCount", "columnCount", "values"]);
        await context.sync();

        if (xRange.rowCount !== yRange.rowCount) {
          throw new Error("X and Y ranges must have the same number of rows");
        }
        if (xRange.columnCount !== 1 || yRange.columnCount !== 1) {
          throw new Error("X and Y ranges must each contain exactly one column");
        }

        const x: string[] = [];
        const y: string[] = [];
        for (let i = 0; i < xRange.rowCount; i++) {
          const xValue = xRange.values[i][0];
          const yValue = yRange.values[i][0];
          const xIsBlank = xValue === null || xValue === "";
          const yIsBlank = yValue === null || yValue === "";
          //console.log(`x = ${xValue}; y = ${yValue}`);

          if (xIsBlank && yIsBlank) {
            continue;
          }
          if (xIsBlank || yIsBlank) {
            throw new Error(`Both X and Y values must be present in row ${i + 1}`);
          }
          if (
            (typeof xValue !== "string" && typeof xValue !== "number") ||
            (typeof yValue !== "string" && typeof yValue !== "number")
          ) {
            throw new Error(`X and Y values in row ${i + 1} must be text or numbers`);
          }

          x.push(xValue.toString());
          y.push(yValue.toString());
        }
        if (x.length === 0) {
          throw new Error("X and Y ranges must contain at least one paired value");
        }
        const contingencyBuilder = new QualitativeContengencyTableBuilder();
        const contingencyTable = contingencyBuilder.build(x, y);
        const concordanceCalculator = new ConcordanceCalculator(contingencyTable);
        const concordanceResults = concordanceCalculator.calculate();
        //to-do: copy to Excel
        const concordanceCell = props.qualOutputRangeValue;
        if (concordanceCell === "") {
          throw Error("Please select the concordance output range.");
        } else {
          const concordanceRange = currentWorksheet.getRange(concordanceCell);
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
      } catch (err) {
        if (err instanceof Error) {
          props.notify("error", err.message);
        }
      } // try
    });
  };

  return (
    <div className="container">
      <div>Assess concordance of qualitative data and calculate Cohen's Kappa.</div>
      <div>
        <RangeInput
          label={props.uitext["lbl_x_range"]}
          rangeValue={props.qualXRangeValue}
          setRangeValue={props.setQualXRangeValue}
          validationMessage={props.uitext["msg_x_range"]}
          uitext={props.uitext}
        />
      </div>
      <div>
        <RangeInput
          label={props.uitext["lbl_y_range"]}
          rangeValue={props.qualYRangeValue}
          setRangeValue={props.setQualYRangeValue}
          validationMessage={props.uitext["msg_y_range"]}
          uitext={props.uitext}
        />
      </div>

      <div>
        <RangeInput
          label={props.uitext["lbl_output_range"]}
          rangeValue={props.qualOutputRangeValue}
          setRangeValue={props.setQualOutputRangeValue}
          validationMessage={props.uitext["msg_output_range"]}
          uitext={props.uitext}
        />
      </div>

      <div className={styles.field}>
        <Button appearance="primary" onClick={qualComparison}>
          {props.uitext["btn_run"]}
        </Button>
      </div>
    </div>
  );
}
