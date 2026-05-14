/**
 * Comparison Pane
 *
 * The Bland-Altman and Regression analysis was developed first so those
 * components appear hear. To reduce the complexity and size of this file
 * the other functions were moved to separate subcomponents but the properties
 * for the complete add-in are all defined here as they are in many cases
 * required to be used in more than one subcomponent.
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
  Button,
  Checkbox,
  CheckboxProps,
  Field,
  Input,
  Card,
  Tooltip,
  tokens
} from "@fluentui/react-components";

const PrecisionLayout = React.lazy(() => import("./PrecisionLayout"));
const Qualitative = React.lazy(() => import("./Qualitative"));
const Regression = React.lazy(() => import("./Regression"));
const Precision = React.lazy(() => import("./Precision"));

// Properties for this component
interface ComparisonProps {
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
    marginLeft: tokens.spacingHorizontalXS,
    marginBottom: tokens.spacingVerticalXS,
    marginTop: tokens.spacingVerticalXS,
  },
  label: {
    fontWeight: "bold",
  },
});



export default function ComparisonPane(props: ComparisonProps) {
  const styles = useStyles();
  const [apsAbsValue, setApsAbsValue] = React.useState<string>("");
  const [apsRelValue, setApsRelValue] = React.useState<string>("");
  const [xRangeValue, setXRangeValue] = React.useState<string>("");
  const [yRangeValue, setYRangeValue] = React.useState<string>("");
  const [compOutRangeValue, setCompOutRangeValue] = React.useState<string>("");
  const [baRangeValue, setBaRangeValue] = React.useState<string>(""); // Bland-Altman output range
  const [scRangeValue, setScRangeValue] = React.useState<string>(""); // Scatter chart output range
  const [cdRangeValue, setCdRangeValue] = React.useState<string>(""); // Chart data output range
  const [errorRatio, setErrorRatio] = React.useState<string>("1.0");
  const [regressionType, setRegressionType] = React.useState<string>("paba"); // Default to "paba" (Passing-Bablok)
  const [useCalcErrorRatio, setUseCalcErrorRatio] = React.useState<boolean>(false); // Default to false
  const [importApsSettings, setImportApsSettings] = React.useState<boolean>(false); // Default to false
  const [labelOutput, setLabelOutput] = React.useState<boolean>(true); // Default to true
  const [differencePlotType, setDifferencePlotType] = React.useState<string>("rel"); // Default to "rel"
  const [concordanceOutputRange, setConcordanceOutputRange] = React.useState<string>("");
  const [ciMethod, setCiMethod] = React.useState<string>("default"); // Default to "bootstrap"
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
  const [importLayoutChecked, setImportLayoutChecked] =
    React.useState<CheckboxProps["checked"]>(false);

  const regressionTypeId = useId("regression-type");

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
        props.notify("success", "Ranges have been successfully updated");
      } catch (error) {
        if (error instanceof Error) {
          props.notify("error", error.message);
        }
      }
    }
  }; // setDefaultValues

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
        props.notify("success", "Ranges have been successfully updated");
      } catch (error) {
        if (error instanceof Error) {
          props.notify("error", error.message);
        }
      }
    });
  }; // loadWorkbookDefaults

  return (
    <Accordion collapsible={true}>
      <AccordionItem value="2">
        <AccordionHeader>{props.uitext["h_method_comparison"]}</AccordionHeader>
        <AccordionPanel>
          <Card>
            <React.Suspense fallback={<div>Loading...</div>}>
              <Regression
                apsAbsValue={apsAbsValue}
                apsRelValue={apsRelValue}
                xRangeValue={xRangeValue}
                yRangeValue={yRangeValue}
                importApsSettings={importApsSettings}
                setImportApsSettings={setImportApsSettings}
                setXRangeValue={setXRangeValue}
                setYRangeValue={setYRangeValue}
                setApsAbsValue={setApsAbsValue}
                setApsRelValue={setApsRelValue}
                compOutRangeValue={compOutRangeValue}
                setCompOutRangeValue={setCompOutRangeValue}
                baRangeValue={baRangeValue}
                setBaRangeValue={setBaRangeValue}
                scRangeValue={scRangeValue}
                setScRangeValue={setScRangeValue}
                cdRangeValue={cdRangeValue}
                setCdRangeValue={setCdRangeValue}
                errorRatio={errorRatio}
                setErrorRatio={setErrorRatio}
                regressionType={regressionType}
                setRegressionType={setRegressionType}
                useCalcErrorRatio={useCalcErrorRatio}
                setUseCalcErrorRatio={setUseCalcErrorRatio}
                labelOutput={labelOutput}
                setLabelOutput={setLabelOutput}
                differencePlotType={differencePlotType}
                setDifferencePlotType={setDifferencePlotType}
                concordanceOutputRange={concordanceOutputRange}
                setConcordanceOutputRange={setConcordanceOutputRange}
                ciMethod={ciMethod}
                setCiMethod={setCiMethod}
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
                setYThreshold3={setYThreshold3}
                notify={props.notify}
                uitext={props.uitext}
              />
            </React.Suspense>
          </Card>
        </AccordionPanel>
      </AccordionItem>
      <AccordionItem value="4">
        <AccordionHeader>{props.uitext["h_precision"]}</AccordionHeader>
        <AccordionPanel>
          <Card>
            <React.Suspense fallback={<div>Loading...</div>}>
              <Precision
                notify={props.notify}
                factorARangeValue={factorARangeValue}
                factorBRangeValue={factorBRangeValue}
                resultsRangeValue={resultsRangeValue}
                precOutRangeValue={precOutRangeValue}
                setFactorARangeValue={setFactorARangeValue}
                setFactorBRangeValue={setFactorBRangeValue}
                setResultsRangeValue={setResultsRangeValue}
                setPrecOutRangeValue={setPrecOutRangeValue}
                uitext={props.uitext}
              />
            </React.Suspense>
          </Card>
        </AccordionPanel>
      </AccordionItem>
      <AccordionItem value="5">
        <AccordionHeader>{props.uitext["h_precision_layout"]}</AccordionHeader>
        <AccordionPanel>
          <Card>
            <React.Suspense fallback={<div>Loading...</div>}>
              <PrecisionLayout
                notify={props.notify}
                factorARangeValue={factorARangeValue}
                setFactorARangeValue={setFactorARangeValue}
                factorBRangeValue={factorBRangeValue}
                setFactorBRangeValue={setFactorBRangeValue}
                resultsRangeValue={resultsRangeValue}
                setResultsRangeValue={setResultsRangeValue}
                uitext={props.uitext}
              />
            </React.Suspense>
          </Card>
        </AccordionPanel>
      </AccordionItem>
      <AccordionItem value="6">
        <AccordionHeader>{props.uitext["h_qual_comp"]}</AccordionHeader>
        <AccordionPanel>
          <Card>
            <React.Suspense fallback={<div>Loading...</div>}>
              <Qualitative
                notify={props.notify}
                qualXRangeValue={qualXRangeValue}
                qualYRangeValue={qualYRangeValue}
                qualOutputRangeValue={qualOutputRangeValue}
                setQualXRangeValue={setQualXRangeValue}
                setQualYRangeValue={setQualYRangeValue}
                setQualOutputRangeValue={setQualOutputRangeValue}
                uitext={props.uitext}
              />
            </React.Suspense>
          </Card>
        </AccordionPanel>
      </AccordionItem>
      <AccordionItem value="7">
        <AccordionHeader>{props.uitext["h_load_range_defaults"]}</AccordionHeader>
        <AccordionPanel>
          <Card>
            <div className={styles.container}>
              <div className={styles.field}>
                <p>{props.uitext["inf_load_range_defaults"]}</p>
                <Button appearance="primary" onClick={setDefaultValues}>
                  {props.uitext["btn_load_defaults"]}
                </Button>
                <p>{props.uitext["inf_default_template"]}</p>
                <Checkbox
                  label={props.uitext["opt_import_ws"]}
                  checked={importLayoutChecked}
                  onChange={(_ev, data) => setImportLayoutChecked(data.checked)}
                />
                <Field label={props.uitext["lbl_layout_sheet"]} className={styles.field}>
                  <Tooltip relationship="description" content={props.uitext["tip_layout_sheet"]}>
                    <Input
                      value={layoutSheet}
                      onChange={(ev) => setLayoutSheet(ev.target.value)}
                      type="text"
                    />
                  </Tooltip>
                </Field>
                <div className={styles.field}>
                  <table>
                    <thead>
                      <tr>
                        <th>{props.uitext["lbl_attribute"]}</th>
                        <th>{props.uitext["lbl_value"]}</th>
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
  );
} //ComparisonPane
