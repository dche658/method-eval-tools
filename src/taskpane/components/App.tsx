/**
* Main component for the add-in. 
* 
* @author Douglas Chesher
*
*/
import * as React from "react";
import BoxCox from "./BoxCox";
import RefInt from "./RefInt";
import ComparisonPane from "./ComparisonPane";

import {
  makeStyles, useId,
  Accordion, AccordionItem, AccordionHeader, AccordionPanel,
  Link,
  Card,
  Tab, TabList,
  Toaster,
  useToastController,
  ToastTitle,
  Toast,
  ToastTrigger,
  ToastBody,
} from "@fluentui/react-components";

import type {
  SelectTabData,
  SelectTabEvent,
  TabValue,
} from "@fluentui/react-components";

interface AppProps {
  title: string;
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
  }
});



/* Main application entry
*/
const App: React.FC<AppProps> = (props: AppProps) => {

  const [selectedTabValue, setSelectedTabValue] = React.useState<TabValue>("evaluation");

  const toasterId = useId("toaster");

  const { dispatchToast } = useToastController(toasterId);

  const styles = useStyles();

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    if (event.isTrusted) {
      setSelectedTabValue(data.value);
    }
  };

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


  const Evaluation = React.memo(() => {
    return (
      <div role="tabpanel" aria-labelledby="Evaluation">
        <ComparisonPane notify={notify} uitext={props.uitext} />
      </div>
    );
  });

  const Utilities = React.memo(() => {
    return (
      <div role="tabpanel" aria-labelledby="Utilities">
        <Accordion defaultOpenItems="1">
          <AccordionItem value="1">
            <AccordionHeader>{props.uitext["h_ref_int"]}</AccordionHeader>
            <AccordionPanel>
              <Card>
                <RefInt notify={notify} uitext={props.uitext} />

              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="2">
            <AccordionHeader>{props.uitext["h_box_cox"]}</AccordionHeader>
            <AccordionPanel>
              <Card>
                <BoxCox notify={notify} uitext={props.uitext} />
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

        <Accordion defaultOpenItems="1">
          <AccordionItem value="1">
            <AccordionHeader>{props.uitext["h_about"]}</AccordionHeader>
            <AccordionPanel>
              <Card>
                <p>
                  {props.uitext["inf_about"]}
                </p>
                <p>
                  <a href="https://metools.chesher.id.au/help/index.html" target="help">
                    <span className={styles.label}>{props.uitext["lbl_ifu"]}</span>
                  </a>.
                </p>
              </Card>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="2">
            <AccordionHeader>{props.uitext["h_src"]}</AccordionHeader>
            <AccordionPanel>
              <Card>
                <p>
                  {props.uitext["inf_src"]} <a
                    href="https://github.com/dche658/method-eval-tools">
                    https://github.com/dche658/method-eval-tools </a
                  >.
                </p>
                <p>
                  {props.uitext["inf_disclaimer"]}
                </p>
              </Card>
            </AccordionPanel>
          </AccordionItem>
        </Accordion>
      </div>
    );
  });

  return (
    <div className={styles.root}>
      <Toaster toasterId={toasterId} position="top" timeout={5000} pauseOnHover />
      <TabList selectedValue={selectedTabValue} onTabSelect={onTabSelect}>
        <Tab id="Evaluation" value="evaluation">
          {props.uitext["tab_evaluation"]}
        </Tab>
        <Tab id="Utilities" value="utilities">
          {props.uitext["tab_utilities"]}
        </Tab>
        <Tab id="Help" value="help">
          {props.uitext["tab_help"]}
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
