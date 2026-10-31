import * as React from "react";
import {
  makeStyles,
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Card,
  Link,
  Toast,
  ToastBody,
  ToastTitle,
  ToastTrigger,
  Toaster,
  useId,
  useToastController,
} from "@fluentui/react-components";
import BoxCox from "../taskpane/components/BoxCox";
import RefInt from "../taskpane/components/RefInt";
import RefLim from "../taskpane/components/RefLim";

interface AppProps {
  uitext: { [key: string]: string };
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = ({ uitext }) => {
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const styles = useStyles();

  const notify = (intent: string, message: string) => {
    const title = intent === "error" ? "Error" : intent === "warning" ? "Warning" : "Info";
    dispatchToast(
      <Toast appearance="inverted">
        <ToastTitle
          action={
            <ToastTrigger>
              <Link>Dismiss</Link>
            </ToastTrigger>
          }
        >
          {title}
        </ToastTitle>
        <ToastBody>{message}</ToastBody>
      </Toast>,
      { intent: intent === "error" || intent === "warning" ? intent : "info" }
    );
  };

  return (
    <div className={styles.root}>
      <Toaster toasterId={toasterId} position="top" timeout={5000} pauseOnHover />
      <Accordion defaultOpenItems="1">
        <AccordionItem value="1">
          <AccordionHeader>{uitext["h_ref_int"]}</AccordionHeader>
          <AccordionPanel>
            <Card>
              <RefInt notify={notify} uitext={uitext} />
            </Card>
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value="2">
          <AccordionHeader>{uitext["h_box_cox"]}</AccordionHeader>
          <AccordionPanel>
            <Card>
              <BoxCox notify={notify} uitext={uitext} />
            </Card>
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value="3">
          <AccordionHeader>{uitext["h_ri_reflim"]}</AccordionHeader>
          <AccordionPanel>
            <Card>
              <RefLim notify={notify} uitext={uitext} />
            </Card>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </div>
  );
};

export default App;
