import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import {UIStrings} from "../uistrings";

/* global document, Office, module, require, HTMLElement */

const title = "Method Evaluation Tools Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  // Get the language setting for editing document content.
  // To test this, uncomment the following line and then comment out the
  // line that uses Office.context.displayLanguage.
  // const myLanguage = Office.context.contentLanguage;

  // Get the language setting for UI display in the Office application.
  const myLanguage = Office.context.displayLanguage;
  let UIText: { [key: string]: string };

  // Get the resource strings that match the language.
  // Use the UIStrings object from the UIStrings.js file
  // to get the JSON object with the correct localized strings.
  UIText = UIStrings.getLocaleStrings(myLanguage);
  
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} uitext={UIText} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
