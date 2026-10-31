import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./App";
import { UIStrings } from "../uistrings";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  const uitext = UIStrings.getLocaleStrings(Office.context.displayLanguage);

  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App uitext={uitext} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./App", () => {
    const NextApp = require("./App").default;
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <NextApp uitext={UIStrings.getLocaleStrings(Office.context.displayLanguage)} />
      </FluentProvider>
    );
  });
}
