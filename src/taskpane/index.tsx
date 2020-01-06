import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { createStore } from 'botframework-webchat';
import * as React from "react";
import * as ReactDOM from "react-dom";
import { Dispatcher } from "flux";
/* global AppCpntainer, Component, document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;
const title = "Contoso Task Pane Add-in";
const eventDispatcher = new Dispatcher();

const store = createStore({},
  () => next => action => {
  if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
    if (action.payload.activity.type == "event") {
      eventDispatcher.dispatch(action.payload.activity);
    }
  }
  return next(action);
});

const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} store={store} eventDispatcher={eventDispatcher}/>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
