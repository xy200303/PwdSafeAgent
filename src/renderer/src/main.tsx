import React from "react";
import ReactDOM from "react-dom/client";
import { Provider } from "react-redux";
import "@incremark/theme/styles.css";
import "./styles.css";
import { App } from "./ui/App";
import { ErrorBoundary } from "./ui/ErrorBoundary";
import { store } from "./store";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <Provider store={store}>
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    </Provider>
  </React.StrictMode>
);
