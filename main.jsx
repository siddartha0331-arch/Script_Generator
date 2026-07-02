import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

// This is the one line of "glue" that connects the plain index.html above
// to your actual React component tree. createRoot + render is the standard
// React 18 way to boot an app — every Vite/CRA React project has some
// version of this file.
ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
