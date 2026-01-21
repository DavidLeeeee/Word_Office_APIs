/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import React from "react";
import { createRoot } from "react-dom/client";
import App from "../React/App";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    const container = document.getElementById("react-app");
    if (container) {
      const root = createRoot(container);
      root.render(React.createElement(App));
    }
  }
});
