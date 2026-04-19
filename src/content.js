/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "./content.css";

(function () {
    "use strict";
    // The onReady function must be run each time a new page is loaded.
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            // Initialize immediately since Office.onReady already waits for DOM.
        }
    });
})();
