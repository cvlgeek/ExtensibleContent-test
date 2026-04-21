/*
 * Copyright (c) PK-421. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "./content.css";
import "./fluentui-web-components.min.js";
import "./fabric.css";
import "./prism.js";
import "./prism.css";

/*
    document.querySelector('body').addEventListener('contextmenu', (event) => {
        event.preventDefault(); // Prevents the default browser context menu from appearing
        xc.toggleView();
    });
*/


(function () {
    "use strict";
    // The onReady function must be run each time a new page is loaded.
    Office.onReady((info) => {
        //if (info.host === Office.HostType.Excel) {
            // Initialize immediately since Office.onReady already waits for DOM.
            const xc = {}
                xc.toggleView = function(){
                    document.querySelectorAll('.xc-view').forEach(function(el){el.classList.toggle('hidden')});
                };
                xc.defaultContent = `<div id="xc-display-no-content" style="height:100vh; width:100vw; display:grid; place-items: center;">\n\t<div style="text-align:center;">\n\t\tNo content to display.\n\t\t<br>\n\t\t<fluent-button appearance="accent" onclick="xc.toggleView();">Configure</fluent-button>\n\t</div>\n</div>`;
                xc.saveConfiguration = function(){
                    document.querySelectorAll('.xc-view').forEach(function(el){el.classList.toggle('hidden')});
                };
                xc.editor = null;
                xc.init = function(){
                    xc.editor = createEditor('#xc-configuration-content-panel', {
                        language: 'js',
                        value: xc.defaultContent,
                    });
                };
            xc.init();
        //}
    });
})();
