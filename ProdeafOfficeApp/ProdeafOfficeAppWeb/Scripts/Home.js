/// <reference path="../App.js" />

(function () {
    "use strict";

    var lastSelection = '';

    //The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#startWlAuto').click();
            setInterval(checkSelection, 500);
        });
    };

    function checkSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    if (result.value.replace(/\s/g,"")!="" && result.value != lastSelection) {
                        lastSelection = result.value;
                        translate(lastSelection);
                    }
                }
            }
        );
    }

    function translate(text) {
        wl.OpenText(text);
    }

})();