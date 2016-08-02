/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#startWlAuto')[0].click();
            $('#translate-button').click(translate);
        });
    };

    // Reads data from current document selection and displays a notification
    function translate() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var text = result.value;
                    wl.OpenText(text);
                }
            }
        );
    }

})();