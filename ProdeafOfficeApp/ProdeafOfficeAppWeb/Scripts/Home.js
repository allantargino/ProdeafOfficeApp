/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#translate-button').click(translate);
        });
    };

    // Reads data from current document selection and displays a notification
    function translate() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var text = result.value;
                    //var box = $('#traduzir');
                    //box.text(text);
                    //box.hide();
                    //box[0].click();

                    wl.OpenText(text);
                }
            }
        );
    }
})();