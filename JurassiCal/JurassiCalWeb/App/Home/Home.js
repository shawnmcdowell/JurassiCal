/// <reference path="../App.js" />

(function () {
    "use strict";

    // Global constant: binding ID
    window.bindingID = 'myBinding';
    var binding = null;

    //CALL TO POST REQUEST WITH HARDCODED DATA
    var sampleHeaders = [['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'Location', 'Body', 'Attendees']];
    var sampleRows = [
    ['Soccer Game', '9-5-2015', '2:00 PM', '9-5-2015', '4:00 PM', 'Relay Park', 'First game of the season', ''],
    ['Jess\'s birthday party', '9-8-2015', '11:00 AM', '9-8-2015', '1:00 PM', '100 Main Street', 'Jess is celebrating her 3rd birthday!', ''],
    ['Date Night', '9/12/2015', '7:00 PM', '9/12/2015', '9:30 PM', 'Movie Theater', '', '']];

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#set-sample-data').click(insertSampleData);
            $('#get-data-from-selection').click(getDataFromSelection);
            $('#create-calendar-events').click(processBindingData);
        });
    };

    // Reads data from current table selection and displays a notification
    function getDataFromSelection() {
        makePostRequests(null);


        return;

        Office.context.document.getSelectedDataAsync(Office.CoercionType.Table,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    if (result.error.message === "The current selection is not compatible with the specified coercion type.") {
                        //app.showNotification('Error:', 'Please format the data as a table, then try again'); //Excel = "Format as Table", Word slightly different

                    } else {
                        app.showNotification('Error:', result.error.message);
                    }
                } else {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                    console.log("Guess this worked " + result.value);
                    //If successful, create binding to selected data table
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: window.bindingID },
                        function (asyncResult) {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                app.showNotification('Error binding data');
                            } else {
                                //Successful, do nothing
                            }
                        }
                    );
                }
            }
        );
    }

    function getDataFromSelection2() {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Table, { id: window.bindingID },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    //window.location.href = '../Home/Home.html';
                    //Successful, do nothing
                } else {
                    app.showNotification(result.error.name, result.error.message);
                }
            }
        );
    }

    //Creates TableData of sample data, writes it to selected cell in chart, and binds to it
    function insertSampleData() {
        var sampleData = new Office.TableData(
            sampleRows, sampleHeaders);
        //Insert sample data
        Office.context.document.setSelectedDataAsync(sampleData,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    app.showNotification('Could not insert sample data', 'Please choose a different selection range.');
                } else {
                    //If successful, create binding to sample data table
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: window.bindingID },
                        function (asyncResult) {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                app.showNotification('Error binding data');
                            } else {
                                //Successful, do nothing
                            }
                        }
                    );
                }
            }
        );
    }

    //Get binding by window.bindingID and call appropriate function
    function processBindingData() {
        Office.context.document.bindings.getByIdAsync(
            window.bindingID,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    binding = result.value;
                    console.log("process binding data", binding);
                } else {
                    app.showNotification('No binding exists');
                }
            });
    }

    function makePostRequests(data) {
        var tempData = {Subject: 'Soccer Game', Start : '6/26/2015 2:00 PM', End : '6/26/2015 2:00 PM', Location: 'Soccer field #3', Attendees: ['shawnmc@microsoft.com', 'shawnmc@outlook.com']};
        $.post("http://localhost:8010/postrequest", tempData, function (result, status) {
            if (status === "success") {
                console.log(result);
            } else {
                console.log("something went wrong!");
            }
        })
    }
})();