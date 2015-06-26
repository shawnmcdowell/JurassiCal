/// <reference path="../App.js" />

(function () {
    "use strict";

    // Global constant: binding ID
    window.bindingID = 'myBinding';
    var binding = null;

    //CALL TO POST REQUEST WITH HARDCODED DATA
    var sampleHeaders = [['Subject', 'Start', 'End', 'Location', 'Body', 'Attendees']];
    var sampleRows = [
    ['Soccer Game', '6/26/2015 8:00 AM', '6/26/2015 9:00 AM', 'Relay Park', 'First game of the season', 'shawnmc@microsoft.com;shawnmc@outlook.com'],
    ['Jess\'s birthday party', '6/26/2015 9:00 AM', '6/26/2015 10:00 AM', '100 Main Street', 'Jess is celebrating her 3rd birthday!', 'shawnmc@microsoft.com;shawnmc@outlook.com'],
    ['Date Night', '6/26/2015 7:00 PM', '6/26/2015 10:00 PM', 'Issaquah Regal Cinemas/The Ram', 'Jurassic World', 'shawnmc@microsoft.com;shawnmc@outlook.com']];

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#set-sample-data').click(insertSampleData);
            $('#get-data-from-selection').click(addFromSelection);
            $('#create-calendar-events').click(processBindingData);
        });
    };

    // Reads data from current table selection and displays a notification
    function getDataFromSelection() {
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

    function addFromSelection() {
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
                    binding.getDataAsync({valueFormat: Office.ValueFormat.Formatted}, function(result){
                        var header = result.value.headers[0];
                        var rows = result.value.rows;
                        rows.forEach(function (row, rowIndex) {
                            var jsonEvent = "{";
                            header.forEach(function (col, colIndex) {
                                console.log(col, row[colIndex]);
                                switch (col) {
                                    case "Attendees":
                                        var attendees = row[colIndex].split(";");
                                        jsonEvent += '"Attendees": [';
                                        attendees.forEach(function (attendee, attendeeIndex) {
                                            jsonEvent += '{"Address": "' + attendee + '"}'
                                            jsonEvent += attendeeIndex < attendees.length-1 ? ',' : ""
                                        });
                                        jsonEvent += ']';
                                        break;
                                    case "End":
                                    case "Start":
                                        jsonEvent += '"' + col + '": "' + new Date(row[colIndex]) + '"';
                                        break;
                                    default:
                                        //{ "Subject": "Soccer Game", "Start": "6/26/2015 10:00 AM", "End": "6/26/2015 12:00 PM", "Location": "Soccer field #3", "Attendees": [{ Address: "shawnmc@microsoft.com" }, { Address: "shawnmc@outlook.com" }], "Body": "<html><p>Today we are playing the ManU.  We are vistors.</p></html>" };
                                        jsonEvent += '"' + col + '": "'+ row[colIndex] +'"';
                                }
                                jsonEvent +=  colIndex < header.length-1  ? ',' : ""
                            });
                            jsonEvent += "}";

                            //Call the service to create the event
                            $.post("http://localhost:8010/createevent", jsonEvent, function (result, status) {
                                if (status === "success") {
                                    console.log(result);
                                    app.showNotification("Event Creation Succeeded for " + JSON.parse(jsonEvent).Subject);
                                } else {
                                    console.log("something went wrong!");
                                }
                            }, "json");
                        });
                    });
                } else {
                    app.showNotification('No binding exists');
                }
            });
    }

    function makeSamplePostRequests(data) {
        var tempData = { "Subject": "Soccer Game", "Start": "6/26/2015 10:00 AM", "End": "6/26/2015 12:00 PM", "Location": "Soccer field #3", "Attendees": [{ Address: "shawnmc@microsoft.com" }, { Address: "shawnmc@outlook.com" }], "Body": "<html><p>Today we are playing the ManU.  We are vistors.</p></html>" };
        $.post("http://localhost:8010/createevent", JSON.stringify(tempData), function (result, status) {
            if (status === "success") {
                console.log(result);
            } else {
                console.log("something went wrong!");
            }
        }, "json");
    }
})();