Office.onReady(function (info) {
    // Office is ready
    if (info.host === Office.HostType.Excel) {
        // Assign event handlers and functions here
        if (info.platform === Office.PlatformType.OfficeOnline) {
            // Office Online-specific code
        } else {
            // Desktop-specific code
        }
    }
});

// Function to add a new record to the Excel sheet
function addRecord() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var values = [
            [
                document.getElementById("department").value,
                document.getElementById("activities").value,
                document.getElementById("target").value,
                document.getElementById("achievement").value,
                document.getElementById("deadline").value,
                document.getElementById("submitted-on").value,
                document.getElementById("remarks").value,
            ]
        ];

        // Dynamically determine the range based on the existing data in column A
        var rowCount = sheet.getRange("A:A").getUsedRangeOrNullObject().load("rowCount");

        return context.sync().then(function () {
            if (rowCount.rowCount > 0) {
                // If there are existing records, find the next available row
                var newRange = sheet.getRangeByIndexes(rowCount.rowCount + 1, 0, 1, 7);
                newRange.values = values;
            } else {
                // If the column is empty, start at A2
                var newRange = sheet.getRange("A2:H2");
                newRange.values = values;
            }
            return context.sync();
        }).catch(function (error) {
            console.log("Error adding record: " + error);
        });
    }).catch(function (error) {
        console.log("Error adding record: " + error);
    });
}

// Function to search for a department and display details
function searchRecords() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var searchValue = document.getElementById("department").value;
        var range = sheet.getRange("A2:A100"); // Dynamically determine the range based on the existing data
        range.load("values");
        return context.sync()
            .then(function () {
                for (var i = 0; i < range.values[0].length; i++) {
                    if (range.values[0][i] === searchValue) {
                        // Display details or do something with the found data
                        // Implement code to display or process the found data
                        break;
                    }
                }
                console.log("Department not found. If you wish to add the department/Division/Unit, click on 'Add' option below.");
            }).catch(function (error) {
                console.log("Error searching records: " + error);
            });
    }).catch(function (error) {
        console.log("Error searching records: " + error);
    });
}

// Function to delete a record or row with a confirmation prompt for options
function deleteRecord() {
    // Use Office UI dialog API for confirmation
    Office.context.ui.displayDialogAsync(window.location.origin + '/confirmDialog.html', { height: 50, width: 50 });
}

// Function to update a record with a confirmation prompt
function updateRecord() {
    if (confirm("Do you wish to update this record?")) {
        var searchValue = document.getElementById("department").value; // The value to search for
        var updatedActivities = document.getElementById("activities").value;
        var updatedTarget = document.getElementById("target").value;
        var updatedAchievement = document.getElementById("achievement").value;
        var updatedDeadline = document.getElementById("deadline").value;
        var updatedSubmittedOn = document.getElementById("submitted-on").value;
        var updatedRemarks = document.getElementById("remarks").value;

        Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            var range = sheet.getRange("A2:A100"); // Dynamically determine the range based on the existing data
            range.load("values");
            return context.sync()
                .then(function () {
                    var rowIndex = -1;
                    for (var i = 0; i < range.values[0].length; i++) {
                        if (range.values[0][i] === searchValue) {
                            rowIndex = i;
                            break;
                        }
                    }
                    if (rowIndex >= 0) {
                        // Update the record in the Excel sheet
                        var updateRange = sheet.getRangeByIndexes(rowIndex, 1, 1, 6);
                        updateRange.values = [[
                            updatedActivities,
                            updatedTarget,
                            updatedAchievement,
                            updatedDeadline,
                            updatedSubmittedOn,
                            updatedRemarks
                        ]];
                        return context.sync();
                    } else {
                        console.log("Department not found. Unable to update the record.");
                    }
                });
        }).catch(function (error) {
            console.log("Error updating record: " + error);
        });
    }
}

// Function to exit the add-in form
function exitForm() {
    // Implement your task pane closing logic here
    // For example, you might want to close the task pane using Office UI API
    Office.context.ui.closeContainer();
}
