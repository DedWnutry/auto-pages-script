#target indesign

// MAIN FUNCTIONALITY

// Make a reference to the "pages" of the active document.
var pages = app.activeDocument.pages;

// Make a reference to the "masterPages" of the active document.
var masterPages = app.activeDocument.masterSpreads;

// Variable to track the last applied page index
var startPageIndex = 0;

// Applies a master to a range of pages starting from startPageIndex
// Takes 3 arguments:
//   masterName   the name of the master page
//   numberOfPages the number of pages to apply the master to
//   startPageIndex the index from which to start applying the master page
function applyMasterToPages(masterName, numberOfPages, startPageIndex) {
    // Make a reference to the "masterPage" to apply.
    var masterPage = masterPages.item(masterName);

    // Ensure that numberOfPages doesn't exceed the total number of pages
    var maxPages = pages.length;
    if (startPageIndex + numberOfPages > maxPages) {
        numberOfPages = maxPages - startPageIndex; // Adjust if it exceeds the total number of pages
    }

    // Iterate through the "pages" from startPageIndex to startPageIndex + numberOfPages
    for (var i = startPageIndex; i < startPageIndex + numberOfPages; i++) {
        // Make a reference to the "page".
        var page = pages.item(i);

        // Apply the "masterPage" to the "page" by setting its "appliedMaster" attribute.
        page.appliedMaster = masterPage;
    }

    // Return the new startPageIndex for the next call to keep track of where to start
    return startPageIndex + numberOfPages;
}

// DIALOG

// Stores all the names of the document's master pages in an array.
function getMasterPageNames() {
    var masterPageNames = [];
    for (var i = 0; i < masterPages.length; i++) {
        masterPageNames.push(masterPages.item(i).name);
    }
    return masterPageNames;
}

// Displays the input dialog and handles user interaction.
function displayDialog() {
    var dialog = app.dialogs.add({ name: "Add Master To Pages" });
    var masterPageNames = getMasterPageNames();

    with (dialog) {
        with (dialogColumns.add()) {
            with (borderPanels.add()) {
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "Master:" });
                    staticTexts.add({ staticLabel: "Number of Pages (N):" });
                }
                with (dialogColumns.add()) {
                    var masterNameDropdown = dropdowns.add({ stringList: masterPageNames, selectedIndex: 0, minWidth: 200 });
                    var numberOfPagesField = integerEditboxes.add({ editValue: 1, minimumValue: 1, maximumValue: pages.length });
                }
            }
        }

        // Show the dialog and handle "Apply" and "Cancel" buttons
        var dialogReturn = dialog.show();

        if (dialogReturn == true) {
            // When "Apply" is clicked, apply the master page to the specified range
            var masterName = masterPageNames[masterNameDropdown.selectedIndex];
            var numberOfPages = numberOfPagesField.editValue;

            // Apply the master page to the first N pages starting from the last applied page index
            startPageIndex = applyMasterToPages(masterName, numberOfPages, startPageIndex);

            // Keep the dialog open for further use
            displayDialog();
        } else {
            // If "Cancel" is clicked, destroy the dialog and stop the process
            dialog.destroy();
        }
    }
}

displayDialog();
