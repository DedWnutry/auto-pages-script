#target indesign
//Version 1.0

var doc = app.activeDocument; //Gets currently opened document.
var pages = doc.pages; //Makes a reference to the "pages" of the active document.
var masterSpreads = doc.masterSpreads; //List of all master pages (or spreads) in document.
var lastPageUpdated = 0; //Number of the last page updated.
var queuedChangesValues = []; //2D array to store master index and number of pages.
var prevIterationPageMasters = [];
var prevPagesUpdated = [];
var openedDialogs = []; //Array of all currently opened dialogue windows.
var iterationCounter = 0; //Current runtime iteration of the script flow (starts with 1 for better readability).
//Initiate a first iteration of the script and launches dialogue window for User to interact.
createDialog(0, 0);
// Create a main dialog window with input fields and buttons.Can be called multiple times during a single runtime.
function createDialog(masterIndex, numberOfPages) {
    //Creates a new dialog window panel.
    var dialog = new Window('dialog', 'Stage ' + (iterationCounter + 1));
    //Add a currently opened window to the array, that will track all previously opened dialogs.
    openedDialogs.push(dialog);
    //Initialize an array of master pages.
    var masterNames = [];
    //Add all master pages names into dropdown list (with their indices).
    for (var i = 0; i < masterSpreads.length; i++) masterNames.push(masterSpreads[i].name);
    //Creates a dropdown list of master pages and nates data from masterNames.
    var masterDropdown = dialog.add('dropdownlist', undefined, masterNames);
    //Set the default value for dropdown in the case User reverts changes and returns to the last dialogue window based of current iteration, or just assigns a zero if this is a first or current iteration.
    masterDropdown.selection = masterIndex;
    //Add input field to enter number of pages, that User wants to update with new master page.
    dialog.add('statictext', undefined, 'Number of Pages:');
    var pagesNumInputField = dialog.add('edittext', undefined, numberOfPages);
    pagesNumInputField.characters = 5;//Set the visible width of the input field in terms of the number of characters that can be displayed.
    //Add a new panel for handling buttons.
    var buttonGroup = dialog.add('group');
    buttonGroup.orientation = 'row';// Set horizontal layout for gruop.
    //Creates a "Next" button.
    var nextButton = buttonGroup.add('button', undefined, 'Next', { name: 'ok' });
    //Creates a "Revert" button, but only if it's a second or further iteration.
    if (iterationCounter > 0) {
        var revertButton = buttonGroup.add('button', undefined, 'Revert', { name: 'cancel' });

        revertButton.onClick = function () {
            cancelClicked = true;
            dialog.close(2);
        }
    }
    //Creates a "Quit" button.
    var quitButton = buttonGroup.add('button', undefined, 'Quit');
    //Set flag when cancel button is clicked.
    var cancelClicked = false;

    quitButton.onClick = function () {
        cancelClicked = true;
        dialog.close(0);
    }
    nextButton.onClick = function () {
        cancelClicked = true;
        dialog.close(1);
    }
    // Override onClose event to detect "X" button
    dialog.onClose = function () {
        if (!cancelClicked) dialog.close(3); // X button clicked
    }
    //Now uses dialog.show() return value instead of button.onClick events(because, as it turned out, they blocked the script from running at runtime).
    var result = dialog.show();

    if (result == 1) //Corresponds when User presses a "Next" button: records current values, inserted by User into Dropdown list and InputField, listens them to queuedChangesValues 2D array, closes the dialogue window, updates iteration count(for display on next dialog window), applies master pages in loop and creates a new dialogue.
    {
        //Get mater page preset by index from dropdown.
        masterIndex = masterDropdown.selection.index;
        //Conver to int number of pages from input field.
        numberOfPages = parseInt(pagesNumInputField.text, 10);
        //Remember these value to return to them later.
        queuedChangesValues.push([masterIndex, numberOfPages]);
        //Increase iteration count.
        iterationCounter++;
        //Applies new master styles.
        applyMasterPages(masterIndex, numberOfPages);
        //Create a new dialog instantly.
        createDialog(0, 0);
    }
    else if (result == 2) //If User pressed "Rever" button.
    {
        //Lower iteration count.
        iterationCounter--;
        revertChanges();
        //Remove last entry from queuedChangesValues.
        if (queuedChangesValues.length > 0) var lastEntry = queuedChangesValues.pop();
        //Remember index of last entry.
        var index = lastEntry[0];
        //Remember number of pages of last entry.
        var number = lastEntry[1];
        //Create a new dialog (to simulate "returning" to the last "hidden" dialogue window) with prescripted values in the dropdown and inputfield.
        createDialog(index, number);
    }
}
//Apply new master page by index (masterIndex) to the selected number of pages (numberOfPages).
function applyMasterPages(masterIndex, numberOfPages) {
    //Get a total pages amount.
    var maxPages = pages.length;
    //If estimated quantity of pages for update exceeds maximum capacity of pages in document, reduce it to appropriate number.
    if (lastPageUpdated + numberOfPages > maxPages) numberOfPages = maxPages - lastPageUpdated;
    //Select from all pages list of document only required ones, which will be updated.
    var targetPages = pages.everyItem().getElements().slice(lastPageUpdated, lastPageUpdated + numberOfPages);
    //Remember former parent page and enters it into the register.
    for (var i = 0; i < targetPages.length; i++) prevIterationPageMasters.push(targetPages[i].appliedMaster);
    //Update registry entry of pages amount affected on the last iteration.
    prevPagesUpdated.push(numberOfPages);
    //Consistently applies master spread by index to each page from targetPages (forEach loops don't work for InDesign).For some reason also don't work when function called from onClick event of a button.
    for (var i = 0; i < targetPages.length; i++) targetPages[i].appliedMaster = masterSpreads.item(masterIndex);
    //Update a index of last page affected by the script.
    lastPageUpdated += numberOfPages;
}
//Revert changes of document to the last state (previous iteration of script).
function revertChanges() {
    //Get array of pages, changed on the during iteration.
    var targetPages = pages.everyItem().getElements().slice(lastPageUpdated - prevPagesUpdated[prevPagesUpdated.length - 1], lastPageUpdated);
    //Loops through targetPages and returs their previous master spread.
    for (var i = 0; i < targetPages.length; i++) targetPages[i].appliedMaster = prevIterationPageMasters[i];
    //Updates last affected page number.
    lastPageUpdated -= targetPages.length;
    //Remove last registry entry of pages amount affected on the last iteration.
    prevPagesUpdated.pop();
}
