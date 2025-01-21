#target indesign //Tells Adobe ExtendScript to target Adobe InDesign specifically when running the script.
//Version 0.5
var doc = app.activeDocument; //Gets currently opened document.
var pages = doc.pages; //Makes a reference to the "pages" of the active document.
var masterSpreads = doc.masterSpreads; //List of all master pages (or spreads) in document.
var masterIndex = 0; //Index of the parent page, currently selected in dropdown menu.
var numberOfPages = 0; //The number of pages to be changed, entered manually by the user on InputField.
var storedLastPageUpdated = 0; //Number of the last page updated.
var queuedPageChanges = []; //2D array to store master index and number of pages.
var openedDialogs = []; //Array of all currently opened dialogue windows.
var iterationCounter = 1; //Current runtime iteration of the script flow (starts with 1 for better readability).

// Applies the master page to the specified number of pages.
function applyMasterPages()
{
    for (var i = 0; i < queuedPageChanges.length; i++)
    {
        var masterIndex = queuedPageChanges[i][0];
        var numberOfPages = queuedPageChanges[i][1];
        var maxPages = pages.length;

        if (storedLastPageUpdated + numberOfPages > maxPages)
        {
            numberOfPages = maxPages - storedLastPageUpdated; // Adjust if it exceeds the total number of pages
        }

        var targetPages = pages.everyItem().getElements().slice(storedLastPageUpdated, storedLastPageUpdated + numberOfPages);

        for (var j = 0; j < targetPages.length; j++)
        {
           targetPages[j].appliedMaster = masterSpreads.item(masterIndex);
        }
        storedLastPageUpdated += numberOfPages;
    }
}
// Creates a dialog window with input fields and buttons.
function createDialog()
{
    //Creates a new dialog window panel
    var dialog = new Window('dialog', 'Stage ' + iterationCounter);

    openedDialogs.push(dialog);

    //Gets a array of master pages
    var masterNames = [];

    for (var i = 0; i < masterSpreads.length; i++)
    {
        masterNames.push(masterSpreads[i].name);
    }
    //Creates a dropdown list of master pages
    var masterDropdown = dialog.add('dropdownlist', undefined, masterNames);

    if (iterationCounter < queuedPageChanges.length)
    {
        masterDropdown.selection = queuedPageChanges[iterationCounter - 1][0]; //Sets a default selection to the first master in masterNames array.
    }
    else
    {
        masterDropdown.selection = 0;
    }
    //Adds input field to enter number of pages, that user wants to update with new master page.
    dialog.add('statictext', undefined, 'Number of Pages:');

    var numberOfPagesInput = dialog.add('edittext', undefined, 0);
    numberOfPagesInput.characters = 5;
    //Creates a "Apply", "Revert changes" and "Close" buttons.
    var nextButton = dialog.add('button', undefined, 'Next');
    var applyButton = dialog.add('button', undefined, 'Aplly');
    var revertButton = dialog.add('button', undefined, 'Revert');

    nextButton.onClick = function ()
    {
        masterIndex = masterDropdown.selection.index;
        numberOfPages = parseInt(numberOfPagesInput.text, 10);

        queuedPageChanges.push([masterIndex, numberOfPages]);

        dialog.close();

        iterationCounter++;

        createDialog();
    }
    
    applyButton.onClick = function ()
    {
        masterIndex = masterDropdown.selection.index;

        numberOfPages = parseInt(numberOfPagesInput.text, 10);

        queuedPageChanges.push([masterIndex, numberOfPages]);

        dialog.close();
    }

    revertButton.onClick = function ()
    {
        dialog.close();

        iterationCounter--;

        createDialog();

        if (queuedPageChanges.length >= 1)
        {
            queuedPageChanges.splice(-1, -1);
        }
        else
        {
            queuedPageChanges.length = 0; // Clear the array if it has less than 2 elements
        }

        if (queuedPageChanges.length >= 1)
        {
            queuedPageChanges.splice(-1);
        }
    }

    dialog.onClose = function ()
    {
        //"Whenever dialogue window closes"" code here
    }

    dialog.show();
}

function closeAllDialogs()
{
    for (var i = 0; i < openedDialogs.length; i++)
    {
        openedDialogs[i].close();
    }
}

function setup()
{
    createDialog();

    applyMasterPages();
}

setup();
