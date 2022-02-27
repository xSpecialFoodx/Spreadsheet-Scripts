// user config start

// whether to add custom menus to the spreadsheet or not
var AddSpreadsheetMenus = true;

// user config end

// adding the custom menus to the spreadsheet
function AddSpreadsheetMenus_()
{
  AddSpreadsheetToolsMenu();
  AddFixSpreadsheetFilesSizesMenu();
  AddAnnounceSpreadsheetToDisMenu();
}

// a function that gets triggered when the spreadsheet opens
function onOpen()
{
  if (
    AddSpreadsheetMenus != undefined
    && AddSpreadsheetMenus != null
  )
  {
    InitializeGeneralToolsObject_();

    if (
      GeneralToolsObject.VariableIsBoolean(AddSpreadsheetMenus) == true
      && AddSpreadsheetMenus == true
    )
      AddSpreadsheetMenus_();
  }
}
