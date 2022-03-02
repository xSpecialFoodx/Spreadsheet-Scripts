// user config start

// the format of the date that will be received by the spreadsheet (according to the spreadsheet timezone)
var SpreadsheetDateFormat = "MM-dd-yy";

// the smaller this number the safer, but also the more slower it'd become if there's nothing to be replaced
// set as 0 or under for infinite
var MaxRowsPerCheck = 500;

// the minimum amount of numbers in a title id (for example, "779" as 5 length would be displayed as "00779")
var TitleIDMinNumbersAmount = 5;
// will be displayed before the title id number (for example, "00779" with the type "EXAMPLE" would be displayed as "EXAMPLE00779")
var TitleIDType = "EXAMPLE";

// this variable will be used for people to be able to write regions shortly
var KnownRegions = [];

KnownRegions.push({Shorts: ["A"], Full: "ASIA"});
KnownRegions.push({Shorts: ["J", "JP"], Full: "JPN"});
KnownRegions.push({Shorts: ["E", "EU"], Full: "EUR"});
KnownRegions.push({Shorts: ["U", "US"], Full: "USA"});

// used when the version of the entry is 1.00
var NoUpdateSizeText = "None";

// used when there are no fixes needed for the specific entry
var NoFixesText = "None";

// the format that's used for the version cell
var VersionFormat = "0.00";

// the format that's used for the date cell
var DateFormat =
  "mm"
  + '"' + '-' + '"'
  + "dd"
  + '"' + '-' + '"'
  + "yy";

// this variabble will be used for people to be able to have the submitter cell get filled automatically
// based on the user that launched the script (searching by the user's name)
// note:
// the user name can be checked by the ShowUserName function, which pops the user name in a message
// , it's also in the Spreadsheet Tools UI tab
var Submitters = [];

Submitters.push({RealUserName: "example_user", DisplayedUserName: "Example"});

// whether to allow users to add custom fixes to the fixes column
var AllowCustomFixes = true;

// user config end

var SpreadsheetToolsObject = null;// do not delete

// the repositories of the spreadsheet start

var 4RepoListName =  "4 Repo List";// do not delete
var 4AppsListName = "4 Apps List";// do not delete
var 2To4RepoListName = "2 To 4 Repo List";// do not delete
var 4RegionInfo = "4 Region Info";// do not delete
var GameAppInfo = "GameApp Info";// do not delete

// the repositories of the spreadsheet end

// the columns of the sheet names start

var TitleColumnName = "Title";// do not delete
var TitleIDColumnName = "Title ID";// do not delete
var RegionColumnName = "Region";// do not delete
var GenreColumnName = "Genre";// do not delete
var BaseSizeColumnName = "Base Size";// do not delete
var VersionColumnName = "Version";// do not delete
var UpdateSizeColumnName = "Update Size";// do not delete
var DLCColumnName = "DLC";// do not delete
var BPColumnName = "BP";// do not delete
var UploaderColumnName = "Uploader";// do not delete
var DateColumnName = "Date";// do not delete
var LinkColumnName = "Link";// do not delete
var TestedColumnName = "Tested";// do not delete
var InfoColumnName = "Info";// do not delete
var DKsMirrorColumnName = "DK's Mirror";// do not delete
var UpdateLinkColumnName = "Update Link";// do not delete
var SubmitterColumnName = "Submitter";// do not delete
var AnnouncedColumnName = "Announced";// do not delete
var FixesColumnName = "Fixes";// do not delete

// the columns of the sheet names end

// the columns of the sheet start

var TitleColumn = -1;// do not delete
var TitleIDColumn = -1;// do not delete
var RegionColumn = -1;// do not delete
var GenreColumn = -1;// do not delete
var BaseSizeColumn = -1;// do not delete
var VersionColumn = -1;// do not delete
var UpdateSizeColumn = -1;// do not delete
var DLCColumn = -1;// do not delete
var BPColumn = -1;// do not delete
var UploaderColumn = -1;// do not delete
var DateColumn = -1;// do not delete
var LinkColumn = -1;// do not delete
var TestedColumn = -1;// do not delete
var InfoColumn = -1;// do not delete
var DKsMirrorColumn = -1;// do not delete
var UpdateLinkColumn = -1;// do not delete
var SubmitterColumn = -1;// do not delete
var AnnouncedColumn = -1;// do not delete
var FixesColumn = -1;// do not delete

// the columns of the sheet end

// the values of the columns of the sheet start

var TitleValue = null;// do not delete
var TitleIDValue = null;// do not delete
var RegionValue = null;// do not delete
var GenreValue = null;// do not delete
var BaseSizeValue = null;// do not delete
var VersionValue = null;// do not delete
var UpdateSizeValue = null;// do not delete
var DLCValue = null;// do not delete
var BPValue = null;// do not delete
var UploaderValue = null;// do not delete
var DateValue = null;// do not delete
var LinkValue = null;// do not delete
var TestedValue = null;// do not delete
var InfoValue = null;// do not delete
var DKsMirrorValue = null;// do not delete
var UpdateLinkValue = null;// do not delete
var SubmitterValue = null;// do not delete
var AnnouncedValue = null;// do not delete
var FixesValue = null;// do not delete

// the values of the columns of the sheet end

// the trimmed values of the columns of the sheet start

var TitleValueTrimmed = null;// do not delete
var TitleIDValueTrimmed = null;// do not delete
var RegionValueTrimmed = null;// do not delete
var GenreValueTrimmed = null;// do not delete
var BaseSizeValueTrimmed = null;// do not delete
var VersionValueTrimmed = null;// do not delete
var UpdateSizeValueTrimmed = null;// do not delete
var DLCValueTrimmed = null;// do not delete
var BPValueTrimmed = null;// do not delete
var UploaderValueTrimmed = null;// do not delete
var DateValueTrimmed = null;// do not delete
var LinkValueTrimmed = null;// do not delete
var TestedValueTrimmed = null;// do not delete
var InfoValueTrimmed = null;// do not delete
var DKsMirrorValueTrimmed = null;// do not delete
var UpdateLinkValueTrimmed = null;// do not delete
var SubmitterValueTrimmed = null;// do not delete
// there is no trimmed announced in order to keep the announcing check fast
//var AnnouncedValueTrimmed = null;// do not delete
var FixesValueTrimmed = null;// do not delete

// the trimmed values of the columns of the sheet end

class SpreadsheetToolsObjectType
{
  constructor()
  {
    this.UserName = null;

    this.SpreadsheetApplication = null;
    this.Spreadsheet = null;
    this.SpreadsheetName = null;
    this.SpreadsheetUI = null;
    this.SpreadsheetProtection = null;
    this.SpreadsheetTimeZone = null;
    this.SpreadsheetDate = null;

    this.SpreadsheetReposNames = null;
    this.SpreadsheetInfosNames = null;

    this.Sheet = null;
    this.SheetName = null;
    this.SheetProtection = null;

    InitializeGeneralToolsObject_();

    // making a new property for the array type called "FixesSort"
    // , used for sorting the columns fixes that will be written into the Fixes column
    Array.prototype.FixesSort = (
      function()
      {
        var FunctionResult;
        var CurrentResult;

        var CurrentArray = this.clone();
        var CurrentArrayLength = CurrentArray.length;

        if (CurrentArrayLength > 1)
        {
          var CurrentCell;

          var ColumnIndex = -1;

          var CustomFixes = [];
          var CustomFixesAmount = 0;

          var FixesNormalDictionary = {};
          var FixesNormalDictionaryKeysAmount = 0;

          var FixesUrgentDictionary = {};
          var FixesUrgentDictionaryKeysAmount = 0;

          var NormalColumnsNames = [];
          var UrgentColumnsNames = [];

          CurrentResult = [];

          NormalColumnsNames.push(RegionColumnName);
          NormalColumnsNames.push(GenreColumnName);
          NormalColumnsNames.push(BaseSizeColumnName);
          NormalColumnsNames.push(VersionColumnName);
          NormalColumnsNames.push(UpdateSizeColumnName);
          NormalColumnsNames.push(DLCColumnName);
          NormalColumnsNames.push(BPColumnName);
          NormalColumnsNames.push(UploaderColumnName);
          NormalColumnsNames.push(DateColumnName);
          NormalColumnsNames.push(TestedColumnName);
          NormalColumnsNames.push(SubmitterColumnName);

          UrgentColumnsNames.push(TitleIDColumnName);
          UrgentColumnsNames.push(LinkColumnName);
          UrgentColumnsNames.push(DKsMirrorColumnName);
          UrgentColumnsNames.push(UpdateLinkColumnName);

          for (let CurrentArrayIndex = 0; CurrentArrayIndex < CurrentArrayLength; CurrentArrayIndex++)
          {
            CurrentCell = CurrentArray[CurrentArrayIndex];

            switch (CurrentCell)
            {
              case TitleIDColumnName:
                ColumnIndex = TitleIDColumn;

                break;

              case RegionColumnName:
                ColumnIndex = RegionColumn;

                break;

              case GenreColumnName:
                ColumnIndex = GenreColumn;

                break;

              case BaseSizeColumnName:
                ColumnIndex = BaseSizeColumn;

                break;

              case VersionColumnName:
                ColumnIndex = VersionColumn;

                break;

              case UpdateSizeColumnName:
                ColumnIndex = UpdateSizeColumn;

                break;

              case DLCColumnName:
                ColumnIndex = DLCColumn;

                break;

              case BPColumnName:
                ColumnIndex = BPColumn;

                break;

              case UploaderColumnName:
                ColumnIndex = UploaderColumn;

                break;

              case DateColumnName:
                ColumnIndex = DateColumn;

                break;

              case LinkColumnName:
                ColumnIndex = LinkColumn;

                break;

              case TestedColumnName:
                ColumnIndex = TestedColumn;

                break;

              case DKsMirrorColumnName:
                ColumnIndex = DKsMirrorColumn;

                break;

              case UpdateLinkColumnName:
                ColumnIndex = UpdateLinkColumn;

                break;

              case SubmitterColumnName:
                ColumnIndex = SubmitterColumn;

                break;

              default:
                CustomFixes.push(CurrentCell);
                CustomFixesAmount++;

                break;
            }

            if (ColumnIndex != -1)
            {
              if (NormalColumnsNames.includes(CurrentCell) == true)
              {
                FixesNormalDictionary['K' + ColumnIndex.pad(2)] = CurrentCell;
                FixesNormalDictionaryKeysAmount++;
              }
              else if (UrgentColumnsNames.includes(CurrentCell) == true)
              {
                FixesUrgentDictionary['K' + ColumnIndex.pad(2)] = CurrentCell;
                FixesUrgentDictionaryKeysAmount++;
              }

              ColumnIndex = -1;
            }
          }
            
          if (FixesUrgentDictionaryKeysAmount > 0)
          {
            FixesUrgentDictionary = GeneralToolsObject.SortDictionary(FixesUrgentDictionary);

            for (let Key in FixesUrgentDictionary)
              CurrentResult.push(FixesUrgentDictionary[Key]);
          }

          if (FixesNormalDictionaryKeysAmount > 0)
          {
            FixesNormalDictionary = GeneralToolsObject.SortDictionary(FixesNormalDictionary);

            for (let Key in FixesNormalDictionary)
              CurrentResult.push(FixesNormalDictionary[Key]);
          }

          if (CustomFixesAmount > 0)
            CurrentResult = CustomFixes.concat(CurrentResult);
        }
        else
          CurrentResult = CurrentArray;

        FunctionResult = CurrentResult;

        return FunctionResult;
      }
    );
  }
  
  // checking the name of the user that has launched the script
  // (searching in the user's contacts for himself using his email
  // and if found then returning the name of the contact (preferably full name)
  // , if not then returning the user's email's user name (for example: "example@gmail.com -> example"))
  CheckUserName()
  {
    if (this.UserName == null)
    {
      // getting the user that launched the script
      var User = Session.getEffectiveUser();

      if (User != null)
      {
        // getting the email of the user
        var UserEmail = User.getEmail();// example@gmail.com

        if (
          UserEmail != null
          && GeneralToolsObject.VariableIsString(UserEmail) == true
          && UserEmail.length > 0
        )
        {
          // getting the contact that has the user email in it, if there's any
          var UserContact = ContactsApp.getContact(UserEmail);

          // note:
          // using "let UserName" instead of "var UserName" because UserName is being used twice
          // , and having 2 variables with the same name defined as var in the same function is a bad exercise in javascript
          // (since var is all accross the function if it's inside the function (var can also be used for global variables)
          // , and let is only from the block level of the function and below it)

          // if the user has himself in its contacts
          if (UserContact != null)
          {
            // check the user's full name
            var UserFullName = UserContact.getFullName();// ExampleFirst ExampleLast

            if (
              UserFullName != null
              && GeneralToolsObject.VariableIsString(UserFullName) == true
              && UserFullName.length > 0
            )
              this.UserName = UserFullName;// ExampleFirst ExampleLast
            else
            {
              // check the user name of the contact only if failed to check the user's full name
              let UserName = UserContact.getGivenName();// Example

              if (
                UserName != null
                && GeneralToolsObject.VariableIsString(UserName) == true
                && UserName.length > 0
              )
                this.UserName = UserName;// Example
            }
          }
          else
          {
            // normally this returns the user's email's user name
            let UserName = User.getUsername();// example

            if (
              UserName != null
              && GeneralToolsObject.VariableIsString(UserName) == true
              && UserName.length > 0
            )
              this.UserName = UserName;// example
          }
        }
      }
    }
  }

  // checking the name of the spreadsheet
  CheckSpreadsheetName()
  {
    this.SpreadsheetName = this.Spreadsheet != null ? this.Spreadsheet.getName() : null;
  }

  // checking the spreadsheet
  CheckSpreadsheet()
  {
    this.SpreadsheetApplication = SpreadsheetApp;

    // gets the active spreadsheet from the spreadsheet application
    this.Spreadsheet = this.SpreadsheetApplication.getActiveSpreadsheet();

    this.CheckSpreadsheetName();

    // gets the user interface of the spreadsheet from the spreadsheet application
    this.SpreadsheetUI = this.SpreadsheetApplication.getUi();
    this.SpreadsheetProtection = null;
    this.SpreadsheetTimeZone = null;
    this.SpreadsheetDate = null;

    this.SpreadsheetReposNames = [];

    this.SpreadsheetReposNames.push(4RepoListName);
    this.SpreadsheetReposNames.push(4AppsListName);
    this.SpreadsheetReposNames.push(2To4RepoListName);

    this.SpreadsheetInfosNames = [];

    this.SpreadsheetInfosNames.push(4RegionInfo);
    this.SpreadsheetInfosNames.push(GameAppInfo);
  }

  // checks the name of the current sheet
  CheckSheetName()
  {
    this.SheetName = this.Sheet != null ? this.Sheet.getName() : null;
  }

  // searches the sheet by its name in the spreadsheet
  CheckSheet(SheetName)
  {
    if (
      SheetName != null
      && GeneralToolsObject.VariableIsString(SheetName) == true
      && SheetName.length > 0
    )
    {
      if (this.Spreadsheet == null)
        this.CheckSpreadsheet();

      if (this.Spreadsheet != null)
      {
        this.Sheet = this.Spreadsheet.getSheetByName(SheetName);
        this.SheetName = SheetName;
        this.SheetProtection = null;
      }
    }
  }

  // checks the active sheet in the spreadsheet
  CheckActiveSheet()
  {
    if (this.Spreadsheet == null)
      this.CheckSpreadsheet();

    if (this.Spreadsheet != null)
    {
      this.Sheet = this.Spreadsheet.getActiveSheet();

      this.CheckSheetName();
      
      this.SheetProtection = null;
    }
  }

  // checks the fixed name of the sheet
  CheckSheetFixedName()
  {
    return (
      this.SheetName == 4RepoListName
      ? "4"
      : (
        this.SheetName == 4AppsListName
        ? "4APP"
        : (
          this.SheetName == 2To4RepoListName
          ? "2"
          : (
            this.SheetName == 4RegionInfo
            ? "4INFO"
            : (
              this.SheetName == GameAppInfo
              ? "GAMEAPPINFO"
              : null
            )
          )
        )
      )
    );
  }

  // checks the protections of the spreadsheet
  CheckSpreadsheetProtections()
  {
    if (this.Spreadsheet == null)
      this.CheckSpreadsheet();

    return (
      this.Spreadsheet != null
      ? this.Spreadsheet.getProtections(this.SpreadsheetApplication.ProtectionType.SHEET)
      : null
    );
  }

  // protects the spreadsheet (if there aren't any protections in it)
  ProtectSpreadsheet()
  {
    // used mainly for the spreadsheet protection variable check
    if (this.Spreadsheet == null)
      this.CheckSpreadsheet();

    if (this.SpreadsheetProtection == null)
    {
      var SpreadsheetProtections = this.CheckSpreadsheetProtections();

      if (SpreadsheetProtections != null && SpreadsheetProtections.length == 0)
        this.SpreadsheetProtection = this.Spreadsheet.protect();
    }
  }

  // unprotects the spreadsheet (if there is a single protection in it)
  UnprotectSpreadsheet()
  {
    var SpreadsheetProtections = this.CheckSpreadsheetProtections();

    // this is just for reference, the CheckSpreadsheetProtections function already does this check
    if (this.Spreadsheet == null)
      this.CheckSpreadsheet();

    if (SpreadsheetProtections != null && SpreadsheetProtections.length == 1)
      if (this.SpreadsheetProtection != null)
      {
        if (this.SpreadsheetProtection.canEdit() == true)
        {
          this.SpreadsheetProtection.remove();

          this.SpreadsheetProtection = null;
        }
      }
      else if (SpreadsheetProtections[0].canEdit() == true)
        SpreadsheetProtections[0].remove();
  }

  // checks the protections of the current sheet
  CheckSheetProtections()
  {
    return this.Sheet != null ? this.Sheet.getProtections(this.SpreadsheetApplication.ProtectionType.SHEET) : null;
  }

  // protects the current sheet (if there aren't any protections in it)
  ProtectSheet()
  {
    if (this.SheetProtection == null)
    {
      var SheetProtections = this.CheckSheetProtections();

      if (SheetProtections != null && SheetProtections.length == 0)
        this.SheetProtection = this.Sheet.protect();
    }
  }

  // unprotects the current sheet (if there is a single protection in it)
  UnprotectSheet()
  {
    var SheetProtections = this.CheckSheetProtections();

    if (SheetProtections != null && SheetProtections.length == 1)
      if (this.SheetProtection != null)
      {
        if (this.SheetProtection.canEdit() == true)
        {
          this.SheetProtection.remove();

          this.SheetProtection = null;
        }
      }
      else if (SheetProtections[0].canEdit() == true)
        SheetProtections[0].remove();
  }

  // checks the timezone of the spreadsheet
  CheckSpreadsheetTimeZone()
  {
    if (this.SpreadsheetTimeZone == null)
    {
      if (this.Spreadsheet == null)
        this.CheckSpreadsheet();

      if (this.Spreadsheet != null)
        this.SpreadsheetTimeZone = this.Spreadsheet.getSpreadsheetTimeZone();
    }
  }

  // checks the current date of the spreadsheet
  CheckSpreadsheetDate()
  {
    if (this.SpreadsheetDate == null)
      if (
        SpreadsheetDateFormat != undefined
        && SpreadsheetDateFormat != null
        && GeneralToolsObject.VariableIsString(SpreadsheetDateFormat) == true
        && SpreadsheetDateFormat.length > 0
      )
      {
        // this is just for reference since the SpreadsheetTimeZone function already does this check
        if (this.Spreadsheet == null)
          this.CheckSpreadsheet();

        if (this.SpreadsheetTimeZone == null)
          this.CheckSpreadsheetTimeZone();

        if (
          this.SpreadsheetTimeZone != null
          && GeneralToolsObject.VariableIsString(this.SpreadsheetTimeZone) == true
          && this.SpreadsheetTimeZone.length > 0
        )
        {
          var FormattedDate = Utilities.formatDate(new Date(), this.SpreadsheetTimeZone, SpreadsheetDateFormat);

          if (
            FormattedDate != null
            && GeneralToolsObject.VariableIsString(FormattedDate) == true
            && FormattedDate.length > 0
          )
            this.SpreadsheetDate = FormattedDate;
        }
      }
  }

  // checks a cell from the current sheet
  CheckCell(RowsIndex, ColumnsIndex)
  {
    return this.Sheet != null ? this.Sheet.getRange(RowsIndex + 1, ColumnsIndex + 1) : null;
  }

  // applies a value to a cell from the current sheet
  // returns the range that got applied
  ApplyCell(RowsIndex, ColumnsIndex, Value)
  {
    var Cell = this.CheckCell(RowsIndex, ColumnsIndex);

    return Cell != null ? Cell.setValue(Value) : null;
  }

  // checks a range from the current sheet
  CheckRange(RowsIndex, ColumnsIndex, RowsAmount, ColumnsAmount)
  {
    return (
      this.Sheet != null
      ? this.Sheet.getRange(RowsIndex + 1, ColumnsIndex + 1, RowsAmount, ColumnsAmount)
      : null
    );
  }

  // applies values to a range from the current sheet
  // returns the range that got applied
  ApplyRange(RowsIndex, ColumnsIndex, RowsAmount, ColumnsAmount, Values)
  {
    var Range = this.CheckRange(RowsIndex, ColumnsIndex, RowsAmount, ColumnsAmount);

    return Range != null ? Range.setValues(Values) : null;
  }

  // applies values to a row from the current sheet
  ApplyRow(
    RowsIndex
    , ColumnsIndex
    , ColumnsAmount
    , ValuesType// 0 - Real Values, else - Formulas
    , Values
  )
  {
    var Value;

    var RowsValues = [];
    var RowValues = [];
    var RowValuesAmount = 0;
    var RowValue;

    var Cells;

    for (
      let CurrentColumnsIndex = ColumnsIndex;
      CurrentColumnsIndex < ColumnsIndex + ColumnsAmount;
      CurrentColumnsIndex++
    )
    {
      Value = Values[CurrentColumnsIndex - ColumnsIndex];

      if (
        ValuesType == 0
        || (Value.length > 0 && Value[0] == '=')
      )
      {
        RowValue = Value;

        RowValues.push(RowValue);
        RowValuesAmount++;
      }
      else if (RowValuesAmount > 0)
      {
        RowsValues.push(RowValues);

        if (dry_run == false)
        {
          Cells =
            this.CheckRange(
              RowsIndex
              , CurrentColumnsIndex - RowValuesAmount
              , 1
              , RowValuesAmount
            );

          if (ValuesType == 0)
            Cells.setValues(RowsValues);
          else
            Cells.setFormulas(RowsValues);
        }

        RowsValues = [];
        RowValues = [];
        RowValuesAmount = 0;
      }
    }

    if (RowValuesAmount > 0)
    {
      RowsValues.push(RowValues);

      if (dry_run == false)
      {
        Cells =
          this.CheckRange(
            RowsIndex
            , ColumnsIndex + ColumnsAmount - RowValuesAmount
            , 1
            , RowValuesAmount
          );

        if (ValuesType == 0)
          Cells.setValues(RowsValues);
        else
          Cells.setFormulas(RowsValues);
      }
    }
  }

  // applies values to a column from the current sheet
  ApplyColumn(
    RowsIndex
    , ColumnsIndex
    , RowsAmount
    , ValuesType// 0 - Real Values, else - Formulas
    , Values
  )
  {
    var Value;

    var RowsValues = [];
    var RowsValuesAmount = 0;
    var RowValues = null;
    var RowValue;

    var Cells;

    for (
      let CurrentRowsIndex = RowsIndex;
      CurrentRowsIndex < RowsIndex + RowsAmount;
      CurrentRowsIndex++
    )
    {
      Value = Values[CurrentRowsIndex - RowsIndex];

      if (
        ValuesType == 0
        || (Value.length > 0 && Value[0] == '=')
      )
      {
        RowValue = Value;

        RowValues = [];

        RowValues.push(RowValue);

        RowsValues.push(RowValues);
        RowsValuesAmount++;
      }
      else if (RowsValuesAmount > 0)
      {
        if (dry_run == false)
        {
          Cells =
            this.CheckRange(
              CurrentRowsIndex - RowsValuesAmount
              , ColumnsIndex
              , RowsValuesAmount
              , 1
            );

          if (ValuesType == 0)
            Cells.setValues(RowsValues);
          else
            Cells.setFormulas(RowsValues);
        }

        RowsValues = [];
        RowsValuesAmount = 0;
        RowValues = null;
      }
    }

    if (RowsValuesAmount > 0)
      if (dry_run == false)
      {
        Cells =
          this.CheckRange(
            RowsIndex + RowsAmount - RowsValuesAmount
            , ColumnsIndex
            , RowsValuesAmount
            , 1
          );

        if (ValuesType == 0)
          Cells.setValues(RowsValues);
        else
          Cells.setFormulas(RowsValues);
      }
  }

  // checks rows of the current sheet
  CheckRows(
    ValuesType// 0 - Real Values, 1 - Display Values, else - Formulas
    , StartingRow = -1
    , StartingColumn = -1
    , EndingRow = -1
    , EndingColumn = -1
  )
  {
    var Rows = null;

    if (this.Sheet != null)
    {
      var SheetDataRange = this.Sheet.getDataRange();

      if (SheetDataRange != null)
      {
        var CurrentStartingRow = 0;
        var CurrentEndingRow = SheetDataRange.getNumRows() - 1;

        var ErrorFound = false;

        if (
          StartingRow != null
          && GeneralToolsObject.VariableIsNumber(StartingRow) == true
          && GeneralToolsObject.VariableIsString(StartingRow) == false
          && StartingRow != -1
          && StartingRow > CurrentStartingRow
        )
          if (StartingRow <= CurrentEndingRow)
            CurrentStartingRow = StartingRow;
          else
            ErrorFound = true;

        if (ErrorFound == false)
        {
          if (
            EndingRow != null
            && GeneralToolsObject.VariableIsNumber(EndingRow) == true
            && GeneralToolsObject.VariableIsString(EndingRow) == false
            && EndingRow != -1
            && EndingRow < CurrentEndingRow
          )
            if (EndingRow >= CurrentStartingRow)
              CurrentEndingRow = EndingRow;
            else
              ErrorFound = true;

          if (ErrorFound == false)
          {
            var CurrentStartingColumn = 0;
            var CurrentEndingColumn = SheetDataRange.getNumColumns() - 1;

            if (
              StartingColumn != null
              && GeneralToolsObject.VariableIsNumber(StartingColumn) == true
              && GeneralToolsObject.VariableIsString(StartingColumn) == false
              && StartingColumn != -1
              && StartingColumn > CurrentStartingColumn
            )
              if (StartingColumn <= CurrentEndingColumn)
                CurrentStartingColumn = StartingColumn;
              else
                ErrorFound = true;

            if (ErrorFound == false)
            {
              if (
                EndingColumn != null
                && GeneralToolsObject.VariableIsNumber(EndingColumn) == true
                && GeneralToolsObject.VariableIsString(EndingColumn) == false
                && EndingColumn != -1
                && EndingColumn < CurrentEndingColumn
              )
                if (EndingColumn >= CurrentStartingColumn)
                  CurrentEndingColumn = EndingColumn;
                else
                  ErrorFound = true;

              if (ErrorFound == false)
              {
                var RowsRange;

                if (CurrentEndingRow < CurrentStartingRow)
                  CurrentEndingRow = CurrentStartingRow;

                if (CurrentEndingColumn < CurrentStartingColumn)
                  CurrentEndingColumn = CurrentStartingColumn;

                RowsRange =
                  this.CheckRange(
                    CurrentStartingRow
                    , CurrentStartingColumn
                    , CurrentEndingRow - CurrentStartingRow + 1
                    , CurrentEndingColumn - CurrentStartingColumn + 1
                  );

                if (RowsRange != null)
                {
                  var RowsRangeValues =
                    ValuesType == 0
                    ? RowsRange.getValues()
                    : (
                      ValuesType == 1
                      ? RowsRange.getDisplayValues()
                      : RowsRange.getFormulas()
                    );

                  if (RowsRangeValues != null && RowsRangeValues.length > 0)
                    Rows = RowsRangeValues;
                }
              }
            }
          }
        }
      }
    }

    return Rows;
  }

  // checks the locations of the columns based on the inputted rows
  CheckColumnsLocations(Rows)
  {
    TitleColumn = -1;
    TitleIDColumn = -1;
    RegionColumn = -1;
    GenreColumn = -1;
    BaseSizeColumn = -1;
    VersionColumn = -1;
    UpdateSizeColumn = -1;
    DLCColumn = -1;
    BPColumn = -1;
    UploaderColumn = -1;
    DateColumn = -1;
    LinkColumn = -1;
    TestedColumn = -1;
    InfoColumn = -1;
    DKsMirrorColumn = -1;
    UpdateLinkColumn = -1;
    SubmitterColumn = -1;
    AnnouncedColumn = -1;
    FixesColumn = -1;

    if (Rows != null)
    {
      var RowsAmount = Rows.length;

      if (RowsAmount > 0)
      {
        var Row = Rows[0];
        var RowColumns = Row;

        if (RowColumns != null)
        {
          var RowColumnsAmount = RowColumns.length;

          if (RowColumnsAmount > 0)
          {
            var RowColumn;

            for (let RowColumnsIndex = 0; RowColumnsIndex < RowColumnsAmount; RowColumnsIndex++)
            {
              RowColumn = RowColumns[RowColumnsIndex];

              switch (RowColumn)
              {
                case TitleColumnName:
                  TitleColumn = RowColumnsIndex;

                  break;

                case TitleIDColumnName:
                  TitleIDColumn = RowColumnsIndex;

                  break;

                case RegionColumnName:
                  RegionColumn = RowColumnsIndex;

                  break;

                case GenreColumnName:
                  GenreColumn = RowColumnsIndex;

                  break;

                case BaseSizeColumnName:
                  BaseSizeColumn = RowColumnsIndex;

                  break;

                case VersionColumnName:
                  VersionColumn = RowColumnsIndex;

                  break;

                case UpdateSizeColumnName:
                  UpdateSizeColumn = RowColumnsIndex;

                  break;

                case DLCColumnName:
                  DLCColumn = RowColumnsIndex;

                  break;

                case BPColumnName:
                  BPColumn = RowColumnsIndex;

                  break;

                case UploaderColumnName:
                  UploaderColumn = RowColumnsIndex;

                  break;

                case DateColumnName:
                  DateColumn = RowColumnsIndex;

                  break;

                case LinkColumnName:
                  LinkColumn = RowColumnsIndex;

                  break;

                case TestedColumnName:
                  TestedColumn = RowColumnsIndex;

                  break;

                case InfoColumnName:
                  InfoColumn = RowColumnsIndex;

                  break;

                case DKsMirrorColumnName:
                  DKsMirrorColumn = RowColumnsIndex;

                  break;

                case UpdateLinkColumnName:
                  UpdateLinkColumn = RowColumnsIndex;

                  break;

                case SubmitterColumnName:
                  SubmitterColumn = RowColumnsIndex;

                  break;

                case AnnouncedColumnName:
                  AnnouncedColumn = RowColumnsIndex;

                  break;

                case FixesColumnName:
                  FixesColumn = RowColumnsIndex;

                  break;
              }
            }
          }
        }
      }
    }
  }

  // checks a hyper link of a cell in the current sheet
  CheckHyperlink(Row, Column)
  {
    var Hyperlink = null;

    if (Row >= 0 && Column >= 0)
    {
      var Cell = this.CheckCell(Row, Column);

      if (Cell != null)
      {
        var LinkRichText = Cell.getRichTextValue();

        if (LinkRichText != null)
        {
          var LinkUrl = LinkRichText.getLinkUrl();

          if (
            LinkUrl != null
            && GeneralToolsObject.VariableIsString(LinkUrl) == true
            && LinkUrl.length > 0
          )
            Hyperlink = LinkUrl;
        }
      }
    }

    return Hyperlink;
  }
}

// initializing the spreadsheet tools object
function InitializeSpreadsheetToolsObject_()
{
  if (SpreadsheetToolsObject == null)
    SpreadsheetToolsObject = new SpreadsheetToolsObjectType();
}

// shows the user name of the person who launched the script
function ShowUserName()
{
  InitializeGeneralToolsObject_();
  InitializeSpreadsheetToolsObject_();

  SpreadsheetToolsObject.CheckSpreadsheet();

  if (SpreadsheetToolsObject.SpreadsheetUI != null)
  {
    SpreadsheetToolsObject.CheckUserName();

    if (
      SpreadsheetToolsObject.UserName != null
      && GeneralToolsObject.VariableIsString(SpreadsheetToolsObject.UserName) == true
      && SpreadsheetToolsObject.UserName.length > 0
    )
      SpreadsheetToolsObject.SpreadsheetUI.alert(SpreadsheetToolsObject.UserName);
  }
}

// adding the spreadsheet tools menu to the spreadsheet UI
function AddSpreadsheetToolsMenu()
{
  InitializeSpreadsheetToolsObject_();

  SpreadsheetToolsObject.CheckSpreadsheet();
  
  if (SpreadsheetToolsObject.SpreadsheetUI != null)
  {
    var SpreadsheetUIMenu = SpreadsheetToolsObject.SpreadsheetUI.createMenu("Spreadsheet Tools");

    if (SpreadsheetUIMenu != null)
    {
      SpreadsheetUIMenu.addItem("Show User Name", "ShowUserName");
      SpreadsheetUIMenu.addToUi();
    }
  }
}
