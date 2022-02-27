// user config start

// update the rows every 200 replacements
// set as 0 or under for infinite
var FixSpreadsheetFilesSizesMaxFoundMatchesAmount = 200;

// user config end

var FixSheetFilesSizesObject = null;// do not delete

class FixSpreadsheetFilesSizesColumnObjectType
{
  constructor(Column)
  {
    this.Column = Column;
    this.FoundMatch = false;
    this.StartingRow = -1;
    this.EndingRow = -1;
    this.FileSizeNumbers = null;
    this.FileSizeNumberFormat = null;

    InitializeGeneralToolsObject_();
    InitializeSpreadsheetToolsObject_();
  }

  // starting the matches counting
  Start(RowsIndex, FileSizeNumberFormat)
  {
    if (this.FoundMatch == false)
    {
      this.FoundMatch = true;
      this.StartingRow = RowsIndex;
      this.FileSizeNumberFormat = FileSizeNumberFormat;
    }
  }

  // adding a file size number to the file size numbers array
  AddFileSizeNumber(FileSizeNumber)
  {
    if (this.FoundMatch == true)
    {
      if (this.FileSizeNumbers == null)
        this.FileSizeNumbers = [];

      this.FileSizeNumbers.push(FileSizeNumber);

      if (this.EndingRow == -1)
        this.EndingRow = this.StartingRow;
      else
        this.EndingRow++;
    }
  }

  // ending the matches counting
  End()
  {
    if (this.FoundMatch == true)
    {
      if (this.EndingRow != -1)
      {
        var RowsValues = [];
        var RowsValuesAmount = 0;// not necessary, but just for reference
        var RowValues = null;
        var RowValue;

        var FileSizeNumber;

        for (let RowsIndex = this.StartingRow; RowsIndex <= this.EndingRow; RowsIndex++)
        {
          FileSizeNumber = this.FileSizeNumbers[RowsIndex - this.StartingRow];

          RowValue = (
            GeneralToolsObject.VariableIsString(FileSizeNumber) == true
            ? Number(FileSizeNumber)
            : FileSizeNumber
          );

          RowValues = [];

          RowValues.push(RowValue);

          RowsValues.push(RowValues);
          RowsValuesAmount++;
        }

        if (RowsValuesAmount > 0)
          if (dry_run == false)
          {
            var Cells =
              SpreadsheetToolsObject.CheckRange(
                this.StartingRow
                , this.Column
                , this.EndingRow - this.StartingRow + 1
                , 1
              );

            Cells.setValues(RowsValues);
            Cells.setNumberFormat(this.FileSizeNumberFormat);
          }

        this.EndingRow = -1;
      }

      this.FoundMatch = false;

      if (this.StartingRow != -1)
        this.StartingRow = -1;

      if (this.FileSizeNumbers != null)
        this.FileSizeNumbers = null;

      if (this.FileSizeNumberFormat != null)
        this.FileSizeNumberFormat = null;
    }
  }
}

class FixSheetFilesSizesObjectType
{
  constructor()
  {
    InitializeGeneralToolsObject_();
    InitializeSpreadsheetToolsObject_();
    InitializeMutualMethodsObject_();
  }

  // fixing the sheet files sizes
  FixSheetFilesSizes()
  {
    var Rows =
      MutualMethodsObject.CheckRows(
        true
        , true
        , false
        , 0
        //, (
        //  MaxRowsPerCheck != undefined
        //  && MaxRowsPerCheck != null
        //  && GeneralToolsObject.VariableIsNumber(MaxRowsPerCheck) == true
        //  && GeneralToolsObject.VariableIsString(MaxRowsPerCheck) == false
        //  && MaxRowsPerCheck > 0
        //  ? MaxRowsPerCheck - 1
        //  : -1
        //)
        , -1// do a full check once for the likelyhood that there's nothing to be replaced
      );

    // unprotecting the sheet in case it's protected, might happen if the script crashed before
    //SpreadsheetToolsObject.UnprotectSheet();

    if (
      Rows != null
      && Rows.length == 4
      && Rows[0] != null
      && Rows[1] != null
      && Rows[2] != null
      && Rows[3] != null
      && GeneralToolsObject.VariableIsNumber(Rows[1]) == true
      && GeneralToolsObject.VariableIsString(Rows[1]) == false
      && Rows[1] != -1
      && GeneralToolsObject.VariableIsNumber(Rows[3]) == true
      && GeneralToolsObject.VariableIsString(Rows[3]) == false
      && Rows[3] != -1
    )
    {
      var RowsRealValues = Rows[0];
      var RowsRealValuesAmount = Rows[1];
      var RowsDisplayValues = Rows[2];
      var RowsDisplayValuesAmount = Rows[3];
      
      var RowsIndex = 1;
      var RealRowsIndex = RowsIndex;

      var RowRealValues;
      var RowRealValuesColumns;
      var RowRealValuesColumnsAmount;

      var RowDisplayValues = null;
      var RowDisplayValuesColumns = null;
      var RowDisplayValuesColumnsAmount = 0;

      var MaxFoundMatchesAmount =
        FixSpreadsheetFilesSizesMaxFoundMatchesAmount != undefined
        && FixSpreadsheetFilesSizesMaxFoundMatchesAmount != null
        && GeneralToolsObject.VariableIsNumber(FixSpreadsheetFilesSizesMaxFoundMatchesAmount) == true
        && GeneralToolsObject.VariableIsString(FixSpreadsheetFilesSizesMaxFoundMatchesAmount) == false
        ? FixSpreadsheetFilesSizesMaxFoundMatchesAmount
        : -1;

      var ColumnObjects = [];
      var ColumnObjectsAmount = 0;
      var ColumnObject;

      var FileSizeNumberWithFormat;
      
      var FileSizeDoubleNumber = null;
      var FileSizeNumberFormat = null;

      var FoundMatchesAmount = 0;

      var FoundMatch = false;

      var RecheckRows = false;

      var MethodFound = false;

      SpreadsheetToolsObject.CheckColumnsLocations(RowsDisplayValues);

      ColumnObject = new FixSpreadsheetFilesSizesColumnObjectType(BaseSizeColumn);

      ColumnObjects.push(ColumnObject);
      ColumnObjectsAmount++;

      ColumnObject = new FixSpreadsheetFilesSizesColumnObjectType(UpdateSizeColumn);

      ColumnObjects.push(ColumnObject);
      ColumnObjectsAmount++;

      ColumnObject = null;

      //SpreadsheetToolsObject.ProtectSheet();

      while (
        RowsIndex < RowsRealValuesAmount
        && RowsRealValues != null
        && RowsDisplayValues != null
        && RowsRealValuesAmount > 0
        && RowsRealValuesAmount == RowsDisplayValuesAmount
      )
      {
        RowRealValues = RowsRealValues[RowsIndex];
        RowRealValuesColumns = RowRealValues;
        RowRealValuesColumnsAmount = RowRealValuesColumns != null ? RowRealValuesColumns.length : 0;

        if (RowRealValuesColumnsAmount > 0)
        {
          RowDisplayValues = RowsDisplayValues[RowsIndex];
          RowDisplayValuesColumns = RowDisplayValues;
          RowDisplayValuesColumnsAmount = RowDisplayValuesColumns != null ? RowDisplayValuesColumns.length : 0;

          if (RowDisplayValuesColumnsAmount > 0)
            if (RowRealValuesColumnsAmount == RowDisplayValuesColumnsAmount)
              MethodFound = true;
        }

        if (MethodFound == true)
        {
          MethodFound = false;

          MutualMethodsObject.ResetValues();

          MutualMethodsObject.CheckTitle(RowDisplayValues);

          if (
            TitleValueTrimmed != null
            // using 1 because normally titles wouldn't be constructed of a single character
            && TitleValueTrimmed.length > 1
          )
            for (let ColumnObjectsIndex = 0; ColumnObjectsIndex < ColumnObjectsAmount; ColumnObjectsIndex++)
            {
              ColumnObject = ColumnObjects[ColumnObjectsIndex];

              if (ColumnObject.Column != -1)
              {
                FileSizeNumberWithFormat =
                  MutualMethodsObject.CheckFileSizeNumberWithFormat(
                    RowRealValues
                    , RowDisplayValues
                    , ColumnObject.Column
                  );

                if (FileSizeNumberWithFormat != null)
                {
                  if (
                    FileSizeNumberWithFormat.length == 2
                    && FileSizeNumberWithFormat[0] != null
                    && FileSizeNumberWithFormat[1] != null
                    && GeneralToolsObject.VariableIsNumber(FileSizeNumberWithFormat[0]) == true
                    && GeneralToolsObject.VariableIsString(FileSizeNumberWithFormat[0]) == false
                    && FileSizeNumberWithFormat[0] != -1
                    && GeneralToolsObject.VariableIsString(FileSizeNumberWithFormat[1]) == true
                    && FileSizeNumberWithFormat[1].length > 0
                  )
                  {
                    FoundMatch = true;

                    FoundMatchesAmount++;

                    FileSizeDoubleNumber = FileSizeNumberWithFormat[0];
                    FileSizeNumberFormat = FileSizeNumberWithFormat[1];
                  }

                  FileSizeNumberWithFormat = null;
                }
              }

              if (
                ColumnObject.FoundMatch == true
                && (
                  FoundMatch == false
                  || ColumnObject.EndingRow == -1
                  || ColumnObject.EndingRow + 1 != RealRowsIndex
                  || ColumnObject.FileSizeNumberFormat != FileSizeNumberFormat
                )
              )
              {
                ColumnObject.End();

                if (
                  MaxFoundMatchesAmount > 0
                  && FoundMatchesAmount > MaxFoundMatchesAmount
                )
                {
                  if (FoundMatch == true)
                    FoundMatch = false;

                  if (dry_run == false)
                    if (RecheckRows == false)
                      RecheckRows = true;
                  
                  FoundMatchesAmount = 0;
                }
              }

              // if a match has been found
              
              if (FoundMatch == true)
              {
                FoundMatch = false;

                if (ColumnObject.FoundMatch == false)
                  ColumnObject.Start(RealRowsIndex, FileSizeNumberFormat);

                ColumnObject.AddFileSizeNumber(FileSizeDoubleNumber);
              }

              ColumnObject = null;

              if (FileSizeDoubleNumber != null)
                FileSizeDoubleNumber = null;

              if (FileSizeNumberFormat != null)
                FileSizeNumberFormat = null;
            }
        }

        if (RecheckRows == true)
        {
          // go 10 rows back (at most) for safety

          RealRowsIndex -= 10;

          if (RealRowsIndex < 1)
            RealRowsIndex = 1;
        }
        else
        {
          RowsIndex++;
          RealRowsIndex++;
        }

        if (RecheckRows == true || RowsIndex == RowsRealValuesAmount)
        {
          Rows =
            MutualMethodsObject.CheckRows(
              true
              , true
              , false
              , RealRowsIndex
              , (
                MaxRowsPerCheck != undefined
                && MaxRowsPerCheck != null
                && GeneralToolsObject.VariableIsNumber(MaxRowsPerCheck) == true
                && GeneralToolsObject.VariableIsString(MaxRowsPerCheck) == false
                && MaxRowsPerCheck > 0
                ? RealRowsIndex + MaxRowsPerCheck - 1
                : -1
              )
            );

          if (
            Rows != null
            && Rows.length == 4
            && Rows[0] != null
            && Rows[1] != null
            && Rows[2] != null
            && Rows[3] != null
            && GeneralToolsObject.VariableIsNumber(Rows[1]) == true
            && GeneralToolsObject.VariableIsString(Rows[1]) == false
            && Rows[1] != -1
            && GeneralToolsObject.VariableIsNumber(Rows[3]) == true
            && GeneralToolsObject.VariableIsString(Rows[3]) == false
            && Rows[3] != -1
          )
          {
            RowsRealValues = Rows[0];
            RowsRealValuesAmount = Rows[1];
            RowsDisplayValues = Rows[2];
            RowsDisplayValuesAmount = Rows[3];
          }
          else
          {
            RowsRealValues = null;
            RowsRealValuesAmount = 0;
            RowsDisplayValues = null;
            RowsDisplayValuesAmount = 0;
          }

          if (RowsIndex != 0)
            RowsIndex = 0;

          if (RecheckRows == true)
            RecheckRows = false;
        }
      }

      if (ColumnObjectsAmount > 0)
      {
        for (let ColumnObjectsIndex = 0; ColumnObjectsIndex < ColumnObjectsAmount; ColumnObjectsIndex++)
        {
          ColumnObject = ColumnObjects[ColumnObjectsIndex];

          if (ColumnObject.FoundMatch == true)
            ColumnObject.End();
        }

        ColumnObjectsAmount = 0;
      }

      ColumnObjects = null;

      if (FoundMatchesAmount > 0)
        FoundMatchesAmount = 0;

      //SpreadsheetToolsObject.UnprotectSheet();
    }
  }
}

// initializing the fix sheet files sizes object
function InitializeFixSheetFilesSizesObject_()
{
  if (FixSheetFilesSizesObject == null)
    FixSheetFilesSizesObject = new FixSheetFilesSizesObjectType();
}
  
// fixing the active sheet files sizes
function FixActiveSheetFilesSizes()
{
  InitializeSpreadsheetToolsObject_();
  InitializeFixSheetFilesSizesObject_();

  SpreadsheetToolsObject.CheckSpreadsheet();
  SpreadsheetToolsObject.CheckActiveSheet();

  FixSheetFilesSizesObject.FixSheetFilesSizes();
}

// fixing all sheets' files sizes
function FixAllSheetsFilesSizes()
{
  var SpreadsheetReposNamesAmount;

  InitializeSpreadsheetToolsObject_();
  InitializeFixSheetFilesSizesObject_();

  SpreadsheetToolsObject.CheckSpreadsheet();

  if (SpreadsheetToolsObject.SpreadsheetReposNames != null)
  {
    SpreadsheetReposNamesAmount = SpreadsheetToolsObject.SpreadsheetReposNames.length;

    if (SpreadsheetReposNamesAmount > 0)
    {
      var SpreadsheetReposName;

      for (
        let SpreadsheetReposNamesIndex = 0;
        SpreadsheetReposNamesIndex < SpreadsheetReposNamesAmount;
        SpreadsheetReposNamesIndex++
      )
      {
        SpreadsheetReposName = SpreadsheetToolsObject.SpreadsheetReposNames[SpreadsheetReposNamesIndex];

        SpreadsheetToolsObject.CheckSheet(SpreadsheetReposName);

        FixSheetFilesSizesObject.FixSheetFilesSizes();
      }
    }
  }
}

// adding the fix spreadsheet files sizes menu to the spreadsheet UI
function AddFixSpreadsheetFilesSizesMenu()
{
  InitializeSpreadsheetToolsObject_();

  SpreadsheetToolsObject.CheckSpreadsheet();
  
  if (SpreadsheetToolsObject.SpreadsheetUI != null)
  {
    var SpreadsheetUIMenu = SpreadsheetToolsObject.SpreadsheetUI.createMenu("Fix Sizes");

    if (SpreadsheetUIMenu != null)
    {
      SpreadsheetUIMenu.addItem("Fix Active Sheet Files Sizes", "FixActiveSheetFilesSizes");
      SpreadsheetUIMenu.addSeparator();
      SpreadsheetUIMenu.addItem("Fix All Sheets Files Sizes", "FixAllSheetsFilesSizes");
      SpreadsheetUIMenu.addToUi();
    }
  }
}
