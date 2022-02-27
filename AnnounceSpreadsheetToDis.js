// user config start

// sends these columns in a dis message according to the order of the array
var AnnounceSpreadsheetToDisDisMessageColumns = [];

AnnounceSpreadsheetToDisDisMessageColumns.push(
  {
    Column: TitleColumnName
    // the text style of the title in the dis message
    , TextStyleCode: (
      ItalicsDisTextStyle.Code
      | BoldDisTextStyle.Code
    )
  }
);

AnnounceSpreadsheetToDisDisMessageColumns.push({Column: TitleIDColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: VersionColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: RegionColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: GenreColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: BaseSizeColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: UpdateSizeColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: DLCColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: BPColumnName, TextStyleCode: null});
//AnnounceSpreadsheetToDisDisMessageColumns.push({Column: TestedColumnName, TextStyleCode: null});
AnnounceSpreadsheetToDisDisMessageColumns.push({Column: InfoColumnName, TextStyleCode: null});

// user config end

var AnnounceSheetToDisObject = null;// do not delete

// announcing without sending a dis message or changing the date
var AnnounceSpreadsheetToDisQuietAnnounce = false;// do not delete

// rechecking the rows of which there are fixes needed, if there's any
var AnnounceSpreadsheetToDisRecheckFixes = true;// do not delete

// if true then not requiring title id in order for an entry to get announced
var AnnounceSpreadsheetToDisIgnoreTitleIDRequiredForAnnouncing = true;// do not delete

class AnnounceSheetToDisObjectType
{
  constructor()
  {
    InitializeGeneralToolsObject_();
    InitializeSpreadsheetToolsObject_();
    InitializeDisToolsObject_();
    InitializeMutualMethodsObject_();
  }

  // announcing the current sheet to dis
  AnnounceSheetToDis()
  {
    var Rows =
      MutualMethodsObject.CheckRows(
        true
        , true
        , true
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
      && Rows.length == 6
      && Rows[0] != null
      && Rows[1] != null
      && Rows[2] != null
      && Rows[3] != null
      && Rows[4] != null
      && Rows[5] != null
      && GeneralToolsObject.VariableIsNumber(Rows[1]) == true
      && GeneralToolsObject.VariableIsString(Rows[1]) == false
      && Rows[1] != -1
      && GeneralToolsObject.VariableIsNumber(Rows[3]) == true
      && GeneralToolsObject.VariableIsString(Rows[3]) == false
      && Rows[3] != -1
      && GeneralToolsObject.VariableIsNumber(Rows[5]) == true
      && GeneralToolsObject.VariableIsString(Rows[5]) == false
      && Rows[5] != -1
    )
    {
      var RowsRealValues = Rows[0];
      var RowsRealValuesAmount = Rows[1];
      var RowsDisplayValues = Rows[2];
      var RowsDisplayValuesAmount = Rows[3];
      var RowsFormulas = Rows[4];
      var RowsFormulasAmount = Rows[5];
      
      var RowsIndex = 1;
      var RealRowsIndex = RowsIndex;

      var RowRealValues;
      var RowRealValuesColumns;
      var RowRealValuesColumnsAmount;

      var RowDisplayValues = null;
      var RowDisplayValuesColumns = null;
      var RowDisplayValuesColumnsAmount = 0;

      var RowFormulas = null;
      var RowFormulasColumns = null;
      var RowFormulasColumnsAmount = 0;

      var QuietAnnounce = AnnounceSpreadsheetToDisQuietAnnounce;
      var RecheckFixes = AnnounceSpreadsheetToDisRecheckFixes;
      var IgnoreTitleIDRequiredForAnnouncing = AnnounceSpreadsheetToDisIgnoreTitleIDRequiredForAnnouncing;

      var DisMessageArray;
      var DisMessage;

      var DisMessageColumns = AnnounceSpreadsheetToDisDisMessageColumns;
      var DisMessageColumnsAmount = 
        DisMessageColumns != undefined
        && DisMessageColumns != null
        ? DisMessageColumns.length
        : 0;

      var DisMessageColumnsCell;
      var DisMessageColumnsCellColumn = null;
      var DisMessageColumnsCellColumnValue;
      var DisMessageColumnsCellTextStyleCode = null;

      var FileSizeNumberWithFormat;
      
      var FileSizeDoubleNumber;
      var FileSizeNumberFormat;

      var ColumnsFixes = null;
      var ColumnsFixesAmount = 0;

      var ColumnsFixesText = null;

      var ColumnsFixesDateIndex = -1;

      var ColumnsAdditionalFixes = null;
      var ColumnsAdditionalFixesAmount = 0;

      var ColumnsUrgentFixes = null;
      var ColumnsUrgentFixesAmount = 0;

      var FixesValueTrimmedSplitted;
      var FixesValueTrimmedSplittedAmount;
      var FixesValueTrimmedSplittedIndex;
      var FixesValueTrimmedSplittedCell;
      var FixesValueTrimmedSplittedCellValue;
      var FixesValueTrimmedSplittedCellValueInvalid = false;
      var FixesValueTrimmedSplittedLinkIndex = -1;
      var FixesValueTrimmedSplittedDKsMirrorIndex = -1;
      var FixesValueTrimmedSplittedUpdateLinkIndex = -1;

      var FoundMatch = false;

      var RecheckRows = false;

      var MethodFound = false;

      SpreadsheetToolsObject.CheckColumnsLocations(RowsDisplayValues);

      ////SpreadsheetToolsObject.UnprotectSheet();

      while (
        RowsIndex < RowsRealValuesAmount
        && RowsRealValues != null
        && RowsDisplayValues != null
        && RowsFormulas != null
        && RowsRealValuesAmount > 0
        && RowsRealValuesAmount == RowsDisplayValuesAmount
        && RowsRealValuesAmount == RowsFormulasAmount
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
            {
              RowFormulas = RowsFormulas[RowsIndex];
              RowFormulasColumns = RowFormulas;
              RowFormulasColumnsAmount = RowFormulasColumns != null ? RowFormulasColumns.length : 0;

              if (RowFormulasColumnsAmount > 0)
                if (RowRealValuesColumnsAmount == RowFormulasColumnsAmount)
                  MethodFound = true;
            }
        }

        if (MethodFound == true)
        {
          MethodFound = false;

          MutualMethodsObject.ResetValues();

          MutualMethodsObject.CheckAnnounced(RowDisplayValues);

          if (
            AnnouncedValue != null
            && GeneralToolsObject.VariableIsString(AnnouncedValue) == true
            && (
              AnnouncedValue.length == 2
              && AnnouncedValue == "No"
              || AnnouncedValue.length == NATextLength
              && AnnouncedValue == NAText
            )
          )
            MethodFound = true;
          else if (RecheckFixes == true)
          {
            MutualMethodsObject.CheckFixes(RowDisplayValues);

            if (
              FixesValueTrimmed != null
              && FixesValueTrimmed.length > 0
              && (
                FixesValueTrimmed.length == 3 && FixesValueTrimmed == "Fix"
                || FixesValueTrimmed.length != 4
                || FixesValueTrimmed != "None"
              )
            )
              MethodFound = true;
          }

          if (MethodFound == true)
          {
            MethodFound = false;
            
            MutualMethodsObject.CheckTitle(RowDisplayValues);

            // checking that the entry has title and title id or link at least in order to go on

            if (
              TitleValueTrimmed != null
              // using 1 because normally titles wouldn't be constructed of a single character
              && TitleValueTrimmed.length > 1
            )
            {
              if (TitleIDColumn != -1)
                if (
                  RowDisplayValues[TitleIDColumn] != null
                  && RowDisplayValues[TitleIDColumn].length > 0
                )
                {
                  MutualMethodsObject.CheckTitleID(RowDisplayValues, false);

                  if (
                    TitleIDValueTrimmed != null
                    && TitleIDValueTrimmed.length > 0
                    && (
                      TitleIDValueTrimmed.length != NATextLength
                      || TitleIDValueTrimmed != NAText
                    )
                  )
                  {
                    MethodFound = true;

                    TitleIDValue = null;
                    TitleIDValueTrimmed = null;
                  }
                }

              if (MethodFound == false)
              {
                if (LinkColumn != -1)
                  if (
                    RowDisplayValues[LinkColumn] != null
                    && RowDisplayValues[LinkColumn].length > 0
                  )
                  {
                    MutualMethodsObject.CheckLinkB(RowDisplayValues);

                    if (
                      LinkValueTrimmed != null
                      && LinkValueTrimmed.length > 0
                      && (
                        LinkValueTrimmed.length != NATextLength
                        || LinkValueTrimmed != NAText
                      )
                    )
                    {
                      MethodFound = true;

                      LinkValue = null;
                      LinkValueTrimmed = null;
                    }
                  }

                if (MethodFound == false)
                  if (DKsMirrorColumn != -1)
                    if (
                      RowDisplayValues[DKsMirrorColumn] != null
                      && RowDisplayValues[DKsMirrorColumn].length > 0
                    )
                    {
                      MutualMethodsObject.CheckDKsMirrorB(RowDisplayValues);

                      if (
                        DKsMirrorValueTrimmed != null
                        && DKsMirrorValueTrimmed.length > 0
                        && (
                          DKsMirrorValueTrimmed.length != NATextLength
                          || DKsMirrorValueTrimmed != NAText
                        )
                      )
                      {
                        MethodFound = true;

                        DKsMirrorValue = null;
                        DKsMirrorValueTrimmed = null;
                      }
                    }

                if (MethodFound == false)
                  if (UpdateLinkColumn != -1)
                    if (
                      RowDisplayValues[UpdateLinkColumn] != null
                      && RowDisplayValues[UpdateLinkColumn].length > 0
                    )
                    {
                      MutualMethodsObject.CheckUpdateLinkB(RowDisplayValues);

                      if (
                        UpdateLinkValueTrimmed != null
                        && UpdateLinkValueTrimmed.length > 0
                        && (
                          UpdateLinkValueTrimmed.length != NATextLength
                          || UpdateLinkValueTrimmed != NAText
                        )
                      )
                      {
                        MethodFound = true;

                        UpdateLinkValue = null;
                        UpdateLinkValueTrimmed = null;
                      }
                    }
              }
            }

            // checking if there's anything to apply to the current row, or that it needs to be announced
            // note:
            // doing it since spreadsheet commands are slow
            // , so it's faster to do this because trying to do as least as possible spreadsheet commands

            if (MethodFound == true)
            {
              MethodFound = false;

              if (
                AnnouncedValue != null
                && GeneralToolsObject.VariableIsString(AnnouncedValue) == true
                && (
                  AnnouncedValue.length == 2
                  && AnnouncedValue == "No"
                  || AnnouncedValue.length == NATextLength
                  && AnnouncedValue == NAText
                )
              )
              {
                // checking if the current row needs to be announced
                if (
                  AnnouncedValue.length == 2
                  && AnnouncedValue == "No"
                )
                  if (MethodFound == false)
                    MethodFound = true;

                // fixes haven't been checked if entered here, so make sure to check
                MutualMethodsObject.CheckFixes(RowDisplayValues);
              }
                  
              if (MutualMethodsObject.CheckInfo(RowDisplayValues, true) == true)
                if (MethodFound == false)
                  MethodFound = true;

              MutualMethodsObject.CheckTitleID(RowDisplayValues, false);
              MutualMethodsObject.CheckRegion(RowDisplayValues, false);
              MutualMethodsObject.CheckGenre(RowDisplayValues, false);
              MutualMethodsObject.CheckDLC(RowDisplayValues, false);
              MutualMethodsObject.CheckDate(RowDisplayValues, false);
              
              MutualMethodsObject.CheckBaseSizeB(RowDisplayValues);
              MutualMethodsObject.CheckVersion(RowDisplayValues);
              MutualMethodsObject.CheckBP(RowDisplayValues);
              MutualMethodsObject.CheckUploader(RowDisplayValues);
              MutualMethodsObject.CheckLinkB(RowDisplayValues);
              MutualMethodsObject.CheckTested(RowDisplayValues);
              MutualMethodsObject.CheckDKsMirrorB(RowDisplayValues);
              MutualMethodsObject.CheckUpdateLinkB(RowDisplayValues);
              MutualMethodsObject.CheckSubmitter(RowDisplayValues);

              // checking fixes start

              ColumnsFixes = [];

              ////SpreadsheetToolsObject.UnprotectSheet();

              if (
                RegionColumn != -1
                && RowDisplayValues[RegionColumn] != null
              )
                if (
                  RegionValueTrimmed != null
                  && RegionValueTrimmed != RowDisplayValues[RegionColumn]
                )
                  if (
                    RowDisplayValues[RegionColumn].length == 0
                    || RegionValueTrimmed != null
                    && RegionValueTrimmed.length == NATextLength
                    && RegionValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[RegionColumn]) == true
                  )
                  {
                    if (MethodFound == false)
                      MethodFound = true;

                    ColumnsFixes.push(RegionColumnName);
                    ColumnsFixesAmount++;

                    MutualMethodsObject.ApplyRegion(RealRowsIndex, RegionValueTrimmed, RowDisplayValues);
                  }
                  else
                  {
                    RegionValue = null;
                    RegionValueTrimmed = null;
                    
                    if (MutualMethodsObject.CheckRegion(RowDisplayValues, true) == true)
                      if (MethodFound == false)
                        MethodFound = true;
                  }

              if (
                GenreColumn != -1
                && RowDisplayValues[GenreColumn] != null
              )
                if (
                  GenreValueTrimmed != null
                  && GenreValueTrimmed != RowDisplayValues[GenreColumn]
                )
                  if (
                    RowDisplayValues[GenreColumn].length == 0
                    || GenreValueTrimmed != null
                    && GenreValueTrimmed.length == NATextLength
                    && GenreValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[GenreColumn]) == true
                  )
                  {
                    if (MethodFound == false)
                      MethodFound = true;

                    ColumnsFixes.push(GenreColumnName);
                    ColumnsFixesAmount++;

                    MutualMethodsObject.ApplyGenre(RealRowsIndex, GenreValueTrimmed, RowDisplayValues);
                  }
                  else
                  {
                    GenreValue = null;
                    GenreValueTrimmed = null;
                    
                    if (MutualMethodsObject.CheckGenre(RowDisplayValues, true) == true)
                      if (MethodFound == false)
                        MethodFound = true;
                  }

              if (
                DLCColumn != -1
                && RowDisplayValues[DLCColumn] != null
              )
                if (
                  DLCValueTrimmed != null
                  && DLCValueTrimmed != RowDisplayValues[DLCColumn]
                )
                  if (
                    RowDisplayValues[DLCColumn].length == 0
                    || DLCValueTrimmed != null
                    && DLCValueTrimmed.length == NATextLength
                    && DLCValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[DLCColumn]) == true
                  )
                  {
                    if (MethodFound == false)
                      MethodFound = true;

                    ColumnsFixes.push(DLCColumnName);
                    ColumnsFixesAmount++;

                    MutualMethodsObject.ApplyDLC(RealRowsIndex, DLCValueTrimmed, RowDisplayValues);
                  }
                  else
                  {
                    DLCValue = null;
                    DLCValueTrimmed = null;
                    
                    if (MutualMethodsObject.CheckDLC(RowDisplayValues, true) == true)
                      if (MethodFound == false)
                        MethodFound = true;
                  }

              if (
                BaseSizeColumn != -1
                && RowDisplayValues[BaseSizeColumn] != null
              )
                if (
                  BaseSizeValueTrimmed != null
                  && BaseSizeValueTrimmed != RowDisplayValues[BaseSizeColumn]
                )
                {
                  if (MethodFound == false)
                    MethodFound = true;

                  if (
                    RowDisplayValues[BaseSizeColumn].length == 0
                    || BaseSizeValueTrimmed.length == NATextLength
                    && BaseSizeValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[BaseSizeColumn]) == true
                  )
                  {
                    ColumnsFixes.push(BaseSizeColumnName);
                    ColumnsFixesAmount++;
                  }

                  MutualMethodsObject.ApplyBaseSizeB(RealRowsIndex, BaseSizeValueTrimmed, RowRealValues);
                }

              if (
                VersionColumn != -1
                && RowDisplayValues[VersionColumn] != null
              )
                if (
                  VersionValueTrimmed != null
                  && VersionValueTrimmed != RowDisplayValues[VersionColumn]
                )
                {
                  if (MethodFound == false)
                    MethodFound = true;

                  if (
                    RowDisplayValues[VersionColumn].length == 0
                    || GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                    && VersionValueTrimmed.length == NATextLength
                    && VersionValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[VersionColumn]) == true
                  )
                  {
                    ColumnsFixes.push(VersionColumnName);
                    ColumnsFixesAmount++;
                  }
                  
                  MutualMethodsObject.ApplyVersionB(RealRowsIndex, VersionValueTrimmed, RowRealValues);
                }

              if (
                UpdateSizeColumn != -1
                && RowDisplayValues[UpdateSizeColumn] != null
              )
                if (
                  VersionValueTrimmed != null
                  && GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == true
                  && (
                    GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                    ? Number(VersionValueTrimmed)
                    : VersionValueTrimmed
                  ) == 1
                )
                {
                  if (
                    (
                      NoUpdateSizeText != undefined
                      && NoUpdateSizeText != null
                      && GeneralToolsObject.VariableIsString(NoUpdateSizeText) == true
                      && NoUpdateSizeText.length > 0
                      ? NoUpdateSizeText
                      : NAText
                    ) != RowDisplayValues[UpdateSizeColumn]
                  )
                    if (MethodFound == false)
                      MethodFound = true;

                  MutualMethodsObject.ApplyUpdateSizeB(
                    RealRowsIndex
                    , (
                      NoUpdateSizeText != undefined
                      && NoUpdateSizeText != null
                      && GeneralToolsObject.VariableIsString(NoUpdateSizeText) == true
                      && NoUpdateSizeText.length > 0
                      ? NoUpdateSizeText
                      : NAText
                    )
                    , RowRealValues
                  );
                }
                else
                {
                  MutualMethodsObject.CheckUpdateSizeB(RowDisplayValues);

                  if (
                    UpdateSizeValueTrimmed != null
                    && UpdateSizeValueTrimmed != RowDisplayValues[UpdateSizeColumn]
                  )
                  {
                    if (MethodFound == false)
                      MethodFound = true;

                    if (
                      RowDisplayValues[UpdateSizeColumn].length == 0
                      || UpdateSizeValueTrimmed.length == NATextLength
                      && UpdateSizeValueTrimmed == NAText
                      && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[UpdateSizeColumn]) == true
                    )
                    {
                      ColumnsFixes.push(UpdateSizeColumnName);
                      ColumnsFixesAmount++;
                    }
                    
                    MutualMethodsObject.ApplyUpdateSizeB(RealRowsIndex, UpdateSizeValueTrimmed, RowRealValues);
                  }
                }

              if (
                UploaderColumn != -1
                && RowDisplayValues[UploaderColumn] != null
              )
                if (
                  UploaderValueTrimmed != null
                  && UploaderValueTrimmed != RowDisplayValues[UploaderColumn]
                )
                {
                  if (
                    RowDisplayValues[UploaderColumn].length == 0
                    || UploaderValueTrimmed.length == NATextLength
                    && UploaderValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[UploaderColumn]) == true
                  )
                  {
                    ColumnsFixes.push(UploaderColumnName);
                    ColumnsFixesAmount++;
                  }
                  
                  // do not trim the uploader in its cell
                  if (
                    UploaderValue == null
                    || GeneralToolsObject.VariableIsString(UploaderValue) == false
                    || UploaderValue.length == 0
                    || UploaderValue.length == UploaderValueTrimmed.length
                    && UploaderValue == UploaderValueTrimmed
                    || UploaderValueTrimmed.length == NATextLength
                    && UploaderValueTrimmed == NAText
                  )
                  {
                    if (MethodFound == false)
                      MethodFound = true;

                    MutualMethodsObject.ApplyUploader(RealRowsIndex, UploaderValueTrimmed, RowRealValues);
                  }
                }

              if (
                SubmitterColumn != -1
                && RowDisplayValues[SubmitterColumn] != null
              )
                if (
                  SubmitterValueTrimmed != null
                  && SubmitterValueTrimmed != RowDisplayValues[SubmitterColumn]
                )
                {
                  if (
                    RowDisplayValues[SubmitterColumn].length == 0
                    || SubmitterValueTrimmed.length == NATextLength
                    && SubmitterValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[SubmitterColumn]) == true
                  )
                  {
                    ColumnsFixes.push(SubmitterColumnName);
                    ColumnsFixesAmount++;
                  }
                  
                  // do not trim the submitter in its cell
                  if (
                    SubmitterValue == null
                    || GeneralToolsObject.VariableIsString(SubmitterValue) == false
                    || SubmitterValue.length == 0
                    || SubmitterValue.length == SubmitterValueTrimmed.length
                    && SubmitterValue == SubmitterValueTrimmed
                    || SubmitterValueTrimmed.length == NATextLength
                    && SubmitterValueTrimmed == NAText
                  )
                  {
                    if (MethodFound == false)
                      MethodFound = true;

                    MutualMethodsObject.ApplySubmitter(RealRowsIndex, SubmitterValueTrimmed, RowRealValues);
                  }
                }

              ////SpreadsheetToolsObject.ProtectSheet();

              // checking fixes end

              // checking existing fixes start
              
              if (
                FixesValueTrimmed != null
                && FixesValueTrimmed.length > 0
                && (
                  FixesValueTrimmed.length != 3
                  || FixesValueTrimmed != "Fix"
                )
                && (
                  FixesValueTrimmed.length != 4
                  || FixesValueTrimmed != "None"
                )
              )
              {
                FixesValueTrimmedSplitted = FixesValueTrimmed.split(", ");

                if (FixesValueTrimmedSplitted != null)
                {
                  FixesValueTrimmedSplittedAmount = FixesValueTrimmedSplitted.length;

                  if (FixesValueTrimmedSplittedAmount > 0)
                  {
                    FixesValueTrimmedSplittedIndex = 0;

                    while (FixesValueTrimmedSplittedIndex < FixesValueTrimmedSplittedAmount)
                    {
                      FixesValueTrimmedSplittedCell = FixesValueTrimmedSplitted[FixesValueTrimmedSplittedIndex];
                      FixesValueTrimmedSplittedCellValue = null;

                      switch (FixesValueTrimmedSplittedCell)
                      {
                        case TitleIDColumnName:
                          if (TitleIDColumn != -1)
                            FixesValueTrimmedSplittedCellValue = TitleIDValueTrimmed;

                          break;

                        case RegionColumnName:
                          if (RegionColumn != -1)
                            FixesValueTrimmedSplittedCellValue = RegionValueTrimmed;

                          break;

                        case GenreColumnName:
                          if (GenreColumn != -1)
                            FixesValueTrimmedSplittedCellValue = GenreValueTrimmed;

                          break;

                        case BaseSizeColumnName:
                          if (BaseSizeColumn != -1)
                            FixesValueTrimmedSplittedCellValue = BaseSizeValueTrimmed;

                          break;

                        case VersionColumnName:
                          if (VersionColumn != -1)
                            FixesValueTrimmedSplittedCellValue = VersionValueTrimmed;

                          break;

                        case UpdateSizeColumnName:
                          if (UpdateSizeColumn != -1)
                            FixesValueTrimmedSplittedCellValue = UpdateSizeValueTrimmed;

                          break;

                        case DLCColumnName:
                          if (DLCColumn != -1)
                            FixesValueTrimmedSplittedCellValue = DLCValueTrimmed;

                          break;

                        case BPColumnName:
                          if (BPColumn != -1)
                            FixesValueTrimmedSplittedCellValue = BPValueTrimmed;

                          break;

                        case UploaderColumnName:
                          if (UploaderColumn != -1)
                            FixesValueTrimmedSplittedCellValue = UploaderValueTrimmed;

                          break;

                        case DateColumnName:
                          if (DateColumn != -1)
                            FixesValueTrimmedSplittedCellValue = DateValueTrimmed;

                          break;

                        case LinkColumnName:
                          if (LinkColumn != -1)
                            FixesValueTrimmedSplittedCellValue = LinkValueTrimmed;

                          break;

                        case TestedColumnName:
                          if (TestedColumn != -1)
                            FixesValueTrimmedSplittedCellValue = TestedValueTrimmed;

                          break;

                        case DKsMirrorColumnName:
                          if (DKsMirrorColumn != -1)
                            FixesValueTrimmedSplittedCellValue = DKsMirrorValueTrimmed;

                          break;

                        case UpdateLinkColumnName:
                          if (UpdateLinkColumn != -1)
                            FixesValueTrimmedSplittedCellValue = UpdateLinkValueTrimmed;

                          break;

                        case SubmitterColumnName:
                          if (SubmitterColumn != -1)
                            FixesValueTrimmedSplittedCellValue = SubmitterValueTrimmed;

                          break;

                        default:
                          if (
                            AllowCustomFixes == undefined
                            || AllowCustomFixes == null
                            || GeneralToolsObject.VariableIsBoolean(AllowCustomFixes) == false
                            || AllowCustomFixes == false
                            || GeneralToolsObject.VariableIsNotNA(FixesValueTrimmedSplittedCell) == false
                            || NoFixesText != undefined
                            && NoFixesText != null
                            && GeneralToolsObject.VariableIsString(NoFixesText) == true
                            && NoFixesText.length > 0
                            && FixesValueTrimmedSplittedCell.toUpperCase() == NoFixesText.toUpperCase()
                          )
                            FixesValueTrimmedSplittedCellValueInvalid = true;
                      }

                      if (
                        FixesValueTrimmedSplittedCellValueInvalid == true
                        || FixesValueTrimmedSplittedCellValue != null
                        && GeneralToolsObject.VariableIsString(FixesValueTrimmedSplittedCellValue) == true
                        && FixesValueTrimmedSplittedCellValue.length > 0
                        && (
                          FixesValueTrimmedSplittedCellValue.length != NATextLength
                          || FixesValueTrimmedSplittedCellValue != NAText
                        )
                      )
                      {
                        FixesValueTrimmedSplitted.splice(FixesValueTrimmedSplittedIndex, 1);
                        FixesValueTrimmedSplittedAmount--;

                        if (FixesValueTrimmedSplittedCellValueInvalid == true)
                          FixesValueTrimmedSplittedCellValueInvalid = false;
                        // if found 1 link then all the others aren't needed anymore
                        else if (
                          FixesValueTrimmedSplittedCell == LinkColumnName
                          || FixesValueTrimmedSplittedCell == DKsMirrorColumnName
                          || FixesValueTrimmedSplittedCell == UpdateLinkColumnName
                        )
                        {
                          FixesValueTrimmedSplittedLinkIndex =
                            FixesValueTrimmedSplitted.indexOf(LinkColumnName);

                          if (FixesValueTrimmedSplittedLinkIndex != -1)
                          {
                            FixesValueTrimmedSplitted.splice(FixesValueTrimmedSplittedLinkIndex, 1);

                            if (FixesValueTrimmedSplittedLinkIndex >= FixesValueTrimmedSplittedIndex)
                              FixesValueTrimmedSplittedAmount--;
                            else
                              FixesValueTrimmedSplittedIndex--;

                            FixesValueTrimmedSplittedLinkIndex = -1;
                          }

                          FixesValueTrimmedSplittedDKsMirrorIndex =
                            FixesValueTrimmedSplitted.indexOf(DKsMirrorColumnName);

                          if (FixesValueTrimmedSplittedDKsMirrorIndex != -1)
                          {
                            FixesValueTrimmedSplitted.splice(FixesValueTrimmedSplittedDKsMirrorIndex, 1);

                            if (FixesValueTrimmedSplittedDKsMirrorIndex >= FixesValueTrimmedSplittedIndex)
                              FixesValueTrimmedSplittedAmount--;
                            else
                              FixesValueTrimmedSplittedIndex--;

                            FixesValueTrimmedSplittedDKsMirrorIndex = -1;
                          }

                          FixesValueTrimmedSplittedUpdateLinkIndex =
                            FixesValueTrimmedSplitted.indexOf(UpdateLinkColumnName);

                          if (FixesValueTrimmedSplittedUpdateLinkIndex != -1)
                          {
                            FixesValueTrimmedSplitted.splice(FixesValueTrimmedSplittedUpdateLinkIndex, 1);

                            if (FixesValueTrimmedSplittedUpdateLinkIndex >= FixesValueTrimmedSplittedIndex)
                              FixesValueTrimmedSplittedAmount--;
                            else
                              FixesValueTrimmedSplittedIndex--;

                            FixesValueTrimmedSplittedUpdateLinkIndex = -1;
                          }
                        }
                      }
                      else
                        FixesValueTrimmedSplittedIndex++;

                      FixesValueTrimmedSplittedCell = null;

                      if (FixesValueTrimmedSplittedCellValue != null)
                        FixesValueTrimmedSplittedCellValue = null;
                    }

                    if (FixesValueTrimmedSplittedIndex > 0)
                      FixesValueTrimmedSplittedIndex = 0;

                    if (FixesValueTrimmedSplittedAmount > 0)
                    {
                      ColumnsFixes =
                        ColumnsFixesAmount > 0
                        ? ColumnsFixes.concat(FixesValueTrimmedSplitted)
                        : FixesValueTrimmedSplitted;

                      ColumnsFixesAmount =
                        ColumnsFixesAmount > 0
                        ? ColumnsFixesAmount + FixesValueTrimmedSplittedAmount
                        : FixesValueTrimmedSplittedAmount;

                      FixesValueTrimmedSplittedAmount = 0;
                    }
                  }

                  FixesValueTrimmedSplitted = null;
                }
              }

              // checking existing fixes end

              // additional fixes start

              ColumnsAdditionalFixes = [];

              if (
                BaseSizeValueTrimmed != null
                && BaseSizeValueTrimmed.length > 0
                && (
                  BaseSizeValueTrimmed.length != NATextLength
                  || BaseSizeValueTrimmed != NAText
                )
                || UpdateSizeValueTrimmed != null
                && UpdateSizeValueTrimmed.length > 0
                && (
                  UpdateSizeValueTrimmed.length != NATextLength
                  || UpdateSizeValueTrimmed != NAText
                )
                //&& (
                //  NoUpdateSizeText == undefined
                //  || NoUpdateSizeText == null
                //  || GeneralToolsObject.VariableIsString(NoUpdateSizeText) == false
                //  || UpdateSizeValueTrimmed.length != NoUpdateSizeText.length
                //  || UpdateSizeValueTrimmed != NoUpdateSizeText
                //)
                || VersionValueTrimmed != null
                && GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == true
                && (
                  GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                  ? Number(VersionValueTrimmed)
                  : VersionValueTrimmed
                ) >= 1
              )
              {
                if (
                  RegionColumn != -1
                  && (
                    RegionValueTrimmed == null
                    || RegionValueTrimmed.length == 0
                    || RegionValueTrimmed.length == NATextLength
                    && RegionValueTrimmed == NAText
                  )
                )
                {
                  ColumnsAdditionalFixes.push(RegionColumnName);
                  ColumnsAdditionalFixesAmount++;
                }

                if (
                  GenreColumn != -1
                  && (
                    GenreValueTrimmed == null
                    || GenreValueTrimmed.length == 0
                    || GenreValueTrimmed.length == NATextLength
                    && GenreValueTrimmed == NAText
                  )
                )
                {
                  ColumnsAdditionalFixes.push(GenreColumnName);
                  ColumnsAdditionalFixesAmount++;
                }
              }

              if (
                VersionValueTrimmed != null
                && GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == true
                && (
                  GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                  ? Number(VersionValueTrimmed)
                  : VersionValueTrimmed
                ) >= 1
              )
                if (
                  BaseSizeColumn != -1
                  && (
                    BaseSizeValueTrimmed == null
                    || BaseSizeValueTrimmed.length == 0
                    || BaseSizeValueTrimmed.length == NATextLength
                    && BaseSizeValueTrimmed == NAText
                  )
                )
                {
                  ColumnsAdditionalFixes.push(BaseSizeColumnName);
                  ColumnsAdditionalFixesAmount++;
                }

              if (
                BaseSizeValueTrimmed != null
                && BaseSizeValueTrimmed.length > 0
                && (
                  BaseSizeValueTrimmed.length != NATextLength
                  || BaseSizeValueTrimmed != NAText
                )
              )
                if (
                  VersionColumn != -1
                  && (
                    VersionValueTrimmed == null
                    || GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                    && (
                      VersionValueTrimmed.length == 0
                      || VersionValueTrimmed.length == NATextLength
                      && VersionValueTrimmed == NAText
                    )
                  )
                )
                {
                  ColumnsAdditionalFixes.push(VersionColumnName);
                  ColumnsAdditionalFixesAmount++;
                }

              if (
                VersionValueTrimmed != null
                && GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == true
                && (
                  GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                  ? Number(VersionValueTrimmed)
                  : VersionValueTrimmed
                ) > 1
              )
                if (
                  UpdateSizeColumn != -1
                  && (
                    UpdateSizeValueTrimmed == null
                    || UpdateSizeValueTrimmed.length == 0
                    || UpdateSizeValueTrimmed.length == NATextLength
                    && UpdateSizeValueTrimmed == NAText
                    || NoUpdateSizeText != undefined
                    && NoUpdateSizeText != null
                    && GeneralToolsObject.VariableIsString(NoUpdateSizeText) == true
                    && UpdateSizeValueTrimmed.length == NoUpdateSizeText.length
                    && UpdateSizeValueTrimmed == NoUpdateSizeText
                  )
                )
                {
                  ColumnsAdditionalFixes.push(UpdateSizeColumnName);
                  ColumnsAdditionalFixesAmount++;
                }

              if (
                UpdateSizeValueTrimmed != null
                && UpdateSizeValueTrimmed.length > 0
                && (
                  UpdateSizeValueTrimmed.length != NATextLength
                  || UpdateSizeValueTrimmed != NAText
                )
              )
                if (
                  VersionColumn != -1
                  && (
                    VersionValueTrimmed == null
                    || GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                    && (
                      VersionValueTrimmed.length == 0
                      || VersionValueTrimmed.length == NATextLength
                      && VersionValueTrimmed == NAText
                    )
                    || (
                      NoUpdateSizeText != undefined
                      && NoUpdateSizeText != null
                      && GeneralToolsObject.VariableIsString(NoUpdateSizeText) == true
                      && UpdateSizeValueTrimmed.length == NoUpdateSizeText.length
                      && UpdateSizeValueTrimmed == NoUpdateSizeText
                    )
                    && (
                      GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == false
                      || (
                        GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                        ? Number(VersionValueTrimmed)
                        : VersionValueTrimmed
                      ) != 1
                    )
                    || (
                      NoUpdateSizeText == undefined
                      || NoUpdateSizeText == null
                      || GeneralToolsObject.VariableIsString(NoUpdateSizeText) == false
                      || UpdateSizeValueTrimmed.length != NoUpdateSizeText.length
                      || UpdateSizeValueTrimmed != NoUpdateSizeText
                    )
                    && (
                      GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == false
                      || (
                        GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                        ? Number(VersionValueTrimmed)
                        : VersionValueTrimmed
                      ) <= 1
                    )
                  )
                )
                {
                  ColumnsAdditionalFixes.push(VersionColumnName);
                  ColumnsAdditionalFixesAmount++;
                }

              if (
                BPColumn != -1
                && (
                  BPValueTrimmed == null
                  || BPValueTrimmed.length == 0
                  || BPValueTrimmed.length == NATextLength
                  && BPValueTrimmed == NAText
                )
              )
              {
                ColumnsAdditionalFixes.push(BPColumnName);
                ColumnsAdditionalFixesAmount++;
              }

              if (
                DateColumn != -1
                && (
                  DateValueTrimmed == null
                  || DateValueTrimmed.length == 0
                  || DateValueTrimmed.length == NATextLength
                  && DateValueTrimmed == NAText
                )
              )
              {
                ColumnsAdditionalFixes.push(DateColumnName);
                ColumnsAdditionalFixesAmount++;
              }

              if (
                TestedColumn != -1
                && (
                  TestedValueTrimmed == null
                  || TestedValueTrimmed.length == 0
                  || TestedValueTrimmed.length == NATextLength
                  && TestedValueTrimmed == NAText
                )
              )
              {
                ColumnsAdditionalFixes.push(TestedColumnName);
                ColumnsAdditionalFixesAmount++;
              }

              if (ColumnsAdditionalFixesAmount > 0)
              {
                ColumnsFixes =
                  ColumnsFixesAmount > 0
                  ? ColumnsFixes.concat(ColumnsAdditionalFixes)
                  : ColumnsAdditionalFixes;

                ColumnsFixesAmount =
                  ColumnsFixesAmount > 0
                  ? ColumnsFixesAmount + ColumnsAdditionalFixesAmount
                  : ColumnsAdditionalFixesAmount;

                ColumnsAdditionalFixes = null;
                ColumnsAdditionalFixesAmount = 0;
              }

              // additional fixes end

              // urgent fixes start

              ColumnsUrgentFixes = [];

              ////SpreadsheetToolsObject.UnprotectSheet();

              if (
                TitleIDColumn != -1
                && RowDisplayValues[TitleIDColumn] != null
              )
                if (
                  TitleIDValueTrimmed != null
                  && TitleIDValueTrimmed != RowDisplayValues[TitleIDColumn]
                )
                  if (
                    RowDisplayValues[TitleIDColumn].length == 0
                    || TitleIDValueTrimmed.length == NATextLength
                    && TitleIDValueTrimmed == NAText
                    && GeneralToolsObject.VariableIsNotNA(RowDisplayValues[TitleIDColumn]) == true
                  )
                  {
                    if (MethodFound == false)
                      MethodFound = true;

                    ColumnsUrgentFixes.push(TitleIDColumnName);
                    ColumnsUrgentFixesAmount++;

                    MutualMethodsObject.ApplyTitleID(RealRowsIndex, TitleIDValueTrimmed, RowDisplayValues);
                  }
                  else
                  {
                    TitleIDValue = null;
                    TitleIDValueTrimmed = null;
                    
                    if (MutualMethodsObject.CheckTitleID(RowDisplayValues, true) == true)
                      if (MethodFound == false)
                        MethodFound = true;
                  }

              ////SpreadsheetToolsObject.ProtectSheet();
              
              if (
                (
                  LinkColumn == -1
                  || RowDisplayValues[LinkColumn] == null
                  || LinkValueTrimmed == null
                  || LinkValueTrimmed.length == 0
                  || LinkValueTrimmed.length == NATextLength
                  && LinkValueTrimmed == NAText
                )
                && (
                  DKsMirrorColumn == -1
                  || RowDisplayValues[DKsMirrorColumn] == null
                  || DKsMirrorValueTrimmed == null
                  || DKsMirrorValueTrimmed.length == 0
                  || DKsMirrorValueTrimmed.length == NATextLength
                  && DKsMirrorValueTrimmed == NAText
                )
                && (
                  UpdateLinkColumn == -1
                  || RowDisplayValues[UpdateLinkColumn] == null
                  || UpdateLinkValueTrimmed == null
                  || UpdateLinkValueTrimmed.length == 0
                  || UpdateLinkValueTrimmed.length == NATextLength
                  && UpdateLinkValueTrimmed == NAText
                )
              )
              {
                if (LinkColumn != -1)
                {
                  ColumnsUrgentFixes.push(LinkColumnName);
                  ColumnsUrgentFixesAmount++;
                }

                if (DKsMirrorColumn != -1)
                {
                  ColumnsUrgentFixes.push(DKsMirrorColumnName);
                  ColumnsUrgentFixesAmount++;
                }

                if (UpdateLinkColumn != -1)
                {
                  ColumnsUrgentFixes.push(UpdateLinkColumnName);
                  ColumnsUrgentFixesAmount++;
                }
              }

              if (ColumnsUrgentFixesAmount > 0)
              {
                ColumnsFixes =
                  ColumnsFixesAmount > 0
                  ? ColumnsFixes.concat(ColumnsUrgentFixes)
                  : ColumnsUrgentFixes;

                ColumnsFixesAmount =
                  ColumnsFixesAmount > 0
                  ? ColumnsFixesAmount + ColumnsUrgentFixesAmount
                  : ColumnsUrgentFixesAmount;

                ColumnsUrgentFixes = null;
                ColumnsUrgentFixesAmount = 0;
              }

              // urgent fixes end

              // getting the sorted unique columns fixes and its text start

              if (ColumnsFixesAmount > 0)
              {
                ColumnsFixes = ColumnsFixes.unique().FixesSort();
                ColumnsFixesAmount = ColumnsFixes.length;
              }

              ColumnsFixesText =
                ColumnsFixesAmount > 0
                ? (
                  ColumnsFixesAmount > 1
                  ? ColumnsFixes.join(", ")
                  : ColumnsFixes[0]
                )
                : (
                  NoFixesText != undefined
                  && NoFixesText != null
                  && GeneralToolsObject.VariableIsString(NoFixesText) == true
                  && NoFixesText.length > 0
                  ? NoFixesText
                  : NAText
                );

              // getting the sorted unique columns fixes and its text end

              if (
                ColumnsFixesText != null
                && GeneralToolsObject.VariableIsString(ColumnsFixesText) == true
                && ColumnsFixesText.length > 0
                && (
                  FixesValueTrimmed == null
                  || GeneralToolsObject.VariableIsString(FixesValueTrimmed) == false
                  || FixesValueTrimmed.length == 0
                  || ColumnsFixesText != FixesValueTrimmed
                )
              )
                if (MethodFound == false)
                  MethodFound = true;
            }

            // verifying that no rows have been changed

            if (MethodFound == true)
            {
              MethodFound = false;

              if (
                (
                  TitleIDValueTrimmed != null
                  && TitleIDValueTrimmed.length > 0
                  ? MutualMethodsObject.VerifyTitleID(RealRowsIndex, RowRealValues)
                  : MutualMethodsObject.VerifyTitle(RealRowsIndex, RowRealValues)
                ) == true
              )
                MethodFound = true;
              else
                if (dry_run == false)
                  // recheck rows if verification failed
                  if (RecheckRows == false)
                    RecheckRows = true;
            }

            // if found something to be applied to the current row or that it needs to be announced then enter

            if (MethodFound == true)
            {
              MethodFound = false;

              if (dry_run == false)
                // recheck rows after every change in the spreadsheet
                if (RecheckRows == false)
                  RecheckRows = true;

              MutualMethodsObject.CheckLinkA(RealRowsIndex, RowDisplayValues);
              MutualMethodsObject.CheckDKsMirrorA(RealRowsIndex, RowDisplayValues);
              MutualMethodsObject.CheckUpdateLinkA(RealRowsIndex, RowDisplayValues);

              // urgent fixes again start

              ColumnsUrgentFixes = [];
              
              if (
                (
                  LinkColumn == -1
                  || RowDisplayValues[LinkColumn] == null
                  || LinkValueTrimmed == null
                  || LinkValueTrimmed.length == 0
                  || LinkValueTrimmed.length == NATextLength
                  && LinkValueTrimmed == NAText
                )
                && (
                  DKsMirrorColumn == -1
                  || RowDisplayValues[DKsMirrorColumn] == null
                  || DKsMirrorValueTrimmed == null
                  || DKsMirrorValueTrimmed.length == 0
                  || DKsMirrorValueTrimmed.length == NATextLength
                  && DKsMirrorValueTrimmed == NAText
                )
                && (
                  UpdateLinkColumn == -1
                  || RowDisplayValues[UpdateLinkColumn] == null
                  || UpdateLinkValueTrimmed == null
                  || UpdateLinkValueTrimmed.length == 0
                  || UpdateLinkValueTrimmed.length == NATextLength
                  && UpdateLinkValueTrimmed == NAText
                )
              )
              {
                if (LinkColumn != -1)
                {
                  ColumnsUrgentFixes.push(LinkColumnName);
                  ColumnsUrgentFixesAmount++;
                }

                if (DKsMirrorColumn != -1)
                {
                  ColumnsUrgentFixes.push(DKsMirrorColumnName);
                  ColumnsUrgentFixesAmount++;
                }

                if (UpdateLinkColumn != -1)
                {
                  ColumnsUrgentFixes.push(UpdateLinkColumnName);
                  ColumnsUrgentFixesAmount++;
                }
              }

              if (ColumnsUrgentFixesAmount > 0)
              {
                ColumnsFixes =
                  ColumnsFixesAmount > 0
                  ? ColumnsFixes.concat(ColumnsUrgentFixes)
                  : ColumnsUrgentFixes;

                ColumnsFixesAmount =
                  ColumnsFixesAmount > 0
                  ? ColumnsFixesAmount + ColumnsUrgentFixesAmount
                  : ColumnsUrgentFixesAmount;

                // getting the sorted unique columns fixes and its text again start

                ColumnsFixes = ColumnsFixes.unique().FixesSort();
                ColumnsFixesAmount = ColumnsFixes.length;

                ColumnsFixesText =
                  ColumnsFixesAmount > 1
                  ? ColumnsFixes.join(", ")
                  : ColumnsFixes[0];

                // getting the sorted unique columns fixes and its text again end

                ColumnsUrgentFixes = null;
                ColumnsUrgentFixesAmount = 0;
              }
              else if (
                IgnoreTitleIDRequiredForAnnouncing == true
                || TitleIDColumn == -1
                || TitleIDValueTrimmed != null
                && TitleIDValueTrimmed.length > 0
                && (
                  TitleIDValueTrimmed.length != NATextLength
                  || TitleIDValueTrimmed != NAText
                )
              )
                FoundMatch = true;

              // urgent fixes again end

              // applying values start

              ////SpreadsheetToolsObject.UnprotectSheet();

              FileSizeNumberWithFormat =
                MutualMethodsObject.CheckFileSizeNumberWithFormat(
                  RowRealValues
                  , RowDisplayValues
                  , BaseSizeColumn
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
                  FileSizeDoubleNumber = FileSizeNumberWithFormat[0];
                  FileSizeNumberFormat = FileSizeNumberWithFormat[1];

                  MutualMethodsObject.ApplyBaseSizeA(
                    RealRowsIndex, FileSizeDoubleNumber, FileSizeNumberFormat, RowDisplayValues, RowRealValues
                  );

                  FileSizeDoubleNumber = null;
                  FileSizeNumberFormat = null;
                }

                FileSizeNumberWithFormat = null;
              }

              if (
                VersionValueTrimmed != null
                && GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == true
              )
                MutualMethodsObject.ApplyVersionA(RealRowsIndex, VersionValueTrimmed, RowRealValues);

              if (
                VersionValueTrimmed != null
                && GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == true
                && (
                  GeneralToolsObject.VariableIsString(VersionValueTrimmed) == true
                  ? Number(VersionValueTrimmed)
                  : VersionValueTrimmed
                ) == 1
              )
                MutualMethodsObject.ApplyUpdateSizeB(
                  RealRowsIndex
                  , (
                    NoUpdateSizeText != undefined
                    && NoUpdateSizeText != null
                    && GeneralToolsObject.VariableIsString(NoUpdateSizeText) == true
                    && NoUpdateSizeText.length > 0
                    ? NoUpdateSizeText
                    : NAText
                  )
                  , RowRealValues
                );
              else
              {
                FileSizeNumberWithFormat =
                  MutualMethodsObject.CheckFileSizeNumberWithFormat(
                    RowRealValues
                    , RowDisplayValues
                    , UpdateSizeColumn
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
                    FileSizeDoubleNumber = FileSizeNumberWithFormat[0];
                    FileSizeNumberFormat = FileSizeNumberWithFormat[1];

                    MutualMethodsObject.ApplyUpdateSizeA(
                      RealRowsIndex, FileSizeDoubleNumber, FileSizeNumberFormat, RowDisplayValues, RowRealValues
                    );

                    FileSizeDoubleNumber = null;
                    FileSizeNumberFormat = null;
                  }

                  FileSizeNumberWithFormat = null;
                }
              }

              if (
                LinkValueTrimmed != null
                && LinkValueTrimmed.length > 0
                // apply it even if the cell has the same link, for safety
              )
                MutualMethodsObject.ApplyLink(RealRowsIndex, LinkValueTrimmed, RowRealValues, RowFormulas);

              if (
                DKsMirrorValueTrimmed != null
                && DKsMirrorValueTrimmed.length > 0
                // apply it even if the cell has the same link, for safety
              )
                MutualMethodsObject.ApplyDKsMirror(RealRowsIndex, DKsMirrorValueTrimmed, RowRealValues, RowFormulas);

              if (
                UpdateLinkValueTrimmed != null
                && UpdateLinkValueTrimmed.length > 0
                // apply it even if the cell has the same link, for safety
              )
                MutualMethodsObject.ApplyUpdateLink(RealRowsIndex, UpdateLinkValueTrimmed, RowRealValues, RowFormulas);

              // do not announce if a match hasn't been found
              if (FoundMatch == true)
              {
                FoundMatch = false;

                if (
                  AnnouncedValue != null
                  && GeneralToolsObject.VariableIsString(AnnouncedValue) == true
                  && (
                    AnnouncedValue.length == 2
                    && AnnouncedValue == "No"
                    || AnnouncedValue.length == NATextLength
                    && AnnouncedValue == NAText
                  )
                )
                {
                  MethodFound = true;

                  if (QuietAnnounce == false)
                  {
                    SpreadsheetToolsObject.CheckSpreadsheetDate();

                    if (
                      SpreadsheetToolsObject.SpreadsheetDate != null
                      && SpreadsheetToolsObject.SpreadsheetDate.length > 0
                    )
                    {
                      MutualMethodsObject.ApplyDateA(
                        RealRowsIndex, SpreadsheetToolsObject.SpreadsheetDate, RowRealValues
                      );

                      // since applied date even if it was empty, based on the spreadsheet date
                      // , so it doesn't need to be fixed anymore, in case it's in the columns fixes

                      ColumnsFixesDateIndex = ColumnsFixes.indexOf(DateColumnName);

                      if (ColumnsFixesDateIndex != -1)
                      {
                        ColumnsFixes.splice(ColumnsFixesDateIndex, 1);
                        ColumnsFixesAmount--;

                        ColumnsFixesText =
                          ColumnsFixesAmount > 0
                          ? (
                            ColumnsFixesAmount > 1
                            ? ColumnsFixes.join(", ")
                            : ColumnsFixes[0]
                          )
                          : (
                            NoFixesText != undefined
                            && NoFixesText != null
                            && GeneralToolsObject.VariableIsString(NoFixesText) == true
                            && NoFixesText.length > 0
                            ? NoFixesText
                            : NAText
                          );

                        ColumnsFixesDateIndex = -1;
                      }
                    }
                  }

                  MutualMethodsObject.ApplyAnnounced(RealRowsIndex, "Yes", RowRealValues);
                }
              }
              else
                // set the announced as not available in case it hasn't been announced yet
                // , in order to mention that it can't get announced the way it is at the moment
                if (
                  AnnouncedValue != null
                  && GeneralToolsObject.VariableIsString(AnnouncedValue) == true
                  && AnnouncedValue.length == 2
                  && AnnouncedValue == "No"
                )
                  MutualMethodsObject.ApplyAnnounced(RealRowsIndex, NAText, RowRealValues);

              MutualMethodsObject.ApplyFixes(RealRowsIndex, ColumnsFixesText, RowRealValues);

              ////SpreadsheetToolsObject.ProtectSheet();

              // applying values end

              // loading the applied values from the row display values into the row real values start

              if (TitleIDColumn != -1)
                if (
                  RowDisplayValues[TitleIDColumn] != null
                  && GeneralToolsObject.VariableIsString(RowDisplayValues[TitleIDColumn]) == true
                  && RowDisplayValues[TitleIDColumn].length > 0
                  && (
                    RowRealValues[TitleIDColumn] == null
                    || GeneralToolsObject.VariableIsString(RowRealValues[TitleIDColumn]) == false
                    || RowRealValues[TitleIDColumn].length == 0
                    || RowDisplayValues[TitleIDColumn].length != RowRealValues[TitleIDColumn].length
                    || RowDisplayValues[TitleIDColumn] != RowRealValues[TitleIDColumn]
                  )
                )
                  RowRealValues[TitleIDColumn] = RowDisplayValues[TitleIDColumn];

              if (RegionColumn != -1)
                if (
                  RowDisplayValues[RegionColumn] != null
                  && GeneralToolsObject.VariableIsString(RowDisplayValues[RegionColumn]) == true
                  && RowDisplayValues[RegionColumn].length > 0
                  && (
                    RowRealValues[RegionColumn] == null
                    || GeneralToolsObject.VariableIsString(RowRealValues[RegionColumn]) == false
                    || RowRealValues[RegionColumn].length == 0
                    || RowDisplayValues[RegionColumn].length != RowRealValues[RegionColumn].length
                    || RowDisplayValues[RegionColumn] != RowRealValues[RegionColumn]
                  )
                )
                  RowRealValues[RegionColumn] = RowDisplayValues[RegionColumn];

              if (GenreColumn != -1)
                if (
                  RowDisplayValues[GenreColumn] != null
                  && GeneralToolsObject.VariableIsString(RowDisplayValues[GenreColumn]) == true
                  && RowDisplayValues[GenreColumn].length > 0
                  && (
                    RowRealValues[GenreColumn] == null
                    || GeneralToolsObject.VariableIsString(RowRealValues[GenreColumn]) == false
                    || RowRealValues[GenreColumn].length == 0
                    || RowDisplayValues[GenreColumn].length != RowRealValues[GenreColumn].length
                    || RowDisplayValues[GenreColumn] != RowRealValues[GenreColumn]
                  )
                )
                  RowRealValues[GenreColumn] = RowDisplayValues[GenreColumn];

              if (DLCColumn != -1)
                if (
                  RowDisplayValues[DLCColumn] != null
                  && GeneralToolsObject.VariableIsString(RowDisplayValues[DLCColumn]) == true
                  && RowDisplayValues[DLCColumn].length > 0
                  && (
                    RowRealValues[DLCColumn] == null
                    || GeneralToolsObject.VariableIsString(RowRealValues[DLCColumn]) == false
                    || RowRealValues[DLCColumn].length == 0
                    || RowDisplayValues[DLCColumn].length != RowRealValues[DLCColumn].length
                    || RowDisplayValues[DLCColumn] != RowRealValues[DLCColumn]
                  )
                )
                  RowRealValues[DLCColumn] = RowDisplayValues[DLCColumn];

              if (DateColumn != -1)
                if (
                  RowDisplayValues[DateColumn] != null
                  && GeneralToolsObject.VariableIsString(RowDisplayValues[DateColumn]) == true
                  && RowDisplayValues[DateColumn].length > 0
                  && (
                    RowRealValues[DateColumn] == null
                    || GeneralToolsObject.VariableIsString(RowRealValues[DateColumn]) == false
                    || RowRealValues[DateColumn].length == 0
                    || RowDisplayValues[DateColumn].length != RowRealValues[DateColumn].length
                    || RowDisplayValues[DateColumn] != RowRealValues[DateColumn]
                  )
                )
                  RowRealValues[DateColumn] = RowDisplayValues[DateColumn];

              if (InfoColumn != -1)
                if (
                  RowDisplayValues[InfoColumn] != null
                  && GeneralToolsObject.VariableIsString(RowDisplayValues[InfoColumn]) == true
                  && RowDisplayValues[InfoColumn].length > 0
                  && (
                    RowRealValues[InfoColumn] == null
                    || GeneralToolsObject.VariableIsString(RowRealValues[InfoColumn]) == false
                    || RowRealValues[InfoColumn].length == 0
                    || RowDisplayValues[InfoColumn].length != RowRealValues[InfoColumn].length
                    || RowDisplayValues[InfoColumn] != RowRealValues[InfoColumn]
                  )
                )
                  RowRealValues[InfoColumn] = RowDisplayValues[InfoColumn];

              // loading the applied values from the row display values into the row real values end

              //SpreadsheetToolsObject.UnprotectSheet();

              // RowRealValuesColumnsAmount and RowFormulasColumnsAmount are the same
              // , but differentiating them just for reference
              SpreadsheetToolsObject.ApplyRow(RealRowsIndex, 0, RowRealValuesColumnsAmount, 0, RowRealValues);
              SpreadsheetToolsObject.ApplyRow(RealRowsIndex, 0, RowFormulasColumnsAmount, 1, RowFormulas);

              //SpreadsheetToolsObject.ProtectSheet();

              if (MethodFound == true)
              {
                MethodFound = false;

                // announcing to dis start

                if (QuietAnnounce == false)
                  if (DisMessageColumnsAmount > 0)
                  {
                    DisMessageArray = [];

                    for (
                      let DisMessageColumnsIndex = 0;
                      DisMessageColumnsIndex < DisMessageColumnsAmount;
                      DisMessageColumnsIndex++
                    )
                    {
                      DisMessageColumnsCell = DisMessageColumns[DisMessageColumnsIndex];

                      if (DisMessageColumnsCell != null)
                        if (
                          DisMessageColumnsCell.Column != undefined
                          && DisMessageColumnsCell.Column != null
                          && GeneralToolsObject.VariableIsString(DisMessageColumnsCell.Column) == true
                          && DisMessageColumnsCell.Column.length > 0
                          && (
                            DisMessageColumnsCell.TextStyleCode == undefined
                            || DisMessageColumnsCell.TextStyleCode == null
                            || GeneralToolsObject.VariableIsNumber(DisMessageColumnsCell.TextStyleCode) == true
                            && GeneralToolsObject.VariableIsString(DisMessageColumnsCell.TextStyleCode) == false
                            && DisMessageColumnsCell.TextStyleCode >= 0
                            && DisMessageColumnsCell.TextStyleCode <= (
                              ItalicsDisTextStyle.Code
                              | BoldDisTextStyle.Code
                              | UnderlineDisTextStyle.Code
                              | StrikethroughDisTextStyle.Code
                            )
                          )
                        )
                        {
                          MethodFound = true;

                          DisMessageColumnsCellColumn = DisMessageColumnsCell.Column;
                          DisMessageColumnsCellTextStyleCode = DisMessageColumnsCell.TextStyleCode;
                        }

                      if (MethodFound == true)
                      {
                        MethodFound = false;

                        DisMessageColumnsCellColumnValue = null;

                        switch(DisMessageColumnsCellColumn)
                        {
                          case TitleColumnName:
                            // try to avoid showing the trimmed title if there's no need
                            if (TitleColumn != -1)
                              DisMessageColumnsCellColumnValue =
                                TitleValue == null
                                || GeneralToolsObject.VariableIsString(TitleValue) == false
                                || TitleValue.length == 0
                                || TitleValue.length == TitleValueTrimmed.length
                                && TitleValue == TitleValueTrimmed
                                || TitleValueTrimmed.length == NATextLength
                                && TitleValueTrimmed == NAText
                                ? TitleValueTrimmed
                                : TitleValue;

                            break;

                          case TitleIDColumnName:
                            if (TitleIDColumn != -1)
                              DisMessageColumnsCellColumnValue = TitleIDValueTrimmed;

                            break;

                          case RegionColumnName:
                            if (RegionColumn != -1)
                              DisMessageColumnsCellColumnValue = RegionValueTrimmed;

                            break;

                          case GenreColumnName:
                            if (GenreColumn != -1)
                              DisMessageColumnsCellColumnValue = GenreValueTrimmed;

                            break;

                          case BaseSizeColumnName:
                            if (BaseSizeColumn != -1)
                              DisMessageColumnsCellColumnValue = BaseSizeValueTrimmed;

                            break;

                          case VersionColumnName:
                            if (VersionColumn != -1)
                              DisMessageColumnsCellColumnValue = VersionValueTrimmed;

                            break;

                          case UpdateSizeColumnName:
                            if (UpdateSizeColumn != -1)
                              DisMessageColumnsCellColumnValue = UpdateSizeValueTrimmed;

                            break;

                          case DLCColumnName:
                            if (DLCColumn != -1)
                              DisMessageColumnsCellColumnValue = DLCValueTrimmed;

                            break;

                          case BPColumnName:
                            if (BPColumn != -1)
                              DisMessageColumnsCellColumnValue = BPValueTrimmed;

                            break;

                          case TestedColumnName:
                            if (TestedColumn != -1)
                              DisMessageColumnsCellColumnValue = TestedValueTrimmed;

                            break;

                          case InfoColumnName:
                            // try to avoid showing the trimmed info if there's no need
                            if (InfoColumn != -1)
                              DisMessageColumnsCellColumnValue =
                                InfoValue == null
                                || GeneralToolsObject.VariableIsString(InfoValue) == false
                                || InfoValue.length == 0
                                || InfoValue.length == InfoValueTrimmed.length
                                && InfoValue == InfoValueTrimmed
                                || InfoValueTrimmed.length == NATextLength
                                && InfoValueTrimmed == NAText
                                ? InfoValueTrimmed
                                : InfoValue;

                            break;
                        }

                        if (DisMessageColumnsCellColumnValue != null)
                          switch(DisMessageColumnsCellColumn)
                          {
                            case TitleColumnName:
                            case TitleIDColumnName:
                            case RegionColumnName:
                            case GenreColumnName:
                            case BaseSizeColumnName:
                            case VersionColumnName:
                            case InfoColumnName:
                              if (
                                GeneralToolsObject.VariableIsString(DisMessageColumnsCellColumnValue) == true
                                && DisMessageColumnsCellColumnValue.length > 0
                                && GeneralToolsObject.VariableIsNotNA(DisMessageColumnsCellColumnValue) == true
                              )
                                MethodFound = true;

                              break;

                            case UpdateSizeColumnName:
                              if (
                                GeneralToolsObject.VariableIsString(DisMessageColumnsCellColumnValue) == true
                                && DisMessageColumnsCellColumnValue.length > 0
                                && GeneralToolsObject.VariableIsNotNA(DisMessageColumnsCellColumnValue) == true
                                && (
                                  NoUpdateSizeText == undefined
                                  || NoUpdateSizeText == null
                                  || GeneralToolsObject.VariableIsString(NoUpdateSizeText) == false
                                  || NoUpdateSizeText.length == 0
                                  || DisMessageColumnsCellColumnValue.length != NoUpdateSizeText.length
                                  || DisMessageColumnsCellColumnValue != NoUpdateSizeText
                                )
                              )
                                MethodFound = true;

                              break;

                            case DLCColumnName:
                              if (
                                GeneralToolsObject.VariableIsString(DisMessageColumnsCellColumnValue) == true
                                && DisMessageColumnsCellColumnValue.length == 3
                                && DisMessageColumnsCellColumnValue == "Yes"
                              )
                                MethodFound = true;

                              break;

                            case BPColumnName:
                            case TestedColumnName:
                              if (
                                GeneralToolsObject.VariableIsString(DisMessageColumnsCellColumnValue) == true
                                && DisMessageColumnsCellColumnValue.length == 4
                                && DisMessageColumnsCellColumnValue == "TRUE"
                              )
                                MethodFound = true;

                              break;
                          }

                        if (MethodFound == true)
                        {
                          MethodFound = false;

                          switch(DisMessageColumnsCellColumn)
                          {
                            case TitleIDColumnName:
                            case RegionColumnName:
                            case GenreColumnName:
                              DisMessageColumnsCellColumnValue = '[' + DisMessageColumnsCellColumnValue + ']';

                              break;

                            case BaseSizeColumnName:
                              DisMessageColumnsCellColumnValue =
                                '[' + "Base" + ' ' + DisMessageColumnsCellColumnValue + ']';

                              break;

                            case VersionColumnName:
                              DisMessageColumnsCellColumnValue =
                                '[' + 'v' + DisMessageColumnsCellColumnValue + ']';

                              break;

                            case UpdateSizeColumnName:
                              DisMessageColumnsCellColumnValue =
                                '[' + "Update" + ' ' + DisMessageColumnsCellColumnValue + ']';

                              break;

                            case DLCColumnName:
                              DisMessageColumnsCellColumnValue = '[' + "DLC" + ']';

                              break;

                            case BPColumnName:
                              DisMessageColumnsCellColumnValue = '[' + "BP" + ']';

                              break;

                            case TestedColumnName:
                              DisMessageColumnsCellColumnValue = '[' + "Tested" + ']';

                              break;

                            case InfoColumnName:
                              DisMessageColumnsCellColumnValue =
                                DisMessageColumnsIndex + 1 == DisMessageColumnsAmount
                                ? '-' + ' ' + DisMessageColumnsCellColumnValue
                                : '[' + DisMessageColumnsCellColumnValue + ']';

                              break;
                          }

                          if (
                            DisMessageColumnsCellTextStyleCode != undefined
                            && DisMessageColumnsCellTextStyleCode != null
                          )
                            DisMessageColumnsCellColumnValue =
                              DisToolsObject.ApplyTextStyle(
                                DisMessageColumnsCellColumnValue, DisMessageColumnsCellTextStyleCode
                              );

                          DisMessageArray.push(DisMessageColumnsCellColumnValue);
                        }

                        if (DisMessageColumnsCellColumnValue != null)
                          DisMessageColumnsCellColumnValue = null;
                      }

                      DisMessageColumnsCell = null;

                      if (DisMessageColumnsCellColumn != null)
                        DisMessageColumnsCellColumn = null;

                      if (DisMessageColumnsCellTextStyleCode != null)
                        DisMessageColumnsCellTextStyleCode = null;
                    }

                    // needs to contain at least a single column
                    if (DisMessageArray.length > 0)
                    {
                      DisMessageArray =
                        ['[' + SpreadsheetToolsObject.CheckSheetFixedName() + ']']
                        .concat(DisMessageArray)
                        .unique();

                      DisMessage = DisMessageArray.join(' ');

                      DisToolsObject.SendDisMessage(DisMessage);

                      Utilities.sleep(1300);// avoid sending messages to dis too often

                      DisMessage = null;
                    }

                    DisMessageArray = null;
                  }

                // announcing to dis end
              }
            }

            if (ColumnsFixesAmount > 0)
            {
              ColumnsFixes = null;
              ColumnsFixesAmount = 0;
            }

            if (ColumnsFixesText != null)
              ColumnsFixesText = null;
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
              , true
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
            && Rows.length == 6
            && Rows[0] != null
            && Rows[1] != null
            && Rows[2] != null
            && Rows[3] != null
            && Rows[4] != null
            && Rows[5] != null
            && GeneralToolsObject.VariableIsNumber(Rows[1]) == true
            && GeneralToolsObject.VariableIsString(Rows[1]) == false
            && Rows[1] != -1
            && GeneralToolsObject.VariableIsNumber(Rows[3]) == true
            && GeneralToolsObject.VariableIsString(Rows[3]) == false
            && Rows[3] != -1
            && GeneralToolsObject.VariableIsNumber(Rows[5]) == true
            && GeneralToolsObject.VariableIsString(Rows[5]) == false
            && Rows[5] != -1
          )
          {
            RowsRealValues = Rows[0];
            RowsRealValuesAmount = Rows[1];
            RowsDisplayValues = Rows[2];
            RowsDisplayValuesAmount = Rows[3];
            RowsFormulas = Rows[4];
            RowsFormulasAmount = Rows[5];
          }
          else
          {
            RowsRealValues = null;
            RowsRealValuesAmount = 0;
            RowsDisplayValues = null;
            RowsDisplayValuesAmount = 0;
            RowsFormulas = null;
            RowsFormulasAmount = 0;
          }

          if (RowsIndex != 0)
            RowsIndex = 0;

          if (RecheckRows == true)
            RecheckRows = false;
        }
      }

      ////SpreadsheetToolsObject.ProtectSheet();
    }
  }
}

// initializing the announce sheet to dis object
function InitializeAnnounceSheetToDisObject_()
{
  if (AnnounceSheetToDisObject == null)
    AnnounceSheetToDisObject = new AnnounceSheetToDisObjectType();
}

// announcing the active sheet to dis
function AnnounceActiveSheetToDis()
{
  InitializeSpreadsheetToolsObject_();
  InitializeAnnounceSheetToDisObject_();

  SpreadsheetToolsObject.CheckSpreadsheet();
  SpreadsheetToolsObject.CheckActiveSheet();

  AnnounceSheetToDisObject.AnnounceSheetToDis();
}

// announcing all sheets to dis
function AnnounceAllSheetsToDis()
{
  var SpreadsheetReposNamesAmount;

  InitializeSpreadsheetToolsObject_();
  InitializeAnnounceSheetToDisObject_();

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

        AnnounceSheetToDisObject.AnnounceSheetToDis();
      }
    }
  }
}

// testing an announce to dis
function AnnounceTestToDis()
{
  var RandomNumber = Math.random() * 100;

  InitializeDisToolsObject_();

  DisToolsObject.SendDisMessage(
    RandomNumber < 20
    ? "test1"
    : (
      RandomNumber < 40
      ? "test2"
      : (
        RandomNumber < 60
        ? "test3"
        : (
          RandomNumber < 80
          ? "test4"
          : "test5"
        )
      )
    )
  );
}

// adding the announce spreadsheet to dis menu to the spreadsheet UI
function AddAnnounceSpreadsheetToDisMenu()
{
  InitializeSpreadsheetToolsObject_();

  SpreadsheetToolsObject.CheckSpreadsheet();
  
  if (SpreadsheetToolsObject.SpreadsheetUI != null)
  {
    var SpreadsheetUIMenu = SpreadsheetToolsObject.SpreadsheetUI.createMenu("Announce");

    if (SpreadsheetUIMenu != null)
    {
      SpreadsheetUIMenu.addItem("Announce Active Sheet To Dis", "AnnounceActiveSheetToDis");
      SpreadsheetUIMenu.addSeparator();
      SpreadsheetUIMenu.addItem("Announce All Sheets To Dis", "AnnounceAllSheetsToDis");
      //SpreadsheetUIMenu.addSeparator();
      //SpreadsheetUIMenu.addItem("Announce Test To Dis", "AnnounceTestToDis");
      SpreadsheetUIMenu.addToUi();
    }
  }
}
