// user config start

// the color of the links' display texts
var MutualMethodsLinksColor = "orange";

// user config end

var MutualMethodsObject = null;// do not delete

// applying the cells values in MB (1 KB = 1000 B, 1 MB = 1000 KB, etc)
var MutualMethodsFileSizeMeasureType = 2;// do not delete

class MutualMethodsObjectType
{
  constructor()
  {
    this.FileSizeDoubleNumberFormat = null;
    this.FileSizeIntegerNumberFormat = null;

    InitializeGeneralToolsObject_();
    InitializeSpreadsheetToolsObject_();

    this.FileSizeDoubleNumberFormat = 
      '[' + '<' + (
        MutualMethodsFileSizeMeasureType == 0 ? 1000 * 1000 : 1000
      ).toString() + ']';

    this.FileSizeDoubleNumberFormat += "0.00";

    if (MutualMethodsFileSizeMeasureType == 0)
      this.FileSizeDoubleNumberFormat += ',';

    this.FileSizeDoubleNumberFormat +=
      '"' + ' ' + (
        MutualMethodsFileSizeMeasureType <= 1
        ? "KB"
        : (
          MutualMethodsFileSizeMeasureType == 2
          ? "MB"
          : (
            MutualMethodsFileSizeMeasureType == 3
            ? "GB"
            : "TB"
          )
        )
      ) + '"' + ';';

    this.FileSizeDoubleNumberFormat +=
      '[' + '<' + (
        MutualMethodsFileSizeMeasureType == 0 ? 1000 * 1000 * 1000 : 1000 * 1000
      ).toString() + ']';
      
    this.FileSizeDoubleNumberFormat += "0.00" + ',';

    if (MutualMethodsFileSizeMeasureType == 0)
      this.FileSizeDoubleNumberFormat += ',';

    this.FileSizeDoubleNumberFormat +=
      '"' + ' ' + (
        MutualMethodsFileSizeMeasureType <= 1
        ? "MB"
        : (
          MutualMethodsFileSizeMeasureType == 2
          ? "GB"
          : (
            MutualMethodsFileSizeMeasureType == 3
            ? "TB"
            : "PB"
          )
        )
      ) + '"' + ';';
      
    this.FileSizeDoubleNumberFormat += "0.00" + ',' + ',';

    if (MutualMethodsFileSizeMeasureType == 0)
      this.FileSizeDoubleNumberFormat += ',';

    this.FileSizeDoubleNumberFormat +=
      '"' + ' ' + (
        MutualMethodsFileSizeMeasureType <= 1
        ? "GB"
        : (
          MutualMethodsFileSizeMeasureType == 2
          ? "TB"
          : (
            MutualMethodsFileSizeMeasureType == 3
            ? "PB"
            : "EB"
          )
        )
      ) + '"';
      
    this.FileSizeIntegerNumberFormat = this.FileSizeDoubleNumberFormat.replaceAll("0.00", "0");
  }

  // a method for checking text cells
  CheckTextMethod(Value)
  {
    return (
      Value != null
      ? (
        GeneralToolsObject.VariableIsString(Value) == true
        && Value.length > 0
        && GeneralToolsObject.VariableIsNotNA(Value) == true
        ? Value
        : NAText
      )
      : null
    );
  }

  // a method for checking check box cells
  CheckCheckBoxMethod(Value)
  {
    var FunctionResult;
    var CurrentResult;

    if (Value != null)
      if (GeneralToolsObject.VariableIsString(Value) == true)
      {
        var ValueLength = Value.length;

        if (
          ValueLength > 0
          && GeneralToolsObject.VariableIsNotNA(Value) == true
        )
        {
          var CapitalValue = Value.toUpperCase();

          CurrentResult =
            ValueLength == 5 && CapitalValue == "FALSE"
            ? "FALSE"
            : (
              ValueLength == 4 && CapitalValue == "TRUE"
              ? "TRUE"
              : Value
            );
        }
        else
          CurrentResult = NAText;
      }
      else
        CurrentResult = NAText;
    else
      CurrentResult = null;

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // a method for checking file size cells
  CheckFileSizeMethod(Value)
  {
    var FunctionResult;
    var CurrentResult;

    if (Value != null)
      if (GeneralToolsObject.VariableIsString(Value) == true)
      {
        var ValueLength = Value.length;

        if (
          ValueLength > 0
          && GeneralToolsObject.VariableIsNotNA(Value) == true
        )
        {
          var OptimalFileSizeText = GeneralToolsObject.CheckOptimalFileSizeText(Value, null);

          CurrentResult =
            OptimalFileSizeText != null
            && GeneralToolsObject.VariableIsString(OptimalFileSizeText) == true
            && OptimalFileSizeText.length > 0
            ? OptimalFileSizeText
            : Value;
        }
        else
          CurrentResult = NAText;
      }
      else
        CurrentResult = NAText;
    else
      CurrentResult = null;

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // a method for fixing the update size
  FixUpdateSizeMethod(Value)
  {
    var FunctionResult;
    var CurrentResult = null;

    if (
      Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
    )
    {
      var ValueLength = Value.length;

      CurrentResult = Value;

      if (ValueLength > 0)
      {
        var CapitalValue = Value.toUpperCase();

        if (
          ValueLength == 6
          && CapitalValue == "MERGED"
        )
          CurrentResult = "Merged";
        else if (
          NoUpdateSizeText != undefined
          && NoUpdateSizeText != null
          && GeneralToolsObject.VariableIsString(NoUpdateSizeText) == true
          && ValueLength == NoUpdateSizeText.length
          && CapitalValue == NoUpdateSizeText.toUpperCase()
        )
          CurrentResult = NoUpdateSizeText;
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // a method for checking link cells
  CheckLinkMethod(Row, Column, Value)
  {
    var FunctionResult;
    var CurrentResult;

    if (Value != null)
      if (
        GeneralToolsObject.VariableIsString(Value) == true
        && Value.length > 0
      )
      {
        var Hyperlink = SpreadsheetToolsObject.CheckHyperlink(Row, Column);

        CurrentResult =
          Hyperlink != null
          && GeneralToolsObject.VariableIsString(Hyperlink) == true
          && Hyperlink.length > 0
          ? Hyperlink
          : (
            GeneralToolsObject.ValidateLink(Value) == true
            ? Value
            : NAText
          );
      }
      else
        CurrentResult = NAText;
    else
      CurrentResult = null;

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // a method for applying file size cells
  ApplyFileSizeMethod(Row, Column, FileSizeNumber, FileSizeNumberFormat, RowDisplayValues, RowRealValues = null)
  {
    var FunctionResult;
    var CurrentResult = null;

    if (
      Row > 0
      && Column != -1
      && FileSizeNumber != null
      && FileSizeNumberFormat != null
      && RowDisplayValues != null
      && GeneralToolsObject.VariableIsNumber(FileSizeNumber) == true
      && (
        GeneralToolsObject.VariableIsString(FileSizeNumber) == true
        ? Number(FileSizeNumber)
        : FileSizeNumber
      ) > 0
      && GeneralToolsObject.VariableIsString(FileSizeNumberFormat) == true
      && FileSizeNumberFormat.length > 0
      && RowDisplayValues.length > 0
    )
    {
      var CellValue =
        GeneralToolsObject.VariableIsString(FileSizeNumber) == true
        ? Number(FileSizeNumber)
        : FileSizeNumber;

      var CellFormat = FileSizeNumberFormat;
      var Cell = SpreadsheetToolsObject.CheckCell(Row, Column);

      var OptimalFileSizeNumber = CellValue;
      var OptimalFileSizeMeasureType = MutualMethodsFileSizeMeasureType;
      var OptimalFileSizeMeasure;
      
      if (dry_run == false)
      {
        if (RowRealValues != null)
          RowRealValues[Column] = CellValue;
        else
          Cell.setValue(CellValue);

        Cell.setNumberFormat(CellFormat);
      }

      while (
        OptimalFileSizeNumber >= 1000
        && OptimalFileSizeMeasureType < 6// Exabytes
      )
      {
        OptimalFileSizeNumber /= 1000;
        OptimalFileSizeMeasureType++;
      }

      OptimalFileSizeMeasure = GeneralToolsObject.CheckFileSizeMeasure(null, OptimalFileSizeMeasureType);

      if (
        OptimalFileSizeMeasure != null
        && GeneralToolsObject.VariableIsString(OptimalFileSizeMeasure) == true
        && OptimalFileSizeMeasure.length > 0
      )
      {
        var OptimalFileSizeText =
          GeneralToolsObject.CheckOptimalFileSizeText(
            OptimalFileSizeNumber.toString()
            + ' '
            + OptimalFileSizeMeasure
          );

        if (
          OptimalFileSizeText != null
          && GeneralToolsObject.VariableIsString(OptimalFileSizeText) == true
          && OptimalFileSizeText.length > 0
        )
          CurrentResult = OptimalFileSizeText;
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // a method for applying version cells
  ApplyVersionMethod(Row, Column, Value, RowRealValues = null)
  {
    var FunctionResult;
    var CurrentResult = null;

    if (
      Row > 0
      && Column != -1
      && Value != null
      && GeneralToolsObject.VariableIsNumber(Value) == true
      && (
        GeneralToolsObject.VariableIsString(Value) == true
        ? Number(Value)
        : Value
      ) > 0
    )
    {
      var CellValue =
        GeneralToolsObject.VariableIsString(Value) == true
        ? Number(Value)
        : Value;

      var CellFormat =
        VersionFormat != undefined
        && VersionFormat != null
        && GeneralToolsObject.VariableIsString(VersionFormat) == true
        && VersionFormat.length > 0
        ? VersionFormat
        : null;

      var Cell = SpreadsheetToolsObject.CheckCell(Row, Column);
      
      if (dry_run == false)
      {
        if (RowRealValues != null)
          RowRealValues[Column] = CellValue;
        else
          Cell.setValue(CellValue);

        if (CellFormat != null)
          Cell.setNumberFormat(CellFormat);
      }

      CurrentResult = (GeneralToolsObject.FloatToInteger(CellValue * 100, 5) / 100).toFixed(2);
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // a method for applying date cells
  ApplyDateMethod(Row, Column, Value, RowRealValues = null)
  {
    var FunctionResult;
    var CurrentResult = null;

    if (
      Row > 0
      && Column != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      var CellValue = Value;
      var CellFormat =
        DateFormat != undefined
        && DateFormat != null
        && GeneralToolsObject.VariableIsString(DateFormat) == true
        && DateFormat.length > 0
        ? DateFormat
        : null;

      var Cell = SpreadsheetToolsObject.CheckCell(Row, Column);
      
      if (dry_run == false)
      {
        if (RowRealValues != null)
          RowRealValues[Column] = CellValue;
        else
          Cell.setValue(CellValue);

        if (CellFormat != null)
          Cell.setNumberFormat(CellFormat);
      }

      CurrentResult = CellValue;
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // a method for applying link cells
  ApplyLinkMethod(Row, Column, Value, RowRealValues = null, RowFormulas = null)
  {
    var FunctionResult;
    var CurrentResult = null;

    if (
      Row > 0
      && Column != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      var Link = Value;
      var LinkDomain = GeneralToolsObject.CheckLinkDomain(Link);

      if (
        LinkDomain != null
        && GeneralToolsObject.VariableIsString(LinkDomain) == true
        && LinkDomain.length > 0
      )
      {
        var LinkDomainName = GeneralToolsObject.CheckLinkDomainName(LinkDomain);

        if (
          LinkDomainName != null
          && GeneralToolsObject.VariableIsString(LinkDomainName) == true
          && LinkDomainName.length > 0
        )
        {
          var CellValue = LinkDomainName;
          var CellFormula =
            "=HYPERLINK" + '('
            + '"' + Link + '"'
            + ',' + '"' + CellValue + '"'
            + ')';

          var CellColor =
            MutualMethodsLinksColor != undefined
            && MutualMethodsLinksColor != null
            && GeneralToolsObject.VariableIsString(MutualMethodsLinksColor) == true
            && MutualMethodsLinksColor.length > 0
            ? MutualMethodsLinksColor
            : null;

          var Cell = SpreadsheetToolsObject.CheckCell(Row, Column);
          
          if (dry_run == false)
          {
            if (RowRealValues != null)
              RowRealValues[Column] = CellValue;
            // not necessary to actually set the value, because a cell is either a value or a formula
            // , and in this case, the formula is what matters
            //else
            //  Cell.setValue(CellValue);

            if (RowFormulas != null)
              RowFormulas[Column] = CellFormula;
            else
              Cell.setFormula(CellFormula);

            if (CellColor != null)
              Cell.setFontColor(CellColor);
          }

          CurrentResult = Link;
        }
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // checking the title value
  CheckTitle(RowDisplayValues)
  {
    var Cell = TitleColumn != -1 ? RowDisplayValues[TitleColumn] : null;

    TitleValue =
      Cell != null
      && GeneralToolsObject.VariableIsString(Cell) == true
      && Cell.length > 0
      ? Cell
      : null;

    if (TitleValue != null)
      TitleValueTrimmed = TitleValue.trim();
  }

  // checking the title id value
  CheckTitleID(RowDisplayValues, ChangeRowDisplayValues = false)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    var Cell = TitleIDColumn != -1 ? RowDisplayValues[TitleIDColumn] : null;

    TitleIDValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (TitleIDValue != null)
    {
      TitleIDValueTrimmed = this.CheckTextMethod(TitleIDValue.trim());

      if (TitleIDValueTrimmed != null)
      {
        var TitleIDValueTrimmedLength = TitleIDValueTrimmed.length;

        if (
          TitleIDValueTrimmedLength > 0
          && (TitleIDValueTrimmedLength != NATextLength || TitleIDValueTrimmed != NAText)
        )
        {
          TitleIDValueTrimmed = TitleIDValueTrimmed.toUpperCase();

          if (GeneralToolsObject.VariableIsNumber(TitleIDValueTrimmed) == true)
            if (
              TitleIDMinNumbersAmount != undefined
              && TitleIDMinNumbersAmount != null
              && GeneralToolsObject.VariableIsNumber(TitleIDMinNumbersAmount) == true
              && GeneralToolsObject.VariableIsString(TitleIDMinNumbersAmount) == false
              || TitleIDType != undefined
              && TitleIDType != null
              && GeneralToolsObject.VariableIsString(TitleIDType) == true
              && TitleIDType.length > 0
            )
            {
              if (
                TitleIDMinNumbersAmount != undefined
                && TitleIDMinNumbersAmount != null
                && GeneralToolsObject.VariableIsNumber(TitleIDMinNumbersAmount) == true
                && GeneralToolsObject.VariableIsString(TitleIDMinNumbersAmount) == false
              )
                while (TitleIDValueTrimmedLength < TitleIDMinNumbersAmount)
                {
                  TitleIDValueTrimmed = '0' + TitleIDValueTrimmed;
                  TitleIDValueTrimmedLength++;
                }

              if (
                TitleIDType != undefined
                && TitleIDType != null
                && GeneralToolsObject.VariableIsString(TitleIDType) == true
              )
              {
                var TitleIDTypeLength = TitleIDType.length;

                if (TitleIDTypeLength > 0)
                {
                  TitleIDValueTrimmed = TitleIDType + TitleIDValueTrimmed;
                  TitleIDValueTrimmedLength += TitleIDTypeLength;
                }
              }
            }
            else
              TitleIDValueTrimmed = NAText;
        }
      }

      if (ChangeRowDisplayValues == true)
        if (TitleIDValueTrimmed != null)
          if (RowDisplayValues[TitleIDColumn] != TitleIDValueTrimmed)
          {
            CurrentResult = true;

            RowDisplayValues[TitleIDColumn] = TitleIDValueTrimmed;
          }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // checking the region value
  CheckRegion(RowDisplayValues, ChangeRowDisplayValues = false)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    var Cell = RegionColumn != -1 ? RowDisplayValues[RegionColumn] : null;

    RegionValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (RegionValue != null)
    {
      RegionValueTrimmed = this.CheckTextMethod(RegionValue.trim());

      if (RegionValueTrimmed != null)
      {
        var RegionValueTrimmedLength = RegionValueTrimmed.length;
        
        if (
          RegionValueTrimmedLength > 0
          && (RegionValueTrimmedLength != NATextLength || RegionValueTrimmed != NAText)
        )
        {
          RegionValueTrimmed = RegionValueTrimmed.toUpperCase();// E

          if (
            KnownRegions != undefined
            && KnownRegions != null
          )
          {
            var KnownRegionsAmount = KnownRegions.length;

            if (KnownRegionsAmount > 0)
            {
              var KnownRegion;

              var KnownRegionShorts;
              var KnownRegionShortsAmount;
              var KnownRegionShort;

              var KnownRegionFull;

              var MethodFound = false;

              for (let KnownRegionsIndex = 0; KnownRegionsIndex < KnownRegionsAmount; KnownRegionsIndex++)
              {
                KnownRegion = KnownRegions[KnownRegionsIndex];

                if (KnownRegion != null)
                  if (
                    KnownRegion.Shorts != undefined
                    && KnownRegion.Shorts != null
                    && KnownRegion.Full != undefined
                    && KnownRegion.Full != null
                    && GeneralToolsObject.VariableIsString(KnownRegion.Full) == true
                    && KnownRegion.Full.length > 0
                  )
                  {
                    KnownRegionShorts = KnownRegion.Shorts;
                    KnownRegionShortsAmount = KnownRegionShorts.length;

                    if (KnownRegionShortsAmount > 0)
                    {
                      KnownRegionFull = KnownRegion.Full;// Europe

                      for (
                        let KnownRegionShortsIndex = 0;
                        KnownRegionShortsIndex < KnownRegionShortsAmount;
                        KnownRegionShortsIndex++
                      )
                      {
                        KnownRegionShort = KnownRegionShorts[KnownRegionShortsIndex];// e

                        if (
                          KnownRegionShort != null
                          && GeneralToolsObject.VariableIsString(KnownRegionShort) == true
                          && KnownRegionShort.length > 0
                        )
                          if (RegionValueTrimmed == KnownRegionShort.toUpperCase())// E
                          {
                            MethodFound = true;

                            RegionValueTrimmed = KnownRegionFull;// Europe

                            break;
                          }
                      }

                      if (MethodFound == true)
                      {
                        MethodFound = false;

                        break;
                      }
                    }
                  }
              }
            }
          }
        }
      }

      if (ChangeRowDisplayValues == true)
        if (RegionValueTrimmed != null)
          if (RowDisplayValues[RegionColumn] != RegionValueTrimmed)
          {
            CurrentResult = true;

            RowDisplayValues[RegionColumn] = RegionValueTrimmed;
          }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // checking the genre value
  CheckGenre(RowDisplayValues, ChangeRowDisplayValues = false)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    var Cell = GenreColumn != -1 ? RowDisplayValues[GenreColumn] : null;

    GenreValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (GenreValue != null)
    {
      GenreValueTrimmed = this.CheckTextMethod(GenreValue.trim());

      if (ChangeRowDisplayValues == true)
        if (GenreValueTrimmed != null)
          if (RowDisplayValues[GenreColumn] != GenreValueTrimmed)
          {
            CurrentResult = true;

            RowDisplayValues[GenreColumn] = GenreValueTrimmed;
          }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // checking the base size value (the full and slower function)
  // this function isn't 100% reliable so try to avoid using it
  CheckBaseSizeA(RowDisplayValues)
  {
    var Cell = BaseSizeColumn != -1 ? RowDisplayValues[BaseSizeColumn] : null;

    BaseSizeValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (BaseSizeValue != null)
      BaseSizeValueTrimmed = this.CheckFileSizeMethod(BaseSizeValue.trim());
  }

  // checking the base size value (the short and faster function)
  CheckBaseSizeB(RowDisplayValues)
  {
    var Cell = BaseSizeColumn != -1 ? RowDisplayValues[BaseSizeColumn] : null;

    BaseSizeValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (BaseSizeValue != null)
      BaseSizeValueTrimmed = this.CheckTextMethod(BaseSizeValue.trim());
  }

  // checking the version value
  CheckVersion(RowDisplayValues)
  {
    var Cell = VersionColumn != -1 ? RowDisplayValues[VersionColumn] : null;

    VersionValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (VersionValue != null)
    {
      VersionValueTrimmed = VersionValue.trim();

      VersionValueTrimmed =
        VersionValueTrimmed != null
        ? (
          VersionValueTrimmed.length > 0
          && GeneralToolsObject.VariableIsNotNA(VersionValueTrimmed) == true
          && GeneralToolsObject.VariableIsNumber(VersionValueTrimmed) == true
          && Number(VersionValueTrimmed) > 0
          ? (
            Number(VersionValueTrimmed) >= 1
            ? (GeneralToolsObject.FloatToInteger(Number(VersionValueTrimmed) * 100, 5) / 100).toFixed(2)
            : NAText//Number(VersionValueTrimmed)
          )
          : NAText
        )
        : null;
    }
  }

  // checking the update size value (the full and slower function)
  // this function isn't 100% reliable so try to avoid using it
  CheckUpdateSizeA(RowDisplayValues)
  {
    var Cell = UpdateSizeColumn != -1 ? RowDisplayValues[UpdateSizeColumn] : null;

    UpdateSizeValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (UpdateSizeValue != null)
      UpdateSizeValueTrimmed = this.FixUpdateSizeMethod(this.CheckFileSizeMethod(UpdateSizeValue.trim()));
  }

  // checking the update size value (the short and faster function)
  CheckUpdateSizeB(RowDisplayValues)
  {
    var Cell = UpdateSizeColumn != -1 ? RowDisplayValues[UpdateSizeColumn] : null;

    UpdateSizeValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (UpdateSizeValue != null)
      UpdateSizeValueTrimmed = this.FixUpdateSizeMethod(this.CheckTextMethod(UpdateSizeValue.trim()));
  }

  // checking the dlc value
  CheckDLC(RowDisplayValues, ChangeRowDisplayValues = false)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    var Cell = DLCColumn != -1 ? RowDisplayValues[DLCColumn] : null;

    DLCValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (DLCValue != null)
    {
      DLCValueTrimmed = DLCValue.trim();

      if (DLCValueTrimmed != null)
      {
        var DLCValueTrimmedLength = DLCValueTrimmed.length;

        if (DLCValueTrimmedLength == NATextLength || DLCValueTrimmedLength == 3 || DLCValueTrimmedLength == 2)
        {
          var DLCCapitalValueTrimmed = DLCValueTrimmed.toUpperCase();

          if (DLCValueTrimmedLength == NATextLength && DLCCapitalValueTrimmed == NAText.toUpperCase())
            DLCValueTrimmed = NAText;
          else if (DLCValueTrimmedLength == 3)
            if (DLCCapitalValueTrimmed == "YES")
              DLCValueTrimmed = "Yes";
            else if (
              DLCCapitalValueTrimmed == "N/A"
              || DLCCapitalValueTrimmed == "N\\A"
            )
              DLCValueTrimmed = NAText;
            else
              DLCValueTrimmed = NAText;
          else if (DLCValueTrimmedLength == 2)
            if (DLCCapitalValueTrimmed == "NO")
              DLCValueTrimmed = "No";
            else if (DLCCapitalValueTrimmed == "NA")
              DLCValueTrimmed = NAText;
            else
              DLCValueTrimmed = NAText;
          else
            DLCValueTrimmed = NAText;
        }
        else
          DLCValueTrimmed = NAText;
      }

      if (ChangeRowDisplayValues == true)
        if (DLCValueTrimmed != null)
          if (RowDisplayValues[DLCColumn] != DLCValueTrimmed)
          {
            CurrentResult = true;

            RowDisplayValues[DLCColumn] = DLCValueTrimmed;
          }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // checking the bp value
  CheckBP(RowDisplayValues)
  {
    var Cell = BPColumn != -1 ? RowDisplayValues[BPColumn] : null;

    BPValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (BPValue != null)
      BPValueTrimmed = this.CheckCheckBoxMethod(BPValue.trim());
  }

  // checking the uploader value
  CheckUploader(RowDisplayValues)
  {
    var Cell = UploaderColumn != -1 ? RowDisplayValues[UploaderColumn] : null;

    UploaderValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (UploaderValue != null)
      UploaderValueTrimmed = this.CheckTextMethod(UploaderValue.trim());
  }

  // checking the date value
  CheckDate(RowDisplayValues, ChangeRowDisplayValues = false)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    var Cell = DateColumn != -1 ? RowDisplayValues[DateColumn] : null;

    DateValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (DateValue != null)
    {
      DateValueTrimmed = this.CheckTextMethod(DateValue.trim());

      if (ChangeRowDisplayValues == true)
        if (DateValueTrimmed != null)
          if (RowDisplayValues[DateColumn] != DateValueTrimmed)
          {
            CurrentResult = true;

            RowDisplayValues[DateColumn] = DateValueTrimmed;
          }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // checking the link value (the full and slower function)
  CheckLinkA(RowsIndex, RowDisplayValues)
  {
    var Cell = LinkColumn != -1 ? RowDisplayValues[LinkColumn] : null;

    LinkValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (LinkValue != null)
      LinkValueTrimmed = this.CheckLinkMethod(RowsIndex, LinkColumn, LinkValue.trim());
  }

  // checking the link value (the short and faster function)
  CheckLinkB(RowDisplayValues)
  {
    var Cell = LinkColumn != -1 ? RowDisplayValues[LinkColumn] : null;

    LinkValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (LinkValue != null)
      LinkValueTrimmed = this.CheckTextMethod(LinkValue.trim());
  }

  // checking the tested value
  CheckTested(RowDisplayValues)
  {
    var Cell = TestedColumn != -1 ? RowDisplayValues[TestedColumn] : null;

    TestedValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (TestedValue != null)
      TestedValueTrimmed = this.CheckCheckBoxMethod(TestedValue.trim());
  }

  // checking the info value
  CheckInfo(RowDisplayValues, ChangeRowDisplayValues = false)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    var Cell = InfoColumn != -1 ? RowDisplayValues[InfoColumn] : null;

    InfoValue =
      Cell != null
      && GeneralToolsObject.VariableIsString(Cell) == true
      && Cell.length > 0
      ? Cell
      : null;

    if (InfoValue != null)
    {
      InfoValueTrimmed = this.CheckTextMethod(InfoValue.trim());

      if (ChangeRowDisplayValues == true)
        if (InfoValueTrimmed != null)
          if (RowDisplayValues[InfoColumn] != InfoValueTrimmed)
          {
            CurrentResult = true;

            RowDisplayValues[InfoColumn] = InfoValueTrimmed;
          }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // checking the DK's mirror value (the full and slower function)
  CheckDKsMirrorA(RowsIndex, RowDisplayValues)
  {
    var Cell = DKsMirrorColumn != -1 ? RowDisplayValues[DKsMirrorColumn] : null;

    DKsMirrorValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (DKsMirrorValue != null)
      DKsMirrorValueTrimmed = this.CheckLinkMethod(RowsIndex, DKsMirrorColumn, DKsMirrorValue.trim());
  }

  // checking the DK's mirror value (the short and faster function)
  CheckDKsMirrorB(RowDisplayValues)
  {
    var Cell = DKsMirrorColumn != -1 ? RowDisplayValues[DKsMirrorColumn] : null;

    DKsMirrorValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (DKsMirrorValue != null)
      DKsMirrorValueTrimmed = this.CheckTextMethod(DKsMirrorValue.trim());
  }

  // checking the update link value (the full and slower function)
  CheckUpdateLinkA(RowsIndex, RowDisplayValues)
  {
    var Cell = UpdateLinkColumn != -1 ? RowDisplayValues[UpdateLinkColumn] : null;

    UpdateLinkValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (UpdateLinkValue != null)
      UpdateLinkValueTrimmed = this.CheckLinkMethod(RowsIndex, UpdateLinkColumn, UpdateLinkValue.trim());
  }

  // checking the update link value (the short and faster function)
  CheckUpdateLinkB(RowDisplayValues)
  {
    var Cell = UpdateLinkColumn != -1 ? RowDisplayValues[UpdateLinkColumn] : null;

    UpdateLinkValue =
      Cell != null
      ? (
        GeneralToolsObject.VariableIsString(Cell) == true
        && Cell.length > 0
        ? Cell
        : NAText
      )
      : null;

    if (UpdateLinkValue != null)
      UpdateLinkValueTrimmed = this.CheckTextMethod(UpdateLinkValue.trim());
  }

  // checking the submitter value
  CheckSubmitter(RowDisplayValues)
  {
    // the one who's announcing is the submitter
    // a flaw is that if 2 people or more work on the spreadsheet in the same time and 1 of them is announcing
    // then it might announce other people's work and input the submitter as the person who launched the script
    // in the entries that they work on
    // use this at your own risk
    if (
      AnnouncedValue != null
      && GeneralToolsObject.VariableIsString(AnnouncedValue) == true
      && AnnouncedValue.length == 2
      && AnnouncedValue == "No"
    )
      if (
        Submitters != undefined
        && Submitters != null
      )
      {
        var SubmittersAmount = Submitters.length;

        if (SubmittersAmount > 0)
        {
          SpreadsheetToolsObject.CheckUserName();

          if (
            SpreadsheetToolsObject.UserName != null
            && GeneralToolsObject.VariableIsString(SpreadsheetToolsObject.UserName) == true
            && SpreadsheetToolsObject.UserName.length > 0
          )
          {
            var Submitter;
            var SubmitterRealUserName;
            var SubmitterDisplayedUserName;

            var CapitalUserName = SpreadsheetToolsObject.UserName.toUpperCase();// EXAMPLE_USER

            for (let SubmittersIndex = 0; SubmittersIndex < SubmittersAmount; SubmittersIndex++)
            {
              Submitter = Submitters[SubmittersIndex];

              if (Submitter != null)
                if (
                  Submitter.RealUserName != undefined
                  && Submitter.RealUserName != null
                  && Submitter.DisplayedUserName != undefined
                  && Submitter.DisplayedUserName != null
                  && GeneralToolsObject.VariableIsString(Submitter.RealUserName) == true
                  && Submitter.RealUserName.length > 0
                  && GeneralToolsObject.VariableIsString(Submitter.DisplayedUserName) == true
                  && Submitter.DisplayedUserName.length > 0
                )
                {
                  SubmitterRealUserName = Submitter.RealUserName;// example_user

                  if (CapitalUserName == SubmitterRealUserName.toUpperCase())// EXAMPLE_USER
                  {
                    SubmitterDisplayedUserName = Submitter.DisplayedUserName;// Example

                    SubmitterValue = SubmitterDisplayedUserName;// Example

                    break;
                  }
                }
            }
          }
        }
      }

    if (SubmitterValue != null)
      SubmitterValueTrimmed = SubmitterValue;
    else
    {
      var Cell = SubmitterColumn != -1 ? RowDisplayValues[SubmitterColumn] : null;

      SubmitterValue =
        Cell != null
        ? (
          GeneralToolsObject.VariableIsString(Cell) == true
          && Cell.length > 0
          ? Cell
          : NAText
        )
        : null;

      if (SubmitterValue != null)
        SubmitterValueTrimmed = this.CheckTextMethod(SubmitterValue.trim());
    }
  }

  // checking the announced value
  // better add NAText manually to rows that should get skipped in order to shorten the script's running time
  CheckAnnounced(RowDisplayValues)
  {
    var Cell = AnnouncedColumn != -1 ? RowDisplayValues[AnnouncedColumn] : null;

    if (Cell != null)
      if (GeneralToolsObject.VariableIsString(Cell) == true)
      {
        var CellLength = Cell.length;

        if (CellLength == NATextLength || CellLength == 3 || CellLength == 2)
        {
          var AnnouncedCapitalValue = Cell.toUpperCase();

          if (CellLength == NATextLength && AnnouncedCapitalValue == NAText.toUpperCase())
            AnnouncedValue = NAText;
          else if (CellLength == 3)
            if (AnnouncedCapitalValue == "YES")
              AnnouncedValue = "Yes";
            else if (
              AnnouncedCapitalValue == "N/A"
              || AnnouncedCapitalValue == "N\\A"
            )
              AnnouncedValue = NAText;
            else
              AnnouncedValue = "No";
          else if (CellLength == 2)
            if (AnnouncedCapitalValue == "NO")
              AnnouncedValue = "No";
            else if (AnnouncedCapitalValue == "NA")
              AnnouncedValue = NAText;
            else
              AnnouncedValue = "No";
          else
            AnnouncedValue = "No";
        }
        else
          AnnouncedValue = "No";
      }
      else
        AnnouncedValue = "No";
    else
      AnnouncedValue = null;
  }

  // checking the fixes value
  CheckFixes(RowDisplayValues)
  {
    var Cell = FixesColumn != -1 ? RowDisplayValues[FixesColumn] : null;
    var CellIsString = false;
    var CellLength = 0;
    var CellCapitalValue = null;

    var NoFixesTextLength = 0;
    var NoFixesTextCapitalValue = null;

    if (Cell != null)
    {
      CellIsString = GeneralToolsObject.VariableIsString(Cell);

      if (CellIsString == true)
      {
        CellLength = Cell.length;

        if (CellLength > 0)
          if (CellLength == NATextLength || CellLength == NoFixesTextLength || CellLength == 3 || CellLength == 2)
          {
            CellCapitalValue = Cell.toUpperCase();

            if (CellLength == NATextLength && CellCapitalValue == NAText.toUpperCase())
              FixesValue = "Fix";
            else
            {
              NoFixesTextLength =
                NoFixesText != undefined
                && NoFixesText != null
                && GeneralToolsObject.VariableIsString(NoFixesText) == true
                ? NoFixesText.length
                : 0;

              if (CellLength == NoFixesTextLength)
                NoFixesTextCapitalValue = NoFixesText.toUpperCase();

              if (CellLength == NoFixesTextLength && CellCapitalValue == NoFixesTextCapitalValue)
                FixesValue = "None";
              else if (CellLength == 3)
                if (
                  CellCapitalValue == "N/A"
                  || CellCapitalValue == "N\\A"
                )
                  FixesValue = "Fix";
                else
                  FixesValue = Cell;
              else if (CellLength == 2)
                if (CellCapitalValue == "NO")
                  FixesValue = "Fix";
                else if (CellCapitalValue == "NA")
                  FixesValue = "Fix";
                else
                  FixesValue = Cell;
              else
                FixesValue = Cell;
            }
          }
          else
            FixesValue = Cell;
        else
          FixesValue = "Fix";
      }
      else
        FixesValue = "Fix";
    }
    else
      FixesValue = null;

    if (
      Cell != null
      && CellIsString == true
      && CellLength > 0
    )
      if (CellLength == NATextLength || CellLength == NoFixesTextLength || CellLength == 3 || CellLength == 2)
        if (
          CellLength == NATextLength && CellCapitalValue == NAText.toUpperCase()
          || CellLength == NoFixesTextLength && CellCapitalValue == NoFixesTextCapitalValue
        )
          FixesValueTrimmed = FixesValue;
        else if (CellLength == 3)
          if (
            CellCapitalValue == "N/A"
            || CellCapitalValue == "N\\A"
          )
            FixesValueTrimmed = FixesValue;
          else
            FixesValueTrimmed = FixesValue.trim();
        else if (CellLength == 2)
          if (
            CellCapitalValue == "NO"
            || CellCapitalValue == "NA"
          )
            FixesValueTrimmed = FixesValue;
          else
            FixesValueTrimmed = FixesValue.trim();
        else
          FixesValueTrimmed = FixesValue.trim();
      else
        FixesValueTrimmed = FixesValue.trim();
    else
      FixesValueTrimmed = FixesValue;
  }

  // doing verifications in order to verify a row wasn't added, changed or removed during the process

  // verifying the title value
  VerifyTitle(RowsIndex, RowRealValues)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    if (
      RowsIndex > 0
      && TitleColumn != -1
      && RowRealValues[TitleColumn] != null
    )
    {
      var Cell = SpreadsheetToolsObject.CheckCell(RowsIndex, TitleColumn);

      if (Cell != null)
      {
        var CellValue = Cell.getValue();
        
        if (CellValue != null)
          if (RowRealValues[TitleColumn] == CellValue)
            CurrentResult = true;
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // verifying the title id value
  VerifyTitleID(RowsIndex, RowRealValues)
  {
    var FunctionResult;
    var CurrentResult = false;
    
    if (
      RowsIndex > 0
      && TitleIDColumn != -1
      && RowRealValues[TitleIDColumn] != null
    )
    {
      var Cell = SpreadsheetToolsObject.CheckCell(RowsIndex, TitleIDColumn);

      if (Cell != null)
      {
        var CellValue = Cell.getValue();
        
        if (CellValue != null)
          if (
            RowRealValues[TitleIDColumn] == CellValue
            // if both are N/A the texts might still be different if it got formatted in the row real values
            || GeneralToolsObject.VariableIsNotNA(RowRealValues[TitleIDColumn]) == false
            && GeneralToolsObject.VariableIsNotNA(CellValue) == false
          )
            CurrentResult = true;
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // verifying the region value
  VerifyRegion(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && RegionColumn != -1
    //  ?
    //  RowRealValues[RegionColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, RegionColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the genre value
  VerifyGenre(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && GenreColumn != -1
    //  ?
    //  RowRealValues[GenreColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, GenreColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the base size value
  VerifyBaseSize(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && BaseSizeColumn != -1
    //  ?
    //  RowRealValues[BaseSizeColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, BaseSizeColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the version value
  VerifyVersion(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && VersionColumn != -1
    //  ?
    //  RowRealValues[VersionColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, VersionColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the update size value
  VerifyUpdateSize(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && UpdateSizeColumn != -1
    //  ?
    //  RowRealValues[UpdateSizeColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, UpdateSizeColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the dlc value
  VerifyDLC(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && DLCColumn != -1
    //  ?
    //  RowRealValues[DLCColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, DLCColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the bp value
  VerifyBP(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    // RowsIndex > 0 && BPColumn != -1
    //  ?
    //  RowRealValues[BPColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, BPColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the uploader value
  VerifyUploader(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && UploaderColumn != -1
    //  ?
    //  RowRealValues[UploaderColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, UploaderColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the date value
  VerifyDate(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && DateColumn != -1
    //  ?
    //  RowRealValues[DateColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, DateColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the link value
  VerifyLink(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && LinkColumn != -1
    //  ?
    //  RowRealValues[LinkColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, LinkColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the tested value
  VerifyTested(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && TestedColumn != -1
    //  ?
    //  RowRealValues[TestedColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, TestedColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the info value
  VerifyInfo(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && InfoColumn != -1
    //  ?
    //  RowRealValues[InfoColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, InfoColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the DK's mirror value
  VerifyDKsMirror(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && DKsMirrorColumn != -1
    //  ?
    //  RowRealValues[DKsMirrorColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, DKsMirrorColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the update link value
  VerifyUpdateLink(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && UpdateLinkColumn != -1
    //  ?
    //  RowRealValues[UpdateLinkColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, UpdateLinkColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the submitter value
  VerifySubmitter(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && SubmitterColumn != -1
    //  ?
    //  RowRealValues[SubmitterColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, SubmitterColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the announced value
  VerifyAnnounced(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && AnnouncedColumn != -1
    //  ?
    //  RowRealValues[AnnouncedColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, AnnouncedColumn).getValue()
    //  :
    //  false;
    //)
  }

  // verifying the fixes value
  VerifyFixes(RowsIndex, RowRealValues)
  {
    // Not Implemented
    //return (
    //  RowsIndex > 0 && FixesColumn != -1
    //  ?
    //  RowRealValues[FixesColumn] ==
    //    SpreadsheetToolsObject.CheckCell(RowsIndex, FixesColumn).getValue()
    //  :
    //  false;
    //)
  }

  // applying the title value
  ApplyTitle(RowsIndex, Value, RowRealValues = null)
  {
    // Not Implemented
    //if (
    //  RowsIndex > 0
    //  && TitleColumn != -1
    //  && Value != null
    //  && GeneralToolsObject.VariableIsString(Value) == true
    //  && Value.length > 0
    //)
    //{
    //  if (dry_run == false)
    //    if (RowRealValues != null)
    //      RowRealValues[TitleColumn] = Value;
    //    else
    //      SpreadsheetToolsObject.ApplyCell(RowsIndex, TitleColumn, Value);

    //  TitleValueTrimmed = Value;
    //}
  }

  // applying the title id value
  ApplyTitleID(RowsIndex, Value, RowDisplayValues = null)
  {
    if (
      RowsIndex > 0
      && TitleIDColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowDisplayValues != null)
          RowDisplayValues[TitleIDColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, TitleIDColumn, Value);

      TitleIDValueTrimmed = Value;
    }
  }

  // applying the region value
  ApplyRegion(RowsIndex, Value, RowDisplayValues = null)
  {
    if (
      RowsIndex > 0
      && RegionColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowDisplayValues != null)
          RowDisplayValues[RegionColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, RegionColumn, Value);

      RegionValueTrimmed = Value;
    }
  }

  // applying the genre value
  ApplyGenre(RowsIndex, Value, RowDisplayValues = null)
  {
    if (
      RowsIndex > 0
      && GenreColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowDisplayValues != null)
          RowDisplayValues[GenreColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, GenreColumn, Value);

      GenreValueTrimmed = Value;
    }
  }

  // applying the base size value (the full and slower function)
  ApplyBaseSizeA(RowsIndex, FileSizeNumber, FileSizeNumberFormat, RowDisplayValues, RowRealValues = null)
  {
    BaseSizeValueTrimmed =
      this.ApplyFileSizeMethod(
        RowsIndex, BaseSizeColumn, FileSizeNumber, FileSizeNumberFormat, RowDisplayValues, RowRealValues
      );
  }

  // applying the base size value (the short and faster function)
  ApplyBaseSizeB(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && BaseSizeColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[BaseSizeColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, BaseSizeColumn, Value);

      BaseSizeValueTrimmed = Value;
    }
  }

  // applying the version value (the full and slower function)
  ApplyVersionA(RowsIndex, Value, RowRealValues = null)
  {
    VersionValueTrimmed = this.ApplyVersionMethod(RowsIndex, VersionColumn, Value, RowRealValues);
  }

  // applying the version value (the short and faster function)
  ApplyVersionB(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && VersionColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[VersionColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, VersionColumn, Value);

      VersionValueTrimmed = Value;
    }
  }

  // applying the update size value (the full and slower function)
  ApplyUpdateSizeA(RowsIndex, FileSizeNumber, FileSizeNumberFormat, RowDisplayValues, RowRealValues = null)
  {
    UpdateSizeValueTrimmed =
      this.ApplyFileSizeMethod(
        RowsIndex, UpdateSizeColumn, FileSizeNumber, FileSizeNumberFormat, RowDisplayValues, RowRealValues
      );
  }

  // applying the update size value (the short and faster function)
  ApplyUpdateSizeB(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && UpdateSizeColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[UpdateSizeColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, UpdateSizeColumn, Value);

      UpdateSizeValueTrimmed = Value;
    }
  }

  // applying the dlc value
  ApplyDLC(RowsIndex, Value, RowDisplayValues = null)
  {
    if (
      RowsIndex > 0
      && DLCColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowDisplayValues != null)
          RowDisplayValues[DLCColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, DLCColumn, Value);

      DLCValueTrimmed = Value;
    }
  }

  // applying the bp value
  ApplyBP(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && BPColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[BPColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, BPColumn, Value);

      BPValueTrimmed = Value;
    }
  }

  // applying the uploader value
  ApplyUploader(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && UploaderColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[UploaderColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, UploaderColumn, Value);

      UploaderValueTrimmed = Value;
    }
  }

  // applying the date value (the full and slower function)
  ApplyDateA(RowsIndex, Value, RowRealValues = null)
  {
    DateValueTrimmed = this.ApplyDateMethod(RowsIndex, DateColumn, Value, RowRealValues);
  }

  // applying the date value (the short and faster function)
  ApplyDateB(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && DateColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[DateColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, DateColumn, Value);

      DateValueTrimmed = Value;
    }
  }

  // applying the link value
  ApplyLink(RowsIndex, Value, RowRealValues = null, RowFormulas = null)
  {
    LinkValueTrimmed = this.ApplyLinkMethod(RowsIndex, LinkColumn, Value, RowRealValues, RowFormulas);
  }

  // applying the tested value
  ApplyTested(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && TestedColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[TestedColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, TestedColumn, Value);

      TestedValueTrimmed = Value;
    }
  }

  // applying the info value
  ApplyInfo(RowsIndex, Value, RowDisplayValues = null)
  {
    // Not Implemented
    //if (
    //  RowsIndex > 0
    //  && InfoColumn != -1
    //  && Value != null
    //  && GeneralToolsObject.VariableIsString(Value) == true
    //  && Value.length > 0
    //)
    //{
    //  if (dry_run == false)
    //    if (RowDisplayValues != null)
    //      RowDisplayValues[InfoColumn] = Value;
    //    else
    //      SpreadsheetToolsObject.ApplyCell(RowsIndex, InfoColumn, Value);
    
    //  InfoValueTrimmed = Value;
    //}
  }

  // applying the DK's mirror value
  ApplyDKsMirror(RowsIndex, Value, RowRealValues = null, RowFormulas = null)
  {
    DKsMirrorValueTrimmed = this.ApplyLinkMethod(RowsIndex, DKsMirrorColumn, Value, RowRealValues, RowFormulas);
  }

  // applying the update link value
  ApplyUpdateLink(RowsIndex, Value, RowRealValues = null, RowFormulas = null)
  {
    UpdateLinkValueTrimmed = this.ApplyLinkMethod(RowsIndex, UpdateLinkColumn, Value, RowRealValues, RowFormulas);
  }

  // applying the submitter value
  ApplySubmitter(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && SubmitterColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[SubmitterColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, SubmitterColumn, Value);

      SubmitterValueTrimmed = Value;
    }
  }

  // applying the announced value
  ApplyAnnounced(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && AnnouncedColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[AnnouncedColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, AnnouncedColumn, Value);

      AnnouncedValue = Value;
    }
  }

  // applying the fixes value
  ApplyFixes(RowsIndex, Value, RowRealValues = null)
  {
    if (
      RowsIndex > 0
      && FixesColumn != -1
      && Value != null
      && GeneralToolsObject.VariableIsString(Value) == true
      && Value.length > 0
    )
    {
      if (dry_run == false)
        if (RowRealValues != null)
          RowRealValues[FixesColumn] = Value;
        else
          SpreadsheetToolsObject.ApplyCell(RowsIndex, FixesColumn, Value);

      FixesValueTrimmed = Value;
    }
  }

  // check the spreadsheet rows
  CheckRows(
    CheckRealValues = true
    , CheckDisplayValues = true
    , CheckFormulas = true
    , StartingRow = -1
    , EndingRow = -1
  )
  {
    var FunctionResult;
    var CurrentResult = null;

    if (CheckRealValues == true || CheckDisplayValues == true || CheckFormulas == true)
      if (
        SpreadsheetToolsObject.SheetName != null
        && GeneralToolsObject.VariableIsString(SpreadsheetToolsObject.SheetName) == true
        && SpreadsheetToolsObject.SheetName.length > 0
        && SpreadsheetToolsObject.SpreadsheetReposNames.includes(SpreadsheetToolsObject.SheetName) == true
      )
      {
        var RowsRealValues = null;
        var RowsRealValuesAmount = -1;

        var MethodFound = false;

        if (CheckRealValues == true)
        {
          RowsRealValues = SpreadsheetToolsObject.CheckRows(0, StartingRow, -1, EndingRow, -1);

          if (RowsRealValues != null)
          {
            RowsRealValuesAmount = RowsRealValues.length;

            if (RowsRealValuesAmount > 0)
              MethodFound = true;
          }
        }
        else
          MethodFound = true;

        if (MethodFound == true)
        {
          var RowsDisplayValues = null;
          var RowsDisplayValuesAmount = -1;

          MethodFound = false;

          if (CheckDisplayValues == true)
          {
            RowsDisplayValues = SpreadsheetToolsObject.CheckRows(1, StartingRow, -1, EndingRow, -1);

            if (RowsDisplayValues != null)
            {
              RowsDisplayValuesAmount = RowsDisplayValues.length;

              if (RowsDisplayValuesAmount > 0)
                MethodFound = true;
            }
          }
          else
            MethodFound = true;

          if (MethodFound == true)
          {
            var RowsFormulas = null;
            var RowsFormulasAmount = -1;

            MethodFound = false;

            if (CheckFormulas == true)
            {
              RowsFormulas = SpreadsheetToolsObject.CheckRows(2, StartingRow, -1, EndingRow, -1);

              if (RowsFormulas != null)
              {
                RowsFormulasAmount = RowsFormulas.length;

                if (RowsFormulasAmount > 0)
                  MethodFound = true;
              }
            }
            else
              MethodFound = true;

            if (MethodFound == true)
            {
              MethodFound = false;

              if (CheckRealValues == true)
                if (CheckDisplayValues == true)
                  if (CheckFormulas == true)
                  {
                    if (
                      RowsRealValuesAmount == RowsDisplayValuesAmount
                      && RowsRealValuesAmount == RowsFormulasAmount
                    )
                      MethodFound = true;
                  }
                  else
                  {
                    if (RowsRealValuesAmount == RowsDisplayValuesAmount)
                      MethodFound = true;
                  }
                else
                  MethodFound = true;
              else
                if (CheckDisplayValues == true)
                  if (CheckFormulas == true)
                  {
                    if (RowsDisplayValuesAmount == RowsFormulasAmount)
                      MethodFound = true;
                  }
                  else
                    MethodFound = true;
                else
                  if (CheckFormulas == true)
                    MethodFound = true;
            }

            if (MethodFound == true)
            {
              MethodFound = false;

              CurrentResult = [];

              if (CheckRealValues == true)
              {
                CurrentResult.push(RowsRealValues);
                CurrentResult.push(RowsRealValuesAmount);
              }

              if (CheckDisplayValues == true)
              {
                CurrentResult.push(RowsDisplayValues);
                CurrentResult.push(RowsDisplayValuesAmount);
              }

              if (CheckFormulas == true)
              {
                CurrentResult.push(RowsFormulas);
                CurrentResult.push(RowsFormulasAmount);
              }
            }
          }
        }
      }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // check file size number with format based on the cell real and display value
  CheckFileSizeNumberWithFormat(RowRealValues, RowDisplayValues, Column)
  {
    var FunctionResult;
    var CurrentResult = null;

    if (
      RowRealValues != null
      && RowDisplayValues != null
      && Column != null
      && GeneralToolsObject.VariableIsNumber(Column) == true
      && GeneralToolsObject.VariableIsString(Column) == false
      && Column != -1
    )
    {
      var RowRealValuesColumns = RowRealValues;
      var RowRealValuesColumnsAmount = RowRealValuesColumns.length;

      if (RowRealValuesColumnsAmount > 0)
      {
        var RowDisplayValuesColumns = RowDisplayValues;
        var RowDisplayValuesColumnsAmount = RowDisplayValuesColumns.length;

        var MethodFound = false;

        if (RowDisplayValuesColumnsAmount > 0)
          if (RowRealValuesColumnsAmount == RowDisplayValuesColumnsAmount)
            MethodFound = true;

        if (MethodFound == true)
        {
          var CellRealValue = RowRealValues[Column];
          var CellDisplayValue = null;
          
          var FileSizeDoubleNumber = -1;
          var FileSizeMeasureType = -1;

          var FileSizeNumberFormat = null;

          var OptimalFileSizeDoubleNumber = -1;
          var OptimalFileSizeMeasureType = -1;

          MethodFound = false;

          if (CellRealValue != null)
          {
            CellDisplayValue = RowDisplayValues[Column];

            if (
              CellDisplayValue != null
              && GeneralToolsObject.VariableIsString(CellDisplayValue) == true
              && CellDisplayValue.length > 0
            )
              if (GeneralToolsObject.VariableIsNumber(CellRealValue) == false)
                MethodFound = true;
              else if (GeneralToolsObject.VariableIsNumber(CellDisplayValue) == false)
              {
                MethodFound = true;

                FileSizeDoubleNumber =
                  GeneralToolsObject.VariableIsString(CellRealValue) == true
                  ? Number(CellRealValue)
                  : CellRealValue;
              }
          }

          if (MethodFound == true)
          {
            var FileSizeMeasure = GeneralToolsObject.CheckFileSizeMeasure(CellDisplayValue, -1);

            MethodFound = false;

            if (
              FileSizeMeasure != null
              && GeneralToolsObject.VariableIsString(FileSizeMeasure) == true
              && FileSizeMeasure.length > 0
            )
            {
              FileSizeMeasureType = GeneralToolsObject.CheckFileSizeMeasureType(FileSizeMeasure);

              if (FileSizeMeasureType != -1)
              {
                var FileSizeDoubleDisplayNumber =
                  GeneralToolsObject.CheckFileSizeNumber(CellDisplayValue, FileSizeMeasure);

                if (FileSizeDoubleDisplayNumber != -1)
                {
                  // enter in case it's already in a value that's formatted
                  if (FileSizeDoubleNumber != -1)
                  {
                    var FileSizeIntegerNumber = GeneralToolsObject.FloatToInteger(FileSizeDoubleNumber, 5);
                    var FileSizeIntegerDisplayNumber = GeneralToolsObject.FloatToInteger(FileSizeDoubleDisplayNumber, 5);

                    while (FileSizeIntegerNumber > FileSizeIntegerDisplayNumber)
                    {
                      FileSizeDoubleNumber /= 1000;
                      FileSizeIntegerNumber = GeneralToolsObject.FloatToInteger(FileSizeDoubleNumber, 5);
                    }
                  }
                  else
                    FileSizeDoubleNumber = FileSizeDoubleDisplayNumber;

                  // don't accept values that are smaller than 500 KB
                  if (
                    FileSizeMeasureType > 0
                    && (FileSizeMeasureType > 1 || FileSizeDoubleNumber >= 500)
                  )
                    MethodFound = true;
                }
              }
            }

            if (MethodFound == true)
            {
              MethodFound = false;

              OptimalFileSizeDoubleNumber = FileSizeDoubleNumber;
              OptimalFileSizeMeasureType = FileSizeMeasureType;

              while (OptimalFileSizeMeasureType > 0)
              {
                OptimalFileSizeDoubleNumber *= 1024;
                OptimalFileSizeMeasureType--;
              }

              while (
                OptimalFileSizeMeasureType < MutualMethodsFileSizeMeasureType
                || OptimalFileSizeDoubleNumber >= 1024
                && OptimalFileSizeMeasureType < MutualMethodsFileSizeMeasureType + (
                  MutualMethodsFileSizeMeasureType == 0
                  ? 3
                  : 2
                )
              )
              {
                OptimalFileSizeDoubleNumber /= 1024;
                OptimalFileSizeMeasureType++;
              }

              // need to avoid numbers smaller than 1 or bigger or equal to 1000
              // because can't show numbers under 1 or between 1000 to 1023
              // losing or gaining some size along the way, for example 0.99 GB would be shown as 1 GB
              // 1023 MB would be shown as 1 GB as well, 1011 MB would be shown as 999 MB
              // at most losing or gaining 12 MB
              if (OptimalFileSizeDoubleNumber < 1 || OptimalFileSizeDoubleNumber >= 1000)
                if (
                  (
                    OptimalFileSizeDoubleNumber < 1
                    ? OptimalFileSizeDoubleNumber * 1024
                    : OptimalFileSizeDoubleNumber
                  ) >= 1000 + (1024 - 1000) / 2
                )
                  if (OptimalFileSizeDoubleNumber >= 1000)
                  {
                    if (
                      OptimalFileSizeMeasureType < MutualMethodsFileSizeMeasureType + (
                        MutualMethodsFileSizeMeasureType == 0
                        ? 3
                        : 2
                      )
                    )
                    {
                      OptimalFileSizeDoubleNumber = 1;
                      OptimalFileSizeMeasureType++;
                    }
                  }
                  else
                    OptimalFileSizeDoubleNumber = 1;
                else if (OptimalFileSizeDoubleNumber < 1)
                {
                  if (OptimalFileSizeMeasureType > MutualMethodsFileSizeMeasureType)
                  {
                    OptimalFileSizeDoubleNumber = 999.99;
                    OptimalFileSizeMeasureType--;
                  }
                }
                else
                  OptimalFileSizeDoubleNumber = 999.99;

              // don't accept values that are smaller than 500 KB
              if (
                OptimalFileSizeMeasureType > 0
                && (OptimalFileSizeMeasureType > 1 || OptimalFileSizeDoubleNumber >= 500)
              )
                MethodFound = true;
            }

            if (MethodFound == true)
            {
              var OptimalFileSizeIntegerNumber = GeneralToolsObject.FloatToInteger(OptimalFileSizeDoubleNumber, 5);

              MethodFound = false;

              FileSizeDoubleNumber = OptimalFileSizeDoubleNumber;
              FileSizeMeasureType = OptimalFileSizeMeasureType;

              while (FileSizeMeasureType < MutualMethodsFileSizeMeasureType)
              {
                FileSizeDoubleNumber /= 1000;
                FileSizeMeasureType++;
              }

              while (FileSizeMeasureType > MutualMethodsFileSizeMeasureType)
              {
                FileSizeDoubleNumber *= 1000;
                FileSizeMeasureType--;
              }

              FileSizeNumberFormat =
                OptimalFileSizeDoubleNumber != OptimalFileSizeIntegerNumber
                && (
                  FileSizeDoubleNumber > 0 && FileSizeDoubleNumber < 1
                  || OptimalFileSizeDoubleNumber > 0 && OptimalFileSizeDoubleNumber < 1
                  || FileSizeDoubleNumberFormatOnIntegerNumberFormatIfAboveMB != undefined
                  && FileSizeDoubleNumberFormatOnIntegerNumberFormatIfAboveMB != null
                  && GeneralToolsObject.VariableIsBoolean(
                    FileSizeDoubleNumberFormatOnIntegerNumberFormatIfAboveMB
                  ) == true
                  && (
                    FileSizeMeasureType > 2
                    || OptimalFileSizeMeasureType > 2
                  )
                )
                ? this.FileSizeDoubleNumberFormat
                : this.FileSizeIntegerNumberFormat;
              
              CurrentResult = [];

              CurrentResult.push(FileSizeDoubleNumber);
              CurrentResult.push(FileSizeNumberFormat);
            }
          }
        }
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // resetting the values that are being used in the script for the spreadsheet
  ResetValues()
  {
    TitleValue = null;
    TitleIDValue = null;
    RegionValue = null;
    GenreValue = null;
    BaseSizeValue = null;
    VersionValue = null;
    UpdateSizeValue = null;
    DLCValue = null;
    BPValue = null;
    UploaderValue = null;
    DateValue = null;
    LinkValue = null;
    TestedValue = null;
    InfoValue = null;
    DKsMirrorValue = null;
    UpdateLinkValue = null;
    SubmitterValue = null;
    AnnouncedValue = null;
    FixesValue = null;

    TitleValueTrimmed = null;
    TitleIDValueTrimmed = null;
    RegionValueTrimmed = null;
    GenreValueTrimmed = null;
    BaseSizeValueTrimmed = null;
    VersionValueTrimmed = null;
    UpdateSizeValueTrimmed = null;
    DLCValueTrimmed = null;
    BPValueTrimmed = null;
    UploaderValueTrimmed = null;
    DateValueTrimmed = null;
    LinkValueTrimmed = null;
    TestedValueTrimmed = null;
    InfoValueTrimmed = null;
    DKsMirrorValueTrimmed = null;
    UpdateLinkValueTrimmed = null;
    SubmitterValueTrimmed = null;
    //AnnouncedValueTrimmed = null;
    FixesValueTrimmed = null;
  }
}

// initializing the mutual methods object
function InitializeMutualMethodsObject_()
{
  if (MutualMethodsObject == null)
    MutualMethodsObject = new MutualMethodsObjectType();
}
