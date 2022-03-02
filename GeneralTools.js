// user config start

// the names that will be used for known domains that will be showed on the displayed cells
var DomainsWithNames = [];

DomainsWithNames.push({Domain: "domain.domainending", Name: "Domain"});
DomainsWithNames.push({Domain: "domain2.domainending2", Name: "Domain2"});

// formatting 4 GB as 4.00 GB instead, unlike 4 MB which gets formatted as 4 MB
var FileSizeDoubleNumberFormatOnIntegerNumberFormatIfAboveMB = true;

// user config end

var GeneralToolsObject = null;// do not delete

// whether to run in dry run, and not apply any change, used for debugging mostly
var dry_run = false;// do not delete

// the text that will be used on the displayed cells which are set as "not available", can be changed
var NAText = "N/A";// do not delete
var NATextLength = NAText.length;// do not delete

class GeneralToolsObjectType
{
  constructor()
  {
    // adding leading zeros to a number
    Number.prototype.pad = (
      function(MinNumberLength)
      {
        var NumberText = this.toString();
        var NumberTextLength = NumberText.length;

        while (NumberTextLength < MinNumberLength)
        {
          NumberText = '0' + NumberText;
          NumberTextLength++;
        }

        return NumberText;
      }
    );

    // making a new property for the array type called "clone", used for cloning an array
    Array.prototype.clone = (
      function()
      {
        return this.concat();
      }
    );

    // making a new property for the array type called "unique", used for deleting duplicates in an array
    Array.prototype.unique = (
      function()
      {
        var FunctionResult;

        var CurrentArray = this.clone();
        var CurrentArrayLength = CurrentArray.length;

        if (CurrentArrayLength > 1)
        {
          var CurrentArrayIndexA = 0;
          var CurrentArrayIndexB;

          while (CurrentArrayIndexA < CurrentArrayLength)
          {
            CurrentArrayIndexB = CurrentArrayIndexA + 1;

            while (CurrentArrayIndexB < CurrentArrayLength)
              if (CurrentArray[CurrentArrayIndexA] == CurrentArray[CurrentArrayIndexB])
              {
                CurrentArray.splice(CurrentArrayIndexB, 1);
                CurrentArrayLength--;
              }
              else
                CurrentArrayIndexB++;

            CurrentArrayIndexA++;
          }
        }

        FunctionResult = CurrentArray;

        return FunctionResult;
      }
    );
  }

  // returns true if the variable inserted is a string
  VariableIsString(Variable)
  {
    return (typeof(Variable) == "string");
  }

  // returns true if the variable inserted is a number (might be a text as well, that contains only numbers)
  VariableIsNumber(Variable)
  {
    return (isNaN(Variable) == false);
  }

  // returns true if the variable inserted is a boolean
  VariableIsBoolean(Variable)
  {
    return (typeof(Variable) == "boolean");
  }

  // returns true if the variable inserted is not N/A
  VariableIsNotNA(Variable)
  {
    var FunctionResult;
    var CurrentResult = false;

    if (
      Variable == null
      || this.VariableIsString(Variable) == false
    )
      CurrentResult = true;
    else
    {
      var VariableLength = Variable.length;

      if (VariableLength == 0)
        CurrentResult = true;
      else if (VariableLength == NATextLength || VariableLength == 3 || VariableLength == 2)
      {
        var CapitalVariable = Variable.toUpperCase();
        
        if (
          VariableLength != NATextLength
          || CapitalVariable != NAText.toUpperCase()
        )
          if (VariableLength == 3)
          {
            if (
              CapitalVariable != "N/A"
              && CapitalVariable != "N\\A"
            )
              CurrentResult = true;
          }
          else if (VariableLength == 2)
          {
            if (CapitalVariable != "NA")
              CurrentResult = true;
          }
          else
            CurrentResult = true;

      }
      else
        CurrentResult = true;
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // converts a float variable to integer, rounding up based on the accuracy parameter
  FloatToInteger(Float, Accuracy = 0)
  {
    var FunctionResult;
    var CurrentResult;

    if (Accuracy > 0)
    {
      var CurrentFloat = Float;

      for (let AccuracyIndex = 0; AccuracyIndex < Accuracy; AccuracyIndex++)
        CurrentFloat *= 10;

      CurrentResult = parseInt(Math.ceil(CurrentFloat), 10);
      
      for (let AccuracyIndex = 0; AccuracyIndex < Accuracy; AccuracyIndex++)
        CurrentResult /= 10;

      CurrentResult = parseInt(CurrentResult, 10);
    }
    else
      CurrentResult = parseInt(Float, 10);

    // fixing errors if there're any

    while (CurrentResult - Float >= 1)
      CurrentResult--;

    while (Float - CurrentResult >= 1)
      CurrentResult++;

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // gets the ASCII representation of the inserted character
  CheckCharacterASCII(Character)
  {
    return (
      Character != null
      && this.VariableIsString(Character) == true
      && Character.length == 1
      ? Character.charCodeAt(0)
      : null
    );
  }

  // sorting a dictionary
  SortDictionary(Dictionary)
  {
    var FunctionResult;
    var CurrentResult = {};

    if (Dictionary != null)
    {
      var Keys = [];
      var KeysAmount = 0;

      for (let Key in Dictionary)
      {
        Keys.push(Key);
        KeysAmount++;
      }

      if (KeysAmount > 0)
      {
        var Key;

        Keys.sort();

        for (let KeysIndex = 0; KeysIndex < KeysAmount; KeysIndex++)
        {
          Key = Keys[KeysIndex];

          CurrentResult[Key] = Dictionary[Key];
        }
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // validating that a string variable is a link
  // , a link can be valid and still not pass the validation test
  // , for example "example.com/etc" won't pass the validation test since there's no way to check for it easily
  // , if it starts with "https://" or "http://" or "www." then it would pass the validation test
  ValidateLink(Link)// https://www.example.com/etc
  {
    var FunctionResult;
    var CurrentResult = false;

    if (
      Link != null
      && this.VariableIsString(Link) == true
      && Link.length > 0
    )
    {
      var CurrentLink = Link.toUpperCase();// HTTPS://WWW.EXAMPLE.COM/ETC
      
      if (CurrentLink != null)
      {
        var CurrentLinkLength = CurrentLink.length;

        if (
          CurrentLinkLength > 8
          && CurrentLink.substr(0, 8) == "HTTPS://"// HTTPS://WWW.EXAMPLE.COM/ETC
          || CurrentLinkLength > 7
          && CurrentLink.substr(0, 7) == "HTTP://"// HTTP://WWW.EXAMPLE.COM/ETC
          || CurrentLinkLength > 4
          && CurrentLink.substr(0, 4) == "WWW."// WWW.EXAMPLE.COM/ETC
        )
          CurrentResult = true;
      }
    }

    FunctionResult = CurrentResult;
    
    return FunctionResult;
  }

  // checking the link domain of a link, for example
  // https://www.example.com/etc -> example.com
  CheckLinkDomain(Link)// https://www.example.com/etc
  {
    var LinkDomain = null;

    if (
      Link != null
      && this.VariableIsString(Link) == true
      && Link.length > 0
    )
    {
      var CurrentLinkDomain = Link;

      // stripping https://
      if (
        CurrentLinkDomain.length > 8
        && CurrentLinkDomain.substr(0, 8).toUpperCase() == "HTTPS://"// HTTPS://WWW.EXAMPLE.COM/ETC
      )
        CurrentLinkDomain = CurrentLinkDomain.substr(8);// www.example.com/etc
      // stripping http://
      else if (
        CurrentLinkDomain.length > 7
        && CurrentLinkDomain.substr(0, 7).toUpperCase() == "HTTP://"// HTTP://WWW.EXAMPLE.COM/ETC
      )
        CurrentLinkDomain = CurrentLinkDomain.substr(7);// www.example.com/etc

      // stripping www.
      if (
        CurrentLinkDomain.length > 4
        && CurrentLinkDomain.substr(0, 4).toUpperCase() == "WWW."// WWW.EXAMPLE.COM/ETC
      )
        CurrentLinkDomain = CurrentLinkDomain.substr(4);// example.com/etc

      // stripping forward slashes
      if (CurrentLinkDomain.length > 1)// the link domain can't consist of only a slash
      {
        var CurrentLinkDomainSlashLocation = CurrentLinkDomain.search('/');// example.com/etc

        if (CurrentLinkDomainSlashLocation > 0)
          CurrentLinkDomain = CurrentLinkDomain.substr(0, CurrentLinkDomainSlashLocation);// example.com
      }

      if (CurrentLinkDomain.length > 0)
        LinkDomain = CurrentLinkDomain;// example.com
    }

    return LinkDomain;
  }

  // checking the link domain name of a link domain, for example
  // example.com -> Example
  CheckLinkDomainName(LinkDomain)// example.com
  {
    var LinkDomainName = null;

    if (
      LinkDomain != null
      && this.VariableIsString(LinkDomain) == true
      && LinkDomain.length > 0
    )
      if (
        DomainsWithNames != undefined
        && DomainsWithNames != null
      )
      {
        var DomainsWithNamesAmount = DomainsWithNames.length;

        if (DomainsWithNamesAmount > 0)
        {
          var DomainWithName;
          var DomainWithNameDomain;
          var DomainWithNameName;

          var LinkDomainCapitalValue = LinkDomain.toUpperCase();// EXAMPLE.COM

          for (let DomainsWithNamesIndex = 0; DomainsWithNamesIndex < DomainsWithNamesAmount; DomainsWithNamesIndex++)
          {
            DomainWithName = DomainsWithNames[DomainsWithNamesIndex];

            if (DomainWithName != null)
              if (
                DomainWithName.Domain != undefined
                && DomainWithName.Domain != null
                && DomainWithName.Name != undefined
                && DomainWithName.Name != null
                && this.VariableIsString(DomainWithName.Domain) == true
                && DomainWithName.Domain.length > 0
                && this.VariableIsString(DomainWithName.Name) == true
                && DomainWithName.Name.length > 0
              )
              {
                DomainWithNameDomain = DomainWithName.Domain;// example.com

                if (LinkDomainCapitalValue == DomainWithNameDomain.toUpperCase())// EXAMPLE.COM
                {
                  DomainWithNameName = DomainWithName.Name;// Example

                  LinkDomainName = DomainWithNameName;// Example

                  break;
                }
              }
          }
        }
      }

    return LinkDomainName;
  }

  // checking the file size measure type based on the file size measure
  CheckFileSizeMeasureType(FileSizeMeasure)
  {
    var FileSizeMeasureType = -1;

    if (
      FileSizeMeasure != null
      && this.VariableIsString(FileSizeMeasure) == true
      && FileSizeMeasure.length > 0
    )
      switch (FileSizeMeasure.toUpperCase())
      {
        // Bytes
        case "B":
          FileSizeMeasureType = 0;

          break;

        // Kilobytes
        case "K":
        case "KB":
          FileSizeMeasureType = 1;

          break;

        // Megabytes
        case "M":
        case "MB":
          FileSizeMeasureType = 2;

          break;

        // Gigabytes
        case "G":
        case "GB":
          FileSizeMeasureType = 3;

          break;

        // Terabytes
        case "T":
        case "TB":
          FileSizeMeasureType = 4;

          break;

        // Petabytes
        case "P":
        case "PB":
          FileSizeMeasureType = 5;

          break;

        // Exabytes
        case "E":
        case "EB":
          FileSizeMeasureType = 6;

          break;
      }

    return FileSizeMeasureType;
  }

  // checking the file size measure based on the file size text or the file size measure type or both
  CheckFileSizeMeasure(FileSizeText = null, FileSizeMeasureType = -1)
  {
    var FileSizeMeasure = null;

    var CurrentFileSizeMeasureType =
      FileSizeMeasureType != null
      && this.VariableIsNumber(FileSizeMeasureType) == true
      && this.VariableIsString(FileSizeMeasureType) == false
      ? FileSizeMeasureType
      : -1;

    if (CurrentFileSizeMeasureType == -1)
      if (
        FileSizeText != null
        && this.VariableIsString(FileSizeText) == true
      )
      {
        var FileSizeTextLength = FileSizeText.length;

        if (FileSizeTextLength > 1)// need a number before the file size measure
        {
          var FileSizeTextCharacter;
          var FileSizeTextCharacterASCII;

          var aCharacterASCII = this.CheckCharacterASCII('a');
          var zCharacterASCII = this.CheckCharacterASCII('z');
          var ACharacterASCII = this.CheckCharacterASCII('A');
          var ZCharacterASCII = this.CheckCharacterASCII('Z');

          for (
            // need a number before it, so starting from 1
            let FileSizeTextIndex = 1;
            FileSizeTextIndex < FileSizeTextLength;
            FileSizeTextIndex++
          )
          {
            FileSizeTextCharacter = FileSizeText[FileSizeTextIndex];
            FileSizeTextCharacterASCII = this.CheckCharacterASCII(FileSizeTextCharacter);

            if (
              FileSizeTextCharacterASCII >= aCharacterASCII && FileSizeTextCharacterASCII <= zCharacterASCII
              || FileSizeTextCharacterASCII >= ACharacterASCII && FileSizeTextCharacterASCII <= ZCharacterASCII
            )
            {
              // converting the assumable file size measure to its type is useful both in readability
              // and in securing that the file size measure is valid
              // , because if it isn't then the current file size measure type would be set as -1
              CurrentFileSizeMeasureType = this.CheckFileSizeMeasureType(FileSizeText.substr(FileSizeTextIndex));

              break;
            }
          }
        }
      }

    if (CurrentFileSizeMeasureType != -1)
      switch (CurrentFileSizeMeasureType)
      {
        case 0:
          FileSizeMeasure = "B";// Bytes

          break;

        case 1:
          FileSizeMeasure = "KB";// Kilobytes

          break;

        case 2:
          FileSizeMeasure = "MB";// Megabytes

          break;

        case 3:
          FileSizeMeasure = "GB";// Gigabytes

          break;

        case 4:
          FileSizeMeasure = "TB";// Terabytes

          break;

        case 5:
          FileSizeMeasure = "PB";// Petabytes

          break;

        case 6:
          FileSizeMeasure = "EB";// Exabytes

          break;
      }

    return FileSizeMeasure;
  }

  // checking the file size number based on the file size text and optional to also input the file size measure
  // in order to make the process go faster
  CheckFileSizeNumber(FileSizeText, FileSizeMeasure = null)
  {
    var FileSizeNumber = -1;

    if (
      FileSizeText != null
      && this.VariableIsString(FileSizeText) == true
    )
    {
      var FileSizeTextLength = FileSizeText.length;

      if (FileSizeTextLength > 1)// need a number before the file size measure
      {
        var FileSizeMeasureLength =
          FileSizeMeasure != null
          && this.VariableIsString(FileSizeMeasure) == true
          ? FileSizeMeasure.length
          : 0;

        if (FileSizeMeasureLength == 0)
        {
          var NewFileSizeMeasure = this.CheckFileSizeMeasure(FileSizeText, -1);

          FileSizeMeasureLength =
            NewFileSizeMeasure != null
            && this.VariableIsString(NewFileSizeMeasure) == true
            ? NewFileSizeMeasure.length
            : 0;
        }

        if (
          FileSizeMeasureLength > 0
          && FileSizeTextLength - FileSizeMeasureLength > 0// need a number before the file size measure
        )
        {
          var FileSizeTextSubText = FileSizeText.substr(0, FileSizeTextLength - FileSizeMeasureLength);

          if (
            this.VariableIsNumber(FileSizeTextSubText) == true
            && Number(FileSizeTextSubText) >= 0
          )
            FileSizeNumber = Number(FileSizeTextSubText);
        }
      }
    }

    return FileSizeNumber;
  }

  // checking the optimal file size text based on the file size text and optional to also input the target file size measure
  // which will try to match the file size measure with it (if inputted)
  CheckOptimalFileSizeText(FileSizeText, TargetFileSizeMeasure = null)
  {
    var OptimalFileSizeText = null;

    var FileSizeMeasure = null;
    var FileSizeMeasureType = -1;

    var MethodFound = false;

    if (
      FileSizeText != null
      && this.VariableIsString(FileSizeText) == true
      && FileSizeText.length > 1// need a number before the file size measure
    )
    {
      FileSizeMeasure = this.CheckFileSizeMeasure(FileSizeText, -1);

      if (
        FileSizeMeasure != null
        && this.VariableIsString(FileSizeMeasure) == true
        && FileSizeMeasure.length > 0
      )
      {
        FileSizeMeasureType = this.CheckFileSizeMeasureType(FileSizeMeasure);

        if (FileSizeMeasureType != -1)
          MethodFound = true;
      }
    }

    if (MethodFound == true)
    {
      var MinFileSizeMeasureType = -1;
      var MaxFileSizeMeasureType = -1;

      MethodFound = false;

      if (
        TargetFileSizeMeasure != null
        && this.VariableIsString(TargetFileSizeMeasure) == true
        && TargetFileSizeMeasure.length > 0
      )
      {
        var TargetFileSizeMeasureType = this.CheckFileSizeMeasureType(TargetFileSizeMeasure);

        MethodFound = true;

        if (TargetFileSizeMeasureType != -1)
        {
          MinFileSizeMeasureType = TargetFileSizeMeasureType;
          MaxFileSizeMeasureType = TargetFileSizeMeasureType;
        }
      }

      if (MethodFound == true)
        MethodFound = false;
      else
      {
        MinFileSizeMeasureType = 0;// Bytes
        MaxFileSizeMeasureType = 6;// Exabytes
      }

      if (MinFileSizeMeasureType != -1 && MaxFileSizeMeasureType != -1)
      {
        var FileSizeDoubleNumber = this.CheckFileSizeNumber(FileSizeText, FileSizeMeasure);

        if (FileSizeDoubleNumber != -1)
        {
          var FileSizeIntegerNumber = this.FloatToInteger(FileSizeDoubleNumber, 5);

          while (FileSizeDoubleNumber < 1 && FileSizeMeasureType > MinFileSizeMeasureType)
          {
            FileSizeDoubleNumber *= 1024;
            FileSizeMeasureType--;
          }

          while (FileSizeDoubleNumber >= 1024 && FileSizeMeasureType < MaxFileSizeMeasureType)
          {
            FileSizeDoubleNumber /= 1024;
            FileSizeMeasureType++;
          }

          FileSizeMeasure = this.CheckFileSizeMeasure(null, FileSizeMeasureType);

          if (
            FileSizeMeasure != null
            && this.VariableIsString(FileSizeMeasure) == true
            && FileSizeMeasure.length > 0
          )
            OptimalFileSizeText =
              (
                FileSizeDoubleNumber != FileSizeIntegerNumber
                && (
                  FileSizeDoubleNumber > 0 && FileSizeDoubleNumber < 1
                  || FileSizeDoubleNumberFormatOnIntegerNumberFormatIfAboveMB != undefined
                  && FileSizeDoubleNumberFormatOnIntegerNumberFormatIfAboveMB != null
                  && this.VariableIsBoolean(FileSizeDoubleNumberFormatOnIntegerNumberFormatIfAboveMB) == true
                  && FileSizeMeasureType > 2
                )
                ? (this.FloatToInteger(FileSizeDoubleNumber * 100, 5) / 100).toFixed(2)
                : FileSizeIntegerNumber.toString()
              )
              + ' '
              + FileSizeMeasure;
        }
      }
    }

    return OptimalFileSizeText;
  }
}

// initializes the general tools object
function InitializeGeneralToolsObject_()
{
  if (GeneralToolsObject == null)
    GeneralToolsObject = new GeneralToolsObjectType();
}
