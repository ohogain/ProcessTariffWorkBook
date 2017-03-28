//---------
// <copyright file="ValidateData.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
//---------

using System;
using System.CodeDom.Compiler;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace ProcessTariffWorkbook
{
  public static class ValidateData
  {
    public static void PreRegExDataRecordValidate()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "PreRegExDataRecordValidate() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "PreRegExDataRecordValidate()");

      CheckPricesAreInCorrectFormat();
      CheckForDestinationTypes();
      CheckTablesForDefaultValue();
      CheckRoundingForIncorrectEntry();
      CheckTimeSchemeForIncorrectEntry();
      CheckIfMinCostAndRate4SubseqAreSame();
      CheckForFreephone();
      CheckIfFreephoneIsFree();
      CheckGrouping();
      CheckIntervalLengthsGreaterOrEqualToZero();
      CheckUsingCustomerNames();
      CheckMinimumIntervals();
      CheckMinimumDigits();
      CheckCutOffDuration();
      CheckMultiLevelEnabled();
      CheckAllSchemes();
      CheckDialTime();
      CheckMinimumTime();
      CheckIntervalsAtInitialCostGreaterOrEqualToZero();
      CheckDestinationTypesNames();
      CheckTableNames();

      Console.WriteLine("ValidateData".PadRight(30, '.') + "PreRegExDataRecordValidate() -- finished");
    }
    public static void PostRegExDataRecordValidate()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "PostRegExDataRecordValidate() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "PostRegExDataRecordValidate()");
      CheckForDuplicateBands();        
      CheckForNonUniqueGroupBands();      
      CheckSourceDestinationBandsPresentInPrefixBands();
      ErrorProcessing.WriteOutGroupBandsToErrorLog();
      ErrorProcessing.WriteToErrorlogIfMinCostAnd4ThRateSamePrice();      
      CheckIfAllMatrixBandsUsed();
      CheckForNonMatchingNames();
      CheckChargingType();
      Console.WriteLine("ValidateData".PadRight(30, '.') + "PostRegExDataRecordValidate() -- finished");
    }
    public static bool CheckIfInteger(string value)
    {
      int result = 0;
      bool isInt = int.TryParse(value, out result);
      return isInt;
    }
    public static string CreateDate()
    {      
      return $"{DateTime.Now:dd-MMM-yyyy}";
    }
    public static string CapitaliseWord(string word)
    {
      TextInfo txt = new CultureInfo("en-GB", false).TextInfo;
      return txt.ToTitleCase(word.ToLower());
    }
    public static void CheckForCommasInPrices(string sComma)
    {
      if (sComma.Contains(","))
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckForCommasInPrices()");
        StaticVariable.Errors.Add("ParseInputFile:ReadXLSXFileIntoList()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One of the prices contains a comma. ");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " --> " + sComma);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
    }
    public static string SetToFourDecimalPlaces(string value)
    {
      string result = String.Empty;
      value = value.Trim();
      try
      {
        var parsedValue = 0.0;
        if (Double.TryParse(value, out parsedValue))
        {
          double dValue = parsedValue;
          result = dValue.ToString("0.0000");
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ValidateData::SetToFourDecimalPlaces()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Failed to set price to 4 decimal places:");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      return result;
    }
    public static string AdjustRoundingValue(string value)
    {      
      string result = string.Empty;

      if (value.ToLower().Equals("yes") || value.ToLower().Equals("y") || value.ToLower().Equals("1") || value.ToLower().Equals("roundup") || value.ToLower().Equals("round up"))
      {
        result = "No Round";
      }
      else if (value.ToLower().Equals("no") || value.ToLower().Equals("n") || value.ToLower().Equals("3") || value.ToLower().Equals("exact") || value.ToLower().Equals("noround") || value.ToLower().Equals("no round"))
      {
        result = "Exact";
      }
      else
      {
        result = "Rounding";
      }  
      return result;
    }
    private static void CheckPricesAreInCorrectFormat()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckPricesAreInCorrectFormat() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckPricesAreInCorrectFormat()");
      List<string> resultsList = new List<string>();

      var query =        
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new { dr.CustomerPrefixName, dr.CustomerFirstInitialRate, dr.CustomerFirstSubseqRate, dr.CustomerSecondInitialRate, dr.CustomerSecondSubseqRate, 
          dr.CustomerThirdInitialRate, dr.CustomerThirdSubseqRate, dr.CustomerFourthInitialRate, dr.CustomerFourthSubseqRate, dr.CustomerMinCharge, 
          dr.CustomerConnectionCost, dr.ChargingType };

      foreach (var tok in query)
      {
        if (tok.ChargingType.ToUpper().Equals(Constants.Duration))
        {
          resultsList.Add(tok.CustomerFirstInitialRate);
          resultsList.Add(tok.CustomerFirstSubseqRate);
          resultsList.Add(tok.CustomerSecondInitialRate);
          resultsList.Add(tok.CustomerSecondSubseqRate);
          resultsList.Add(tok.CustomerThirdInitialRate);
          resultsList.Add(tok.CustomerThirdSubseqRate);
          resultsList.Add(tok.CustomerFourthInitialRate);
          resultsList.Add(tok.CustomerFourthSubseqRate);
          resultsList.Add(tok.CustomerMinCharge);
          resultsList.Add(tok.CustomerConnectionCost);
          try
          {
            foreach (string price in resultsList)
            {
              if (CheckIfDouble(price)) continue;
              StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One of the prices is not a double. ");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok.CustomerPrefixName + " --> " + price);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
          catch (Exception e)
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        else if (tok.ChargingType.ToUpper().Equals(Constants.Capped))
        {
          double parsedDoubleValue = 0.0;
          int parsedIntValue = 0;
          if (!double.TryParse(tok.CustomerFirstInitialRate, out parsedDoubleValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Price per minute is not in correct format. it must be a double: " + tok.CustomerFirstInitialRate);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (!double.TryParse(tok.CustomerFirstSubseqRate, out parsedDoubleValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Capped Price is not in correct format. it must be a double: " + tok.CustomerFirstSubseqRate);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (!int.TryParse(tok.CustomerSecondInitialRate, out parsedIntValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Capped Time is not in correct format. it must be a int. time in minutes: " + tok.CustomerSecondInitialRate);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (!double.TryParse(tok.CustomerMinCharge, out parsedDoubleValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Minimum Cost is not in correct format. it must be a double. " + tok.CustomerMinCharge);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (!double.TryParse(tok.CustomerConnectionCost, out parsedDoubleValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Connection Cost is not in correct format. it must be a double. " + tok.CustomerConnectionCost);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        else if (tok.ChargingType.ToUpper().Equals(Constants.Pulse))
        {
          const int minimumPulseLength = 1;
          double parsedDoubleValue = 0.0;
          int parsedIntValue = 0;
          if (!int.TryParse(tok.CustomerFirstInitialRate, out parsedIntValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Pulse Length given in decimal format must be multipled by 100. it must be changed to an int: " + tok.CustomerFirstInitialRate);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (parsedIntValue / 100 < minimumPulseLength)
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Pulse Length is too short: " + tok.CustomerFirstInitialRate);
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "It may not have been multipled by 100. It must be changed to an int from a decimal: ");
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (!int.TryParse(tok.CustomerFirstSubseqRate, out parsedIntValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Pulse Unit is not in correct format. it must be a int, normally 1: " + tok.CustomerFirstSubseqRate);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (!double.TryParse(tok.CustomerMinCharge, out parsedDoubleValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Minimum Cost is not in correct format. it must be a double: " + tok.CustomerMinCharge);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (!double.TryParse(tok.CustomerConnectionCost, out parsedDoubleValue))
          {
            StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Connection Cost is not in correct format. it must be a double: " + tok.CustomerConnectionCost);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }          
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckPricesAreInCorrectFormat() -- finished");
    }
    public static bool CheckIfDouble(string sValue)
    {
      double result = 0;
      bool isDouble = double.TryParse(sValue, out result);
      return isDouble;
    }    
    private static void CheckForDestinationTypes()
    {
      // This function may need to be removed as V6 does not use hard coded destination types.
      // this function will work for V5 however you may have to remove it for V6 if you need an NDS to test.
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForDestinationTypes() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForDestinationTypes()");      
      if (StaticVariable.ExportNds.ToUpper().Equals("TRUE"))
      {
        var query =        
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        where !dr.CustomerDestinationType.ToUpper().Equals("LOCAL") && !dr.CustomerDestinationType.ToUpper().Equals("NATIONAL") &&
            !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL") && !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL MOBILE") &&
            !dr.CustomerDestinationType.ToUpper().Equals("SERVICES") && !dr.CustomerDestinationType.ToUpper().Equals("OTHER") &&
            !dr.CustomerDestinationType.ToUpper().Equals("MOBILE")
          select new {dr.CustomerDestinationType, dr.CustomerPrefixName};

        if (query.Any())
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckForDestinationTypes()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "This destination type is invalid for V5 RingMaster. ");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "It must be either 'Local', 'National', International', 'International Mobile', 'Mobile', 'Services' or 'Other'");
          foreach (var item in query)
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + item.CustomerPrefixName + " - " + item.CustomerDestinationType);            
          }
          StaticVariable.Errors.Add(Environment.NewLine + Constants.FiveSpacesPadding + "Comment out 'ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();' in method to supress killing program");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }            
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForDestinationTypes() -- finished");
    }
    private static void CheckTablesForDefaultValue()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTablesForDefaultValue()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTablesForDefaultValue()");       
      const int tableName = 0;
      const int destinationTable = 3;
      List<string> tmpTableLinksList = new List<string>();
      foreach (string tok in StaticVariable.TableLinks)
      {
        string[] tmpAry = tok.Split('\t');
        tmpTableLinksList.Add(tmpAry[tableName].Trim());
        tmpTableLinksList.Add(tmpAry[destinationTable].Trim());
      }
      tmpTableLinksList = tmpTableLinksList.Distinct().ToList();

      var prefixes =
         from db in StaticVariable.PrefixNumbersRecord
         where db.PrefixNumber.Equals("?")
         select db.TableName;

      var enumerable = prefixes as IList<string> ?? prefixes.ToList();
      var extraTablesInPrefixes = enumerable.Except(tmpTableLinksList).ToList();
      var extraTablesInTableLinks = tmpTableLinksList.Except(enumerable).ToList();

      if (extraTablesInPrefixes.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckTablesForDefaultValue()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "A default prefix does not have a entry in prefix links header. \nIs the prefix file missing?");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Is there an INI file that is not required?");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + extraTablesInPrefixes[0]);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      if (extraTablesInTableLinks.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckTablesForDefaultValue()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "A table in prefix links header file does not have a default prefix - ?");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Or an ini file may be missing for that prefix link.");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTablesForDefaultValue()-- finished");
    }
    private static void CheckRoundingForIncorrectEntry()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckRoundingForIncorrectEntry() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckRoundingForIncorrectEntry()");
      List<string> errors = new List<string>();       
      try
      {
        var query =        
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        select new { db.CustomerRounding, db.CustomerPrefixName, db.ChargingType };

        foreach (var q in query)
        {
          if (q.ChargingType.ToUpper().Equals("PULSE")) continue;
          string custRounding = q.CustomerRounding.ToUpper();
          if (
            !(custRounding.Equals("YES") || custRounding.Equals("1") || custRounding.Equals("Y") ||
              custRounding.Equals("ROUNDUP") || custRounding.Equals("ROUND UP") ||
              custRounding.Equals("NO") || custRounding.Equals("3") || custRounding.Equals("N") ||
              custRounding.Equals("EXACT") || custRounding.Equals("NO ROUND") || custRounding.Equals("NOROUND")))
          {
            errors.Add(q.CustomerPrefixName + " is --> " + custRounding);
          }
        }
        if (errors.Any())
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckRoundingForIncorrectEntry()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Rounding Values are incorrect for these destinations.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The must be 'Yes', 'Y', 'ROUND UP', 'ROUNDUP' or '1' for round up and 'No', 'N', 'EXACT', 'NOROUND', 'NO ROUND' or '3' for no round");
          foreach (string error in errors)
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + error);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ValidateData::CheckRounding()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Exception Message :: " + e.Message);
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckRounding() -- finished");
    }
    private static void CheckTimeSchemeForIncorrectEntry()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTimeSchemeForIncorrectEntry() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTimeSchemeForIncorrectEntry()");
      List<string> errors = new List<string>();
      const int timeScheme = 0;
      bool found = false;
      
      var queryCustomerTimeSchemes =        
      from DataRecord db in StaticVariable.CustomerDetailsDataRecord
      select new { db.CustomerTimeScheme, db.CustomerPrefixName };        

      foreach (var q in queryCustomerTimeSchemes)
      {
        foreach (var timeSchemeName in StaticVariable.TimeSchemes)
        {
          string[] timeschemes = timeSchemeName.Split('\t');
          if (!q.CustomerTimeScheme.ToUpper().Equals(timeschemes[timeScheme].ToUpper())) continue;
          found = true;
          break;
        }
        if (!found)
        {
          errors.Add(q.CustomerPrefixName + " is --> " + q.CustomerTimeScheme);
        }
        found = false;
      }
      if (errors.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckTimeSchemeForIncorrectEntry()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Time Scheme Values are incorrect for these destinations.");
        foreach (string error in errors)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + error);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }     
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckRounding() -- finished");
    }
    private static void CheckIfMinCostAndRate4SubseqAreSame()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "MinCostAndRate4SubseqAreSame() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "MinCostAndRate4SubseqAreSame()");
      List<string> pricesTheSame = new List<string>();

      var query =        
        from DataRecord drm in StaticVariable.CustomerDetailsDataRecord
        where drm.CustomerFourthSubseqRate.Equals(drm.CustomerMinCharge) && Convert.ToDouble(drm.CustomerMinCharge) > 0.0
        select new { drm.CustomerPrefixName, drm.CustomerMinCharge, drm.CustomerFourthSubseqRate };

      foreach (var q in query)
      {
        pricesTheSame.Add(q.CustomerPrefixName + ": MinCost = " + q.CustomerMinCharge + ", 4th Rate Subsequential = " + q.CustomerFourthSubseqRate);
      }
      if (pricesTheSame.Any())
      {
        pricesTheSame.Sort();
        StaticVariable.Errors.Add(Environment.NewLine + "Minimum Cost is the same price as the 4th Rate Subsequent price. Check it out.");
        foreach (string s in pricesTheSame)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + s);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "MinCostAndRate4SubseqAreSame() -- finished");
    }
    private static void CheckForFreephone()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForFreephone() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForFreephone()");
      var result =        
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        where db.CustomerPrefixName.ToUpper().Contains("FREE") || db.StdBand.ToUpper().Equals("FREE") ||
              db.CustomerGroupBand.ToUpper().Equals("FREE") || db.CustomerGroupBand.ToUpper().Equals("TOLL") ||
              db.CustomerPrefixName.ToUpper().Contains("GRAT") || db.StdBand.ToUpper().Equals("GRAT") || 
              db.CustomerPrefixName.ToUpper().Contains("TOLL") || db.StdBand.ToUpper().Equals("TOLL") 
              && !db.StdPrefixName.ToUpper().Contains("INT")
        select db.StdBand;

      if (result.Count().Equals(0))
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckForFreephone()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There is no entry for Freephone.");
        Console.WriteLine("There is no entry for Freephone.............");
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForFreephone() -- finished");
    }
    private static void CheckIfFreephoneIsFree()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfFreephoneIsFree() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIfFreephoneIsFree()");
      bool bFound = false;
      string custName = string.Empty;

      var results =        
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        where db.CustomerPrefixName.ToUpper().Contains("FREE") || db.StdBand.ToUpper().Equals("FREE") ||
              db.CustomerGroupBand.ToUpper().Equals("FREE") || db.CustomerGroupBand.ToUpper().Equals("TOLL") ||
              db.CustomerPrefixName.ToUpper().Contains("GRAT") || db.StdBand.ToUpper().Equals("GRAT") ||
              db.CustomerPrefixName.ToUpper().Contains("TOLL") || db.StdBand.ToUpper().Equals("TOLL")
              && !db.StdPrefixName.ToUpper().Contains("INT")
        select new
        {
          db.CustomerPrefixName,
          db.StdBand,
          db.StdPrefixName,
          db.CustomerFirstInitialRate,
          db.CustomerFirstSubseqRate,
          db.CustomerSecondInitialRate,
          db.CustomerSecondSubseqRate,
          db.CustomerThirdInitialRate,
          db.CustomerThirdSubseqRate,
          db.CustomerFourthInitialRate,
          db.CustomerFourthSubseqRate,
          db.CustomerMinCharge,
          db.CustomerConnectionCost
        };

      foreach (var result in results)
      {
        custName = result.CustomerPrefixName;
        if (CheckIfPriceZero(result.CustomerFirstInitialRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerFirstSubseqRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerSecondInitialRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerSecondSubseqRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerThirdInitialRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerThirdSubseqRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerFourthInitialRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerFourthSubseqRate)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerMinCharge)) { bFound = true; }
        else if (CheckIfPriceZero(result.CustomerConnectionCost)) { bFound = true; }
      }
      if (bFound)
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckIfFreephoneIsFree()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Freephone is not zero priced. - " + custName);
        Console.WriteLine("Freephone is not zero............");
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfFreephoneIsFree() -- finished");
    }
    private static bool CheckIfPriceZero(string sValue)
    {     
      bool notZero = false;
      foreach (char c in sValue)
      {
        if (!c.Equals('0') && !c.Equals('.'))
        {
          notZero = true;
          break;
        }
      }
      return notZero;
    }
    private static void CheckGrouping()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckGrouping() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckGrouping()");                  
      #region using group bands
      List<string> usingGroupBands = new List<string>();
      var queryUsingGroupBands =
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        where !db.CustomerUsingGroupBands.ToUpper().Equals("TRUE") && !db.CustomerUsingGroupBands.ToUpper().Equals("FALSE")
        select new { db.CustomerUsingGroupBands, db.CustomerPrefixName };

      foreach (var q in queryUsingGroupBands)
      {
        usingGroupBands.Add(q.CustomerPrefixName + " --> " + q.CustomerUsingGroupBands);       
      }
      if (usingGroupBands.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckGrouping()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Using Group Bands field are incorrect for these destinations.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The must be 'TRUE' or 'FALSE'");

        foreach (string s in usingGroupBands)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + s);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      #endregion
      #region group band
      List<string> groupBands = new List<string>();
      var queryGroupBand =
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        select new { db.CustomerGroupBand, db.CustomerPrefixName, db.CustomerUsingGroupBands };

      foreach (var q in queryGroupBand)
      {
        if (q.CustomerGroupBand.Length > Constants.V5Tc2BandLengthLimit && StaticVariable.ExportNds.ToUpper().Equals("TRUE") && q.CustomerUsingGroupBands.Equals("TRUE"))
        {
          groupBands.Add(q.CustomerPrefixName + " --> " + q.CustomerGroupBand);
        }
      }
      if (groupBands.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckGrouping()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Group Band field are too long for these destinations.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The must be no greater than 4 chars long.");

        foreach (string band in groupBands)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + band);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      # endregion
      # region group band description
      List<string> groupBandDescriptions = new List<string>();
      var queryGroupBandDescription =
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        select new { db.CustomerGroupBandDescription, db.CustomerPrefixName, db.CustomerUsingGroupBands };

      foreach (var q in queryGroupBandDescription)
      {
        if (q.CustomerGroupBandDescription.Length > Constants.V5Tc2BandDescriptionLength && StaticVariable.ExportNds.Equals("TRUE") && q.CustomerUsingGroupBands.Equals("TRUE"))
        {
          groupBandDescriptions.Add(q.CustomerPrefixName + " --> " + q.CustomerGroupBandDescription);
        }
      }
      if (groupBandDescriptions.Any()) 
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckGrouping()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Group Band Description field are too long for these destinations.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The must be no greater than 20 chars long.");

        foreach (string s in groupBandDescriptions)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + s);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      # endregion
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckGrouping() -- finished");
    }
    public static void CheckIntervalLengthsGreaterOrEqualToZero()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalLengthsGreaterOrEqualToZero()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIntervalLengthsGreaterOrEqualToZero()");
      List<string> results = new List<string>();
      List<string> errors = new List<string>();      
      int nValue = 0;
      const string defaultIntervalLength = "60";
      
      var queryIntervalLengthGreaterThanZero =
        (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
         where !dr.ChargingType.ToUpper().Equals("PULSE")
         select new { dr.CustomerInitialIntervalLength, dr.CustomerSubsequentIntervalLength }).Distinct();
          
      foreach (var tok in queryIntervalLengthGreaterThanZero)
      {
        results.Add(tok.CustomerInitialIntervalLength);
        results.Add(tok.CustomerSubsequentIntervalLength);
      }
      results = results.Distinct().ToList();
     
      foreach (string intervalLength in results)
      {
        int nParsedValue = 0;
        if (int.TryParse(intervalLength, out nParsedValue))
        {
          nValue = nParsedValue;
        }
        else
        {
          errors.Add(intervalLength);
        }
        if (nValue <= 0)
        {
          errors.Add(intervalLength);
        }
      }
      
      if (errors.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckIntervalLengthsGreaterOrEqualToZero()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One or more of the interval lengths is not an integer or is less than 1. ");
        foreach (string error in errors)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The illegal value is - " + error);
          var query =
            from qry in StaticVariable.CustomerDetailsDataRecord
            where qry.CustomerInitialIntervalLength.Equals(error) || qry.CustomerSubsequentIntervalLength.Equals(error)
            select new {qry.CustomerPrefixName, qry.CustomerInitialIntervalLength , qry.CustomerSubsequentIntervalLength };

          foreach (var q in query)
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + q.CustomerPrefixName + " - ");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "initial interval Length    = " + q.CustomerInitialIntervalLength );
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "subsequent interval length = " + q.CustomerSubsequentIntervalLength );
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      var queryIntervalLengthValues =
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        where !dr.ChargingType.ToUpper().Equals("PULSE") && (!dr.CustomerInitialIntervalLength.Equals(defaultIntervalLength) || !dr.CustomerSubsequentIntervalLength.Equals(defaultIntervalLength))
        select new { dr.CustomerInitialIntervalLength, dr.CustomerSubsequentIntervalLength, dr.CustomerPrefixName };    
          
      if (queryIntervalLengthValues.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckIntervalLengthsGreaterOrEqualToZero()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The interval lengths listed below are not the default 60 seconds. ");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "This may be correct. ");
        foreach (var interval in queryIntervalLengthValues)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + interval.CustomerPrefixName + " - ");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "initial interval Length    = " + interval.CustomerInitialIntervalLength);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "subsequent interval length = " + interval.CustomerSubsequentIntervalLength);
        }
      }    
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalLengthsGreaterOrEqualToZero()-- finished");
    }
    public static void CheckUsingCustomerNames()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckUsingCustomerNames()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckUsingCustomerNames()");

      var query =
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        where !dr.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") && !dr.CustomerUsingCustomerNames.ToUpper().Equals("FALSE")
        select new { dr.CustomerUsingCustomerNames, dr.CustomerPrefixName };

      if (query.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckUsingCustomerNames()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Using Customer Names values must be TRUE or FALSE. The destinations below are incorrect.");
        foreach (var tok in query)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok.CustomerPrefixName + " : " + tok.CustomerUsingCustomerNames);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }      
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckUsingCustomerNames()-- finished");
    }
    public static void CheckMinimumIntervals()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumIntervals()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumIntervals()");
      List<string> errList = new List<string>();      

      var query =
        (//from DataRecord dr in StaticVariable.PreRegExDataRecord
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select dr.CustomerMinimumIntervals).Distinct();

      foreach (var tok in query)
      {
        int nParsedValue = 0;
        int nValue = 0;
        if (int.TryParse(tok, out nParsedValue))
        {
          nValue = nParsedValue;
        }
        else
        {
          errList.Add(tok);
        }
        if (nValue < 0)
        {
          errList.Add(tok);
        }
      }
      errList = errList.Distinct().ToList();
      if (errList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckMinimumIntervals()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One of the interval lengths is not an integer. ");
        foreach (string token in errList)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " --> " + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumIntervals()-- finished");
    }
    public static void CheckMinimumDigits()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumDigits()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumDigits()");
      List<string> errList = new List<string>();
      
      var query =
        (//from DataRecord dr in StaticVariable.PreRegExDataRecord
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new { dr.CustomerMinDigits, dr.CustomerPrefixName, dr.ChargingType });

      foreach (var tok in query)
      {       
        if (!tok.ChargingType.ToUpper().Equals("PULSE"))
        {
          int nParsedValue = 0;
          int nValue = 0;
          if (int.TryParse(tok.CustomerMinDigits, out nParsedValue))
          {
            nValue = nParsedValue;
          }
          else
          {
            errList.Add(tok.CustomerPrefixName + " --> " + tok.CustomerMinDigits);
          }
          if (nValue < 0)
          {
            errList.Add(tok.CustomerPrefixName + " --> " + tok.CustomerMinDigits);
          }
        }
        errList = errList.Distinct().ToList();
      }        
      if (errList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckMinimumDigits()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One of the minimum digits is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumDigits()-- finished");
    }
    public static void CheckCutOffDuration()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckCutOffDuration()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckCutOffDuration()");
      List<string> errList = new List<string>();
      List<string> cutOffList = new List<string>();
      
      var query =
        (//from DataRecord dr in StaticVariable.PreRegExDataRecord
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new { dr.CustomerCutOff1Cost, dr.CustomerCutOff2Duration }).Distinct();
      foreach (var tok in query)
      {
        cutOffList.Add(tok.CustomerCutOff1Cost);
        cutOffList.Add(tok.CustomerCutOff2Duration);
      }
      cutOffList = cutOffList.Distinct().ToList();
      foreach (string token in cutOffList)
      {
        int nParsedValue = 0;
        int nValue = 0;
        if (int.TryParse(token, out nParsedValue))
        {
          nValue = nParsedValue;
        }
        else
        {
          errList.Add(token);
        }
        if (nValue < 0)
        {
          errList.Add(token);
        }
      }
      errList = errList.Distinct().ToList();
      if (errList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckCutOffDuration()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One of the Cut-Off values is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "incorrect value  --> " + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckCutOffDuration()-- finished");
    }
    public static void CheckMultiLevelEnabled()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMultiLevelEnabled()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMultiLevelEnabled()");

      var query =
        (//from DataRecord dr in StaticVariable.PreRegExDataRecord
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new { dr.CustomerMultiLevelEnabled }).Distinct();
      foreach (var tok in query)
      {
        if (!(tok.CustomerMultiLevelEnabled.ToUpper().Equals("TRUE")) && !(tok.CustomerMultiLevelEnabled.ToUpper().Equals("FALSE")))
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckMultiLevelEnabled()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Multi Level Enabled values must be TRUE or FALSE. ");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " --> " + tok.CustomerMultiLevelEnabled);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMultiLevelEnabled()-- finished");
    }
    public static void CheckAllSchemes()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckAllSchemes()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckAllSchemes()");

      var query =
        (//from DataRecord dr in StaticVariable.PreRegExDataRecord
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new { dr.CustomerAllSchemes }).Distinct();
      foreach (var tok in query)
      {
        if (!(tok.CustomerAllSchemes.ToUpper().Equals("TRUE")) && !(tok.CustomerAllSchemes.ToUpper().Equals("FALSE")))
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckAllSchemes()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "All Schemes values must be TRUE or FALSE. ");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " --> " + tok.CustomerAllSchemes);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckAllSchemes()-- finished");
    }
    public static void CheckDialTime()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDialTime()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDialTime()");
      List<string> errList = new List<string>();

      var query =
        (//from DataRecord dr in StaticVariable.PreRegExDataRecord
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select dr.CustomerDialTime).Distinct();
      foreach (var tok in query)
      {
        int nParsedValue = 0;
        int nValue = 0;
        if (int.TryParse(tok, out nParsedValue))
        {
          nValue = nParsedValue;
        }
        else
        {
          errList.Add(tok);
        }
        if (nValue < 0)
        {
          errList.Add(tok);
        }
      }
      errList = errList.Distinct().ToList();
      if (errList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckDialTime()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One of the Dial Time values is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " --> " + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDialTime()-- finished");
    }
    public static void CheckMinimumTime()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumTime()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumTime()");
      List<string> errList = new List<string>();      

      var query =
        (//from DataRecord dr in StaticVariable.PreRegExDataRecord
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select dr.CustomerMinimumTime).Distinct();
      foreach (var tok in query)
      {
        int nParsedValue = 0;
        int nValue = 0;
        if (int.TryParse(tok, out nParsedValue))
        {
          nValue = nParsedValue;
        }
        else
        {
          errList.Add(tok);
        }
        if (nValue < 0)
        {
          errList.Add(tok);
        }
      }
      errList = errList.Distinct().ToList();
      if (errList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckMinimumTime()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "One of the minimum digits is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " --> " + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumTime()-- finished");
    }
    public static void CheckIntervalsAtInitialCostGreaterOrEqualToZero()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalsAtInitialCostGreaterOrEqualToZero()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIntervalsAtInitialCostGreaterOrEqualToZero()");
      List<string> errList = new List<string>();      

      var query =
        (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select dr.CustomerIntervalsAtInitialCost).Distinct();

      foreach (var tok in query)
      {
        int nParsedValue = 0;
        int nValue = 0;
        if (int.TryParse(tok, out nParsedValue))
        {
          nValue = nParsedValue;
        }
        else
        {
          errList.Add(tok);
        }
        if (nValue <= 0)
        {
          errList.Add(tok);
        }
      }
            
      if (errList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckIntervalsAtInitialCostGreaterOrEqualToZero()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Intervals At Initial Cost is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          var errQuery =
            from db in StaticVariable.CustomerDetailsDataRecord
            where db.CustomerIntervalsAtInitialCost.Equals(token)
            select new {db.CustomerPrefixName, db.CustomerIntervalsAtInitialCost};

          foreach (var error in errQuery)
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + error.CustomerPrefixName + " : " + error.CustomerIntervalsAtInitialCost);
          }          
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalsAtInitialCostGreaterOrEqualToZero()-- finished");
    }
    public static void CheckTableNames()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTableNames()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTableNames()");     
      var queryTableName =
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        where !dr.CustomerTableName.Contains(StaticVariable.CountryCode) || !dr.CustomerTableName.Contains("_")
        select new { dr.CustomerPrefixName, dr.CustomerTableName };

      if (queryTableName.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckTableNames()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "This table entry is incorrect : ");
        foreach (var table in queryTableName)
        {          
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + table.CustomerPrefixName + " : " + table.CustomerTableName);       
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      var queryTableNameUnique =
        (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord        
        select dr.CustomerTableName).Distinct();

      StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckTableNames()");
      StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Table Names used : ");
      foreach (var uniqueTable in queryTableNameUnique)
      {
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + uniqueTable);
      }           
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTableNames()-- finished");
    }
    public static void CheckDestinationTypesNames()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDestinationTypesNames()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDestinationTypesNames()");
      var query =
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        where !dr.CustomerDestinationType.ToUpper().Equals("LOCAL") && !dr.CustomerDestinationType.ToUpper().Equals("NATIONAL") &&
        !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL") && !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL MOBILE") &&
        !dr.CustomerDestinationType.ToUpper().Equals("MOBILE") && !dr.CustomerDestinationType.ToUpper().Equals("SERVICES") &&
        !dr.CustomerDestinationType.ToUpper().Equals("OTHER")
        select new { dr.CustomerDestinationType, dr.CustomerPrefixName};

      if (query.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckDestinationTypesNames()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Destination types that don't match the V5 default ones of 'Local', 'National', International', International Mobile', 'Mobile', 'Services' & 'Other'");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "If the tariff is for V6 (TDI) then these default destination types are irrelevant");
        foreach (var q in query)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + q.CustomerPrefixName + " : " + q.CustomerDestinationType);
        }
      }                                      
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDestinationTypesNames()-- finished");
    }
    private static void CheckForDuplicateBands()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForDuplicateBands() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForDuplicateBands()");
      List<string> tmpList = new List<string>();
      try
      {
        var query =
          (from DataRecord db in StaticVariable.CustomerDetailsDataRecord
           group db by db.StdBand into newgroup
           where newgroup.Count() > 1
           orderby newgroup.Key
           select newgroup).Distinct();

        foreach (var group in query) //the keys
        {
          foreach (var g in group) //the values.
          {
            tmpList.Add(g.StdBand + " --> " + g.CustomerPrefixName);
          }
        }
        tmpList = tmpList.Distinct().ToList();
        if (tmpList.Any())
        {          
          StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckForDuplicateBands()");
          StaticVariable.Errors.Add("Duplicate Bands:");
          foreach (string tok in tmpList)
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ValidateData::CheckForDuplicateBands()");
        StaticVariable.Errors.Add("Exception Message :: " + e.Message);
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForDuplicateBands() -- finished");
    }
    public static void CheckForMoreThanTwoRegExFiles()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForMoreThanTwoRegExFiles() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForMoreThanTwoRegExFiles()");      
      
      int count = CountNumberOfRegExFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);      
      if (!count.Equals(1))
      {
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There can only be ONE RegEx file in the " + StaticVariable.DatasetsFolder);
        string[] regexes = Directory.GetFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);
        foreach (var regex in regexes)
        {
          if (regex.ToUpper().Contains("REGEX"))
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "- " + Path.GetFileName(regex));
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }

      count = CountNumberOfRegExFiles(StaticVariable.DatasetFolderToUse, Constants.TxtExtensionSearch);      
      if (!count.Equals(1))
      {
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There can only be ONE RegEx file in the " + StaticVariable.DatasetFolderToUse);
        string[] files = Directory.GetFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);
        foreach (var regex in files)
        {
          if (regex.ToUpper().Contains("REGEX"))
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "- " + Path.GetFileName(regex));
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForMoreThanTwoRegExFiles() -- finished");
    }
    private static int CountNumberOfRegExFiles(string folder, string findTextFiles)
    {     
      int fileFound = 0;
      string[] files = Directory.GetFiles(folder, findTextFiles);
      foreach (var file in files)
      {
        if (file.ToUpper().Contains("REGEX"))
        {
          fileFound++;          
        }
      }      
      return fileFound;      
    }   
    private static void CheckForNonMatchingNames()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonMatchingNames()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForNonMatchingNames()");      
      List<string> tmpList = new List<string>();

      var queryNames =
        from db in StaticVariable.CustomerDetailsDataRecord
        where !db.StdPrefixName.ToUpper().Equals(db.CustomerPrefixName.ToUpper())
        select new { db.StdPrefixName, db.CustomerPrefixName };

      if (queryNames.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckForNonMatchingNames()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Std RegEx Names that don't match the Client Names exactly");
        foreach (var names in queryNames)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + names.StdPrefixName.PadRight(44, '.') + " : " + names.CustomerPrefixName);
        }
      }           
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonMatchingNames()-- finished");
    }
    private static void CheckForNonUniqueGroupBands() // to remodel
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonUniqueGroupBands()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForNonUniqueGroupBands()");      
      List<string> tmpList = new List<string>();

      var result =
        (from db in StaticVariable.CustomerDetailsDataRecord
         where (!db.CustomerGroupBand.ToUpper().Equals("NULL", StringComparison.CurrentCultureIgnoreCase) || !db.CustomerGroupBandDescription.ToUpper().Equals("NULL", StringComparison.CurrentCultureIgnoreCase))
           && (db.CustomerUsingGroupBands.ToUpper().Equals("TRUE"))       
        select new
        {
          db.CustomerGroupBand,
          db.CustomerGroupBandDescription,
          db.CustomerFirstInitialRate,
          db.CustomerAllSchemes,
          db.CustomerConnectionCost,
          db.CustomerCutOff1Cost,
          db.CustomerCutOff2Duration,
          db.CustomerDestinationType,
          db.CustomerDialTime,
          db.CustomerFirstSubseqRate,
          db.CustomerFourthInitialRate,
          db.CustomerFourthSubseqRate,
          db.CustomerInitialIntervalLength,
          db.CustomerIntervalsAtInitialCost,
          db.CustomerMinCharge,
          db.CustomerMinDigits,
          db.CustomerMinimumTime,
          db.CustomerMinimumIntervals,
          db.CustomerMultiLevelEnabled,
          db.CustomerRounding,
          db.CustomerSecondInitialRate,
          db.CustomerSecondSubseqRate,
          db.CustomerSubsequentIntervalLength,
          db.CustomerTableName,
          db.CustomerThirdInitialRate,
          db.CustomerThirdSubseqRate,
          db.CustomerTimeScheme,
          db.CustomerUsingGroupBands,
          db.ChargingType
        }).Distinct();
      
      var dupeBands =
        from sr in result
        group sr by sr.CustomerGroupBand into newGroup
        orderby newGroup.Key
        where newGroup.Count() > 1
        select newGroup;

      foreach (var key in dupeBands)
      {
        foreach (var val in key)
        {
          tmpList.Add(val.CustomerGroupBand + " --> " + val.CustomerGroupBandDescription);
        }
      }
      tmpList = tmpList.Distinct().ToList();

      if (tmpList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckForNonUniqueGroupBands()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There are duplicate group bands.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The bands entry may be the same but other fields may be different.");
        foreach (string dupe in tmpList)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + dupe);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonUniqueGroupBands()-- finished");
    }
    public static void CheckSourceDestinationBandsPresentInPrefixBands()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSourceDestinationBandsPresentInPrefixBands()-- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckSourceDestinationBandsPresentInPrefixBands()");
      List<string> uniqueSourceDestinations = new List<string>();
      List<string> errorList = new List<string>();      
      const int sourceDestinationBand = 3;

      //var result = StaticVariable.CustomerDetailsDataRecord.Where(a => a.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? a.CustomerPrefixName : a.StdPrefixName);

      var result2 = StaticVariable.CustomerDetailsDataRecord.Where(b => b.ChargingType.ToUpper().Equals("PULSE"));

      
      

      foreach (string sdb in StaticVariable.SourceDestinationBands)
      {
        string[] sourceDestinationAry = sdb.Split('\t');
        uniqueSourceDestinations.Add(sourceDestinationAry[sourceDestinationBand]);
      }
      uniqueSourceDestinations = uniqueSourceDestinations.Distinct().ToList();
      var bandsInUse = new List<string>();
      bool found = false;

      foreach (DataRecord d in StaticVariable.CustomerDetailsDataRecord)
      {        
        bandsInUse.Add(d.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? d.CustomerGroupBand : d.StdBand);       
      }
      bandsInUse = bandsInUse.Distinct().ToList();      

      foreach (var tok in uniqueSourceDestinations)
      {
        foreach (var band in bandsInUse)
        {
          if (tok.ToLower().Equals(band.ToLower()))
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          errorList.Add(tok);
        }
      }
      if (errorList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData:CheckSourceDestinationBandsPresentInPrefixBands()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There is a Source Destination band in header file that is not found in spreadsheet.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "You may be missing an spreadsheet entry for 'Local', 'National' or 'Regional'?");
        foreach (string band in errorList)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + band);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSourceDestinationBandsPresentInPrefixBands()-- finished");
    }
    public static string AdjustRoundingValueForV6Twb(string value)
    {
      //Console.WriteLine("ValidateData".PadRight(30, '.') + "AdjustRoundingValueForV6Twb() -- started");
      string result = string.Empty;

      if (value.ToLower().Equals("yes") || value.ToLower().Equals("y") || value.ToLower().Equals("1") || value.ToLower().Equals("roundup") || value.ToLower().Equals("round up"))
      {
        result = "YES";
      }
      else if (value.ToLower().Equals("no") || value.ToLower().Equals("n") || value.ToLower().Equals("3") || value.ToLower().Equals("no round") || value.ToLower().Equals("noround") || value.ToLower().Equals("exact"))
      {
        result = "NO";
      }
      else
      {
        result = "NULL";
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::AdjustRoundingValueForV6Twb()");
        StaticVariable.Errors.Add(Environment.NewLine + "The rounding value is incorrect - " + value);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      //Console.WriteLine("ValidateData".PadRight(30, '.') + "AdjustRoundingValueForV6Twb() -- finish");
      return result;
    }        
    private static void CheckIfAllMatrixBandsUsed()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfAllMatrixBandsUsed() -- started");      
      HashSet<string> sourceDestinationBands = new HashSet<string>();
      List<string> allBands = new List<string>();
      const int bandColumn = 3;

      foreach (var band in StaticVariable.SourceDestinationBands)
      {
        var arr = band.Split('\t');
        sourceDestinationBands.Add(arr[bandColumn]);
      }
      var bandQuery =
        (from bnd in StaticVariable.CustomerDetailsDataRecord
         select new {bnd.StdBand, bnd.CustomerGroupBand}).Distinct();

      foreach (var band in bandQuery)
      {
        allBands.Add(band.StdBand);
        allBands.Add(band.CustomerGroupBand);
      }
      allBands = allBands.Distinct().ToList();
      allBands.Sort();

      foreach (var hSet in sourceDestinationBands)
      {
        var found = allBands.Any(band => hSet.ToUpper().Equals(band.ToUpper()));       
        if (!found)
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ValidateData:CheckIfAllMatrixBandsUsed()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The matrix band - " + hSet + " was not found");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfAllMatrixBandsUsed() -- finish");
    }
    private static void CheckChargingType()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckChargingType() -- started");      
      string[] chargingTypes = { Constants.Duration, Constants.Capped, Constants.Pulse };
      bool found = false;
      var queryChargingTypes =
      (from ct in StaticVariable.CustomerDetailsDataRecord
       select ct.ChargingType).Distinct();

      foreach (var type in queryChargingTypes)
      {
        foreach (var ct in chargingTypes)
        {
          if (ct.Equals(type.ToUpper()))
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ValidateData::CheckChargingType()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Charging Type is incorrect.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "It must be either Duration, Capped or Pulse, not: " + type);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        found = false;
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckChargingType() -- finish");
    }
    public static void CheckTariffPlanList()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTariffPlanList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckTariffPlanList()");
      string name = string.Empty;
      string value = string.Empty;
      const int fiveCharsLong = 5;

      if (StaticVariable.TariffPlan.Count.Equals(0))
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckTariffPlanList()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "TariffPlanList is empty");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      else
      {
        StaticVariable.TwbHeader.Add(Environment.NewLine + "Tariff Plan:");
        foreach (string tok in StaticVariable.TariffPlan)
        {
          try
          {
            string[] lines = tok.Split('=');
            name = lines[0];
            value = lines[1];
          }
          catch (Exception e)
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Name \t Value must be tab seperated");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          switch (name.ToUpper())
          {
            case Constants.TariffPlanName:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Tariff Plan Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.TariffPlanName = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.TariffPlanName);
              }
              break;
            case Constants.OperatorName:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Operator Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
              }
              break;
            case Constants.ReleaseDate:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Release Date Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.ReleaseDate = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.ReleaseDate);
              }
              break;
            case Constants.EffectiveFrom:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Effective From Date Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
              }
              break;
            case Constants.Country:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Country Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.Country = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Country);
              }
              break;
            case Constants.CountryCode:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Country Code Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.CountryCode = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.CountryCode);
              }
              break;
            case Constants.CurrencyIsoCode:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Currency (ISOCode) Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
              }
              break;
            case Constants.StartingPointTableName:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Starting Point Table Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
              }
              break;
            case Constants.IsPrivate:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Is Private Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
              }
              break;
            case Constants.Rate1:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Rate 1 Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.Rate1Name = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate1Name);
              }
              break;
            case Constants.Rate2:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Rate 2 Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.Rate2Name = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate2Name);
              }
              break;
            case Constants.Rate3:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Rate 3 Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.Rate3Name = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate3Name);
              }
              break;
            case Constants.Rate4:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Rate 4 Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.Rate4Name = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate4Name);
              }
              break;
            case Constants.TariffReferenceNumber:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: T-ref Number Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  --> " + value + " ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                if (!value.Length.Equals(fiveCharsLong) /*|| !value.StartsWith("T")*/)
                {
                  StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                  StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Tariff reference Number Value Column may have an invalid value.");
                  StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue. It must be 5 chars long.");
                  StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  --> " + value + " ?");
                  ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                }
                else
                {
                  StaticVariable.TariffReferenceNumber = value;
                  StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.TariffReferenceNumber);
                }
              }
              break;
            case Constants.Using:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Using Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
              }
              break;
            case Constants.Version:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Version Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.Version = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Version);
              }
              break;
            case Constants.ExportNds:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Export NDS Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.ExportNds = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.ExportNds);
              }
              break;
            case Constants.CarrierUnitPrice:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList:Carrier Unit Price Value Column has no entry. It must have a value.");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                StaticVariable.CarrierUnitPrice = value;
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.CarrierUnitPrice);
              }
              break;

            case Constants.Holiday:
              if (string.IsNullOrEmpty(value))
              {
                StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTariffPlanList: Holiday Value Column has no entry. It must have at least one value. They are comma seperated");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + "  ?");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              else
              {
                GetHolidaysIntoList(value);
                StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
              }
              break;
            default:
              StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTariffPlanList()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Default Error: A Column has no entry. It must have a value.");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + name + " ? ");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              break;
          }
        }
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTariffPlanList() -- Finished");
    }
    private static void GetHolidaysIntoList(string value)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetHolidaysIntoList() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetHolidaysIntoList()");
      string[] differentHolidays = value.Split(',');
      foreach (string tok in differentHolidays)
      {
        string[] differentDates = tok.Split('-');
        differentDates[0] = differentDates[0].Trim();
        differentDates[1] = differentDates[1].Trim();
        differentDates[2] = differentDates[2].Trim();
        if (differentDates[0].Length.Equals(2) && differentDates[1].Length.Equals(3) && differentDates[2].Length.Equals(4))
        {
          StaticVariable.HolidayList.Add(differentDates[0] + "-" + differentDates[1] + "-" + differentDates[2]);
        }
        else
        {
          StaticVariable.Errors.Add("ProcessRequiredFiles::GetHolidaysIntoList()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The holidays are not in the correct format. They must be like so: DD-Mmm-YYYY.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Check for additional white space. dates must be comma seperated");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetHolidaysIntoList() -- Finished");
    }
    public static void CheckTableLinksList()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTableLinksList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckTableLinksList()");
      const int headerPlusAtLeastOneEntry = 2;
      const int numberOfTableLinksColumns = 4;
      const int tableName = 0;
      const int prefix = 1;
      const int passPrefix = 2;
      const int destination = 3;

      if (StaticVariable.TableLinks.Count < headerPlusAtLeastOneEntry)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTableLinksList()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Table Links List is empty");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      else
      {
        StaticVariable.TwbHeader.Add(Environment.NewLine + "Table Links:");
        StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + "Table Name".PadRight(23, ' ') + "Prefix".PadRight(9, ' ') + "Pass Prefix".PadRight(19, ' ') + "Destination");
        foreach (string tok in StaticVariable.TableLinks)
        {
          string[] aryLine = tok.Split('\t');
          if (!aryLine.Length.Equals(numberOfTableLinksColumns))
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTableLinksList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Table Links has an incorrect entry. There should be 4 columns");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          for (int i = 0; i < numberOfTableLinksColumns; i++)
          {
            if (string.IsNullOrEmpty(aryLine[i]))
            {
              StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTableLinksList()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Table Links has an incorrect entry. One of the columns is empty");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + aryLine[i]);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
          StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + aryLine[tableName].PadRight(23, ' ') + aryLine[prefix].PadRight(9, ' ') + aryLine[passPrefix].PadRight(19, ' ') + aryLine[destination]);
        }
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTableLinksList() -- Finished");
    }
    public static void CheckTimeSchemesList()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTimeSchemesList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckTimeSchemesList()");
      const int numberOfColumns = 3;
      const int schemeName = 0;
      const int holidaysRelevant = 1;
      const int defaultRate = 2;
      if (StaticVariable.TimeSchemes.Count.Equals(0))
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTimeSchemesList()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTimeSchemesList is empty");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      else
      {
        StaticVariable.NumberOfTimeSchemes = StaticVariable.TimeSchemes.Count;
        StaticVariable.TwbHeader.Add(Environment.NewLine + "Time Schemes:");
        StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + "Time Scheme Name\tHolidays Relevant\tDefault Rate");
        foreach (string tok in StaticVariable.TimeSchemes)
        {
          string[] aryLine = tok.Split('\t');
          if (!aryLine.Length.Equals(numberOfColumns))
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTimeSchemesList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Time Schemes has an incorrect entry. There should be 3 columns");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          for (int i = 0; i < numberOfColumns; i++)
          {
            if (string.IsNullOrEmpty(aryLine[i]))
            {
              StaticVariable.Errors.Add("ProcessRequiredFiles::CheckTimeSchemesList()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Time Schemes has an incorrect entry. One of the columns is empty");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + aryLine[i]);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
          StaticVariable.TimeSchemesNames.Add(aryLine[schemeName]);
          StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + aryLine[schemeName].PadRight(19, ' ') + aryLine[holidaysRelevant].PadRight(24, ' ') + aryLine[defaultRate]);
        }
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTimeSchemesList() -- Finished");
    }
    public static void CheckTimeSchemeExceptionsList()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTimeSchemeExceptionsList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckTimeSchemeExceptionsList()");
      StaticVariable.TwbHeader.Add(Environment.NewLine + "Time Schemes Exceptions:");
      StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + "Time Scheme Name".PadRight(19, ' ') + "Day".PadRight(5, ' ') + "Start".PadRight(11, ' ') + "Finish".PadRight(11, ' ') + "Rate");
      List<string> timeSchemeExceptionNames = new List<string>();
      const int schemeName = 0;
      const int day = 1;
      const int start = 2;
      const int finish = 3;
      const int rate = 4;
      foreach (string name in StaticVariable.TimeSchemesExceptions)
      {
        string[] lines = name.Split('\t');
        timeSchemeExceptionNames.Add(lines[schemeName]);
        StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + lines[schemeName].PadRight(19, ' ') + lines[day].PadRight(5, ' ') + lines[start].PadRight(11, ' ') + lines[finish].PadRight(11, ' ') + lines[rate]);
      }
      timeSchemeExceptionNames = timeSchemeExceptionNames.Distinct().ToList();
      foreach (string name in StaticVariable.TimeSchemesNames)
      {
        bool bFound = false;
        foreach (string otherName in timeSchemeExceptionNames)
        {
          if (name.ToUpper().Equals(otherName.ToUpper()))
          {
            bFound = true;
            break;
          }
        }
        if (!bFound)
        {
          StaticVariable.TwbHeader.Add(Environment.NewLine + "The time scheme " + name + " was not defined in Time Scheme Exceptions. It does not need to be if only one rate (24/7) exists.");
        }
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckTimeSchemeExceptionsList() -- Finished");
    }
    public static void CheckSpellingList()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckSpellingList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckSpellingList()");
      string sValue = string.Empty;
      string sName = string.Empty;
      if (StaticVariable.TariffPlan.Count.Equals(0))
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSpellingList()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckSpellingList is empty");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      else
      {
        StaticVariable.TwbHeader.Add(Environment.NewLine + "Spelling:");
        foreach (string tok in StaticVariable.Spelling)
        {
          string[] aryToken = tok.Split('=');
          sValue = aryToken[1];
          sName = aryToken[0];
          if (string.IsNullOrEmpty(sValue))
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSpellingList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "TITLE=SPELLING in Header files has a missing value for ");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + sName);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (sName.ToUpper().Equals(Constants.InternationalMobileSpelling))
          {
            StaticVariable.IntMobileSpelling = sValue;
            StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + "International Mobile Spelling".PadRight(29, ' ') + " = " + StaticVariable.IntMobileSpelling);
          }
          else if (sName.ToUpper().Equals(Constants.InternationalTableSpelling))
          {
            StaticVariable.InternationalTableSpelling = sValue;
            StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + "International Table Spelling".PadRight(29, ' ') + " = " + StaticVariable.InternationalTableSpelling);
          }
          else if (sName.ToUpper().Equals(Constants.NationalTableSpelling))
          {
            StaticVariable.NationalTableSpelling = sValue;
            StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + "National Table Spelling".PadRight(29, ' ') + " = " + StaticVariable.NationalTableSpelling);
          }
          else
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSpellingList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "CheckSpellingList has an extra entry");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + sName + " = " + sValue);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckSpellingList() -- Finished");
    }
    public static void CheckSourceDestinationsBandList()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckSourceDestinationsBandList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckSourceDestinationsBandList()");
      const int numberOfColumns = 4;
      const int band = 3;
      const int table = 0;
      if (!StaticVariable.SourceDestinationBands.Count.Equals(0))
      {
        StaticVariable.TwbHeader.Add(Environment.NewLine + "ProcessRequiredFiles::CheckSourceDestinationsBandList()");
        StaticVariable.TwbHeader.Add(Constants.FiveSpacesPadding + "SourceDestinationsBands: A Matrix is being used." + Environment.NewLine);
        foreach (string tok in StaticVariable.SourceDestinationBands)
        {
          string[] matrixTokens = tok.Split('\t');
          if (!matrixTokens.Length.Equals(numberOfColumns))
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSourceDestinationsBandList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Source Destination (Matrix) columns are incorrect.");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There must be " + numberOfColumns + " columns");
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (matrixTokens[band].Length > Constants.TwbBandLengthLimit)
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSourceDestinationsBandList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Source Destination (Matrix) band length is greater than " + Constants.TwbBandLengthLimit + ".");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "It must be no greater than " + Constants.TwbBandLengthLimit);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          if (StaticVariable.ExportNds.ToUpper().Equals("TRUE") && matrixTokens[band].Length > Constants.V5Tc2BandLengthLimit)
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSourceDestinationsBandList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The NDS band length is over " + Constants.V5Tc2BandLengthLimit + " characters limit.");
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          foreach (string token in matrixTokens)
          {
            if (string.IsNullOrEmpty(token))
            {
              StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSourceDestinationsBandList()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "SourceDestinationsBands has an inncorrect number of columns. There should be " + numberOfColumns + " columns.");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Table Name\tSource\tDestination\tBand");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
          if (!matrixTokens[table].ToUpper().Equals(StaticVariable.NationalTableSpelling.ToUpper()))
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::CheckSourceDestinationsBandList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "SourceDestinationsBands: The table name is not the national name. ");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckSourceDestinationsBandList() -- Finished");
    }
    public static void CheckForStdIntAndBandsFile()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForStdIntAndBandsFile()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckForStdIntAndBandsFile()");
      bool found = false;
      string[] ary = Directory.GetFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);
      foreach (string tok in ary)
      {
        try
        {
          if (Path.GetFileName(tok).ToUpper().Equals(Constants.StdIntAndBands.ToUpper()))
          {
            found = true;
            break;
          }
        }
        catch (Exception e)
        {
          StaticVariable.Errors.Add("ProcessRequiredFiles::CheckForStdIntAndBandsFile");
          StaticVariable.Errors.Add(e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      if (!found)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckForStdIntAndBandsFile");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There is no Std International Bands file in the dataset folder");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "It needs to be called - Std_Int_Names_Bands.txt");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForStdIntAndBandsFile()-- finished");
    }
  }  
}
