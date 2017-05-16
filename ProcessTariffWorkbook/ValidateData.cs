//---------
// <copyright file="ValidateData.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
//---------

using System;

using System.Globalization;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ProcessTariffWorkbook
{
  public static class ValidateData
  {
    public static void PreRegExMatchValidateCustomerDetailsRecord()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "PreRegExMatchValidateCustomerDetailsDataRecord() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "PreRegExMatchValidateCustomerDetailsDataRecord() -- started");
      StaticVariable.MissingCountryExceptions = CreateCountryExceptionsHashset();         
      CheckPricesAreInCorrectFormat();
      CheckTableNames();
      CheckForDestinationTypes();
      CheckRoundingForIncorrectEntry();
      CheckTimeSchemeForIncorrectEntry();
      CheckIfMinCostAndRate4SubseqAreSame();
      CheckGrouping();
      CheckIntervalLengthsGreaterOrEqualToZero();
      CheckUsingCustomerNames();
      CheckMinimumIntervals();
      CheckMinimumDigits();
      CheckMaximumPrices();
      CheckIfInitialIntervalSameAsSubsequentInterval();
      CheckCutOffDuration();
      CheckMultiLevelEnabled();
      CheckAllSchemes();
      CheckDialTime();
      CheckMinimumTime();
      CheckIntervalsAtInitialCostGreaterOrEqualToZero();
      CheckDestinationTypesNames();         
      Console.WriteLine("ValidateData".PadRight(30, '.') + "PreRegExMatchValidateCustomerDetailsDataRecord() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "PreRegExMatchValidateCustomerDetailsDataRecord() -- finished");
    }
    public static void PostRegExMatchValidateCustomerDetailsRecord()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "PostRegExMatchValidateCustomerDetailsDataRecord() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "PostRegExMatchValidateCustomerDetailsDataRecord() -- started");         
      CheckTablesForDefaultValue();
      CheckForMissingDefaultEntries();
      CheckForFreephone();
      CheckIfFreephoneIsFree();
      CheckForDuplicateBands();
      CheckForNonUniqueGroupBands();
      CheckSourceDestinationBandsPresentInPrefixBands();
      CheckGroupBands();
      CheckIfMinCostAnd4ThRateSamePrice();
      CheckIfAllMatrixBandsUsed();
      CheckChargingType();
      CheckDestinationsAssignedMultipleBands();         
      Console.WriteLine("ValidateData".PadRight(30, '.') + "PostRegExMatchValidateCustomerDetailsDataRecord() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "PostRegExMatchValidateCustomerDetailsDataRecord() -- finished");
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
    public static void CheckForCommasInLine(string sComma)
    {
      if (sComma.Contains(","))
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckForCommasInPrices()");
        StaticVariable.ProgressDetails.Add("ParseInputFile:ReadXLSXFileIntoList()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There is a comma in the file. No commas are allowed.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The offending line is: " + sComma);
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
        StaticVariable.ProgressDetails.Add("ValidateData::SetToFourDecimalPlaces()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Failed to set price to 4 decimal places:");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      return result;
    }
    public static string AdjustRoundingValue(string value)
    {
      var result = string.Empty;
      if (value.ToLower().Equals("yes") || value.ToLower().Equals("y") || value.ToLower().Equals("1") ||
            value.ToLower().Equals("roundup") || value.ToLower().Equals("round up"))
      {
        result = "Round Up";
      }
      else if (value.ToLower().Equals("no") || value.ToLower().Equals("n") || value.ToLower().Equals("3") ||
              value.ToLower().Equals("exact") || value.ToLower().Equals("noround") ||
              value.ToLower().Equals("no round"))
      {
      result = "Exact";
      }
      else
      {           
      StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::AdjustRoundingValue()");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The rounding value is incorrect - " + value);
      ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      return result;
    }
    private static void CheckPricesAreInCorrectFormat()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckPricesAreInCorrectFormat() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckPricesAreInCorrectFormat() -- started");
      List<string> resultsList = new List<string>();

      var query =
            from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
            select new
            {
              dr.CustomerPrefixName,
              dr.CustomerFirstInitialRate,
              dr.CustomerFirstSubseqRate,
              dr.CustomerSecondInitialRate,
              dr.CustomerSecondSubseqRate,
              dr.CustomerThirdInitialRate,
              dr.CustomerThirdSubseqRate,
              dr.CustomerFourthInitialRate,
              dr.CustomerFourthSubseqRate,
              dr.CustomerMinCharge,
              dr.CustomerConnectionCost,
              dr.ChargingType
            };

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
                  StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Duration: One of the prices is not a double. ");
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok.CustomerPrefixName + " --> " + price);
                  ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
            }
            catch (Exception e)
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Duration:");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
        }
        else if (tok.ChargingType.ToUpper().Equals(Constants.Capped))
        {
            double parsedDoubleValue = 0.0;
            int parsedIntValue = 0;
            if (!double.TryParse(tok.CustomerFirstInitialRate, out parsedDoubleValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: Capped Price per minute is not in correct format. it must be a double: " + tok.CustomerFirstInitialRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: If there are no capped prices, delete the capped worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (!double.TryParse(tok.CustomerFirstSubseqRate, out parsedDoubleValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: Capped Price is not in correct format. it must be a double: " + tok.CustomerFirstSubseqRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: If there are no capped prices, delete the capped worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (!int.TryParse(tok.CustomerSecondInitialRate, out parsedIntValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: Capped Time is not in correct format. it must be a int. time in minutes: " + tok.CustomerSecondInitialRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: If there are no capped prices, delete the capped worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (!double.TryParse(tok.CustomerMinCharge, out parsedDoubleValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: Capped Minimum Cost is not in correct format. it must be a double. " + tok.CustomerMinCharge);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: If there are no capped prices, delete the capped worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (!double.TryParse(tok.CustomerConnectionCost, out parsedDoubleValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine +
                                                  "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: Capped Connection Cost is not in correct format. it must be a double. " + tok.CustomerConnectionCost);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Capped: If there are no capped prices, delete the capped worksheet");
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
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: Pulse Length given in decimal format must be multipled by 100. it must be changed to an int: " + tok.CustomerFirstInitialRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: If there are no Pulse prices, delete the Pulse worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (parsedIntValue / 100 < minimumPulseLength)
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine +
                                                  "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: Pulse Length is too short: " + tok.CustomerFirstInitialRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "It may not have been multipled by 100. It must be changed to an int from a decimal: ");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: If there are no Pulse prices, delete the Pulse worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (!int.TryParse(tok.CustomerFirstSubseqRate, out parsedIntValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: Pulse Unit is not in correct format. it must be a int, normally 1: " + tok.CustomerFirstSubseqRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: If there are no Pulse prices, delete the Pulse worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (!double.TryParse(tok.CustomerMinCharge, out parsedDoubleValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine +
                                                  "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: Minimum Cost is not in correct format. it must be a double: " + tok.CustomerMinCharge);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: If there are no Pulse prices, delete the Pulse worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (!double.TryParse(tok.CustomerConnectionCost, out parsedDoubleValue))
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckPricesAreInCorrectFormat()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: Connection Cost is not in correct format. it must be a double: " + tok.CustomerConnectionCost);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse: If there are no Pulse prices, delete the Pulse worksheet");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckPricesAreInCorrectFormat() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckPricesAreInCorrectFormat() -- finished");
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
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForDestinationTypes() -- started");
      if (StaticVariable.ExportNdsValue.ToUpper().Equals("TRUE"))
      {
        var query =
          from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
          where
          !dr.CustomerDestinationType.ToUpper().Equals("LOCAL") &&
          !dr.CustomerDestinationType.ToUpper().Equals("NATIONAL") &&
          !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL") &&
          !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL MOBILE") &&
          !dr.CustomerDestinationType.ToUpper().Equals("SERVICES") &&
          !dr.CustomerDestinationType.ToUpper().Equals("OTHER") &&
          !dr.CustomerDestinationType.ToUpper().Equals("MOBILE")
          select new
          {
            dr.CustomerDestinationType,
            dr.CustomerPrefixName
          };

        if (query.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckForDestinationTypes()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "This destination type is invalid for V5 RingMaster. ");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "It must be either 'Local', 'National', International', 'International Mobile', 'Mobile', 'Services' or 'Other'");
          foreach (var item in query)
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The offending destinationan type is - " + item.CustomerDestinationType + " and Customer  name is - " + item.CustomerPrefixName);
          }
          StaticVariable.ProgressDetails.Add(Environment.NewLine + Constants.FiveSpacesPadding + "Comment out 'ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();' in method to supress killing program");
          //ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForDestinationTypes() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForDestinationTypes() -- finished");
    }
    private static void CheckTablesForDefaultValue()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTablesForDefaultValue() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTablesForDefaultValue() -- started");
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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckTablesForDefaultValue()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "A default prefix name (e.g Local,?) does not have a entry (e.g. " + StaticVariable.CountryCodeValue + "_Local) in prefix links header. \nIs the prefix file missing?");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Check if the name spelt correctly or is there an INI file that is not required?" + Environment.NewLine);
        foreach (var item in extraTablesInPrefixes)
        {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + item);
        }
      }
      if (extraTablesInTableLinks.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckTablesForDefaultValue()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "A table in prefix links header does not have a default prefix in its corresponding table. ");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "E.g The Prefix Links has an entry for " + StaticVariable.CountryCodeValue + "_Local but it does not have a default entry (e.g. Local,?) in the " + StaticVariable.CountryCodeValue + "_Local table");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Check if the name spelt correctly or an ini file may be missing for that prefix link.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There may be no error. The default prefix may just be different to the table name." + Environment.NewLine);
        foreach (var item in extraTablesInTableLinks)
        {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + item);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTablesForDefaultValue() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTablesForDefaultValue() -- finished");
    }
    private static void CheckRoundingForIncorrectEntry()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckRoundingForIncorrectEntry() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckRoundingForIncorrectEntry() -- started");
      List<string> errors = new List<string>();
      try
      {
        var query =
            from DataRecord db in StaticVariable.CustomerDetailsDataRecord
            select new
            {
              db.CustomerRounding,
              db.CustomerPrefixName,
              db.ChargingType
            };

        foreach (var q in query)
        {
            if (q.ChargingType.ToUpper().Equals("PULSE"))
              continue;
            string custRounding = q.CustomerRounding.ToUpper();
            if (!(custRounding.Equals("YES") || custRounding.Equals("1") || custRounding.Equals("Y") ||
                  custRounding.Equals("ROUNDUP") || custRounding.Equals("ROUND UP") ||
                  custRounding.Equals("NO") || custRounding.Equals("3") || custRounding.Equals("N") ||
                  custRounding.Equals("EXACT") || custRounding.Equals("NO ROUND") ||
                  custRounding.Equals("NOROUND")))
            {
              errors.Add(q.CustomerPrefixName + " is --> " + custRounding);
            }
        }
        if (errors.Any())
        {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckRoundingForIncorrectEntry()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Rounding Values are incorrect for these destinations.");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The must be 'Yes', 'Y', 'ROUND UP', 'ROUNDUP' or '1' for round up and 'No', 'N', 'EXACT', 'NOROUND', 'NO ROUND' or '3' for no round");
            foreach (string error in errors)
            {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + error);
            }
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      catch (Exception e)
      {
        StaticVariable.ProgressDetails.Add("ValidateData::CheckRounding()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Exception Message :: " + e.Message);
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckRounding() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckRoundingForIncorrectEntry() -- finished");
    }
    private static void CheckTimeSchemeForIncorrectEntry()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTimeSchemeForIncorrectEntry() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTimeSchemeForIncorrectEntry() -- started");
        List<string> errors = new List<string>();
        const int timeScheme = 0;
        bool found = false;

        var queryCustomerTimeSchemes =
              from DataRecord db in StaticVariable.CustomerDetailsDataRecord
              select new
              {
                db.CustomerTimeScheme,
                db.CustomerPrefixName
              };

        foreach (var q in queryCustomerTimeSchemes)
        {
          foreach (var timeSchemeName in StaticVariable.TimeSchemes)
          {
              string[] timeschemes = timeSchemeName.Split('\t');
              if (!q.CustomerTimeScheme.ToUpper().Equals(timeschemes[timeScheme].ToUpper()))
                continue;
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
          StaticVariable.ProgressDetails.Add(Environment.NewLine +
                                            "ValidateData::CheckTimeSchemeForIncorrectEntry()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Time Scheme Values are incorrect for these destinations.");
          foreach (string error in errors)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + error);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckRounding() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTimeSchemeForIncorrectEntry() -- finished");
    }
    private static void CheckIfMinCostAndRate4SubseqAreSame()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "MinCostAndRate4SubseqAreSame() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "MinCostAndRate4SubseqAreSame() -- started");      

      var query =
        from DataRecord drm in StaticVariable.CustomerDetailsDataRecord
        where
        SetToFourDecimalPlaces(drm.CustomerFourthSubseqRate).Equals(SetToFourDecimalPlaces(drm.CustomerMinCharge)) &&
        Convert.ToDouble(drm.CustomerMinCharge) > 0.0
        select new
        {
          drm.CustomerPrefixName,
          drm.CustomerMinCharge,          
        };
      query.ToList().Sort();
     
      if (query.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::MinCostAndRate4SubseqAreSame()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Minimum Cost is the same price as the 4th Rate Subsequent price. This is not normally correct. Recheck.");
        foreach (var name in query)
        {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.CustomerPrefixName + ": Price = " + name.CustomerMinCharge );
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "MinCostAndRate4SubseqAreSame() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "MinCostAndRate4SubseqAreSame() -- finished");
    }
    private static void CheckForFreephone()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForFreephone() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForFreephone() -- started");
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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckForFreephone()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There is no entry for Freephone.");
        Console.WriteLine("There is no entry for Freephone.............");
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForFreephone() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForFreephone() -- started");
    }
    private static void CheckIfFreephoneIsFree()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfFreephoneIsFree() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIfFreephoneIsFree() -- started");
      bool bFound = false;
      List<string> NonZeroFreeCalls = new List<string>();

      var results =
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        where (db.CustomerPrefixName.ToUpper().Contains("FREE") || db.StdBand.ToUpper().Contains("FREE") ||
              db.CustomerGroupBand.ToUpper().Contains("FREE") || db.CustomerGroupBand.ToUpper().Contains("TOLL") ||
              db.CustomerPrefixName.ToUpper().Contains("GRAT") || db.StdBand.ToUpper().Contains("GRAT") ||
              db.CustomerPrefixName.ToUpper().Contains("TOLL") || db.StdBand.ToUpper().Contains("TOLL") ||
              db.StdPrefixName.ToUpper().Contains("FREE") || db.StdBand.ToUpper().Contains("FREE"))
        select new
        {
          db.CustomerPrefixName,
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
        if (CheckIfPriceZero(result.CustomerFirstInitialRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerFirstSubseqRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerSecondInitialRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerSecondSubseqRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerThirdInitialRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerThirdSubseqRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerFourthInitialRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerFourthSubseqRate))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerMinCharge))
        {
            bFound = true;
        }
        else if (CheckIfPriceZero(result.CustomerConnectionCost))
        {
            bFound = true;
        }

        if (bFound)
        {
            NonZeroFreeCalls.Add("Customer Name: " + result.CustomerPrefixName);
            bFound = false;
        }
      }
      if (NonZeroFreeCalls.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckIfFreephoneIsFree()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The destinations below may or may not be free calls.");
        foreach (var item in NonZeroFreeCalls)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + item);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfFreephoneIsFree() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIfFreephoneIsFree() -- finished");
    }
    private static bool CheckIfPriceZero(string sValue)
    {
      bool notZero = false;
      foreach (char c in sValue)
      {
        if (c.Equals('0') || c.Equals('.'))
            continue;
        notZero = true;
        break;
      }
      return notZero;
    }
    private static void CheckGrouping()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckGrouping() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckGrouping() -- started");

      #region using group bands

      List<string> usingGroupBands = new List<string>();
      var queryUsingGroupBands =
            from DataRecord db in StaticVariable.CustomerDetailsDataRecord
            where
            !db.CustomerUsingGroupBands.ToUpper().Equals("TRUE") &&
            !db.CustomerUsingGroupBands.ToUpper().Equals("FALSE")
            select new
            {
              db.CustomerUsingGroupBands,
              db.CustomerPrefixName
            };

      foreach (var q in queryUsingGroupBands)
      {
        usingGroupBands.Add(q.CustomerPrefixName + " --> " + q.CustomerUsingGroupBands);
      }
      if (usingGroupBands.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckGrouping()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Using Group Bands field are incorrect for these destinations.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The must be 'TRUE' or 'FALSE'");

        foreach (string s in usingGroupBands)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + s);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      #endregion

      #region group band

      List<string> groupBands = new List<string>();
      var queryGroupBand =
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        select new
        {
          db.CustomerGroupBand,
          db.CustomerPrefixName,
          db.CustomerUsingGroupBands
        };

      foreach (var q in queryGroupBand)
      {
        if (q.CustomerGroupBand.Length > Constants.V5Tc2BandLengthLimit && StaticVariable.ExportNdsValue.ToUpper().Equals("TRUE") && q.CustomerUsingGroupBands.Equals("TRUE"))
        {
          groupBands.Add(q.CustomerPrefixName + " --> " + q.CustomerGroupBand);
        }
      }
      if (groupBands.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckGrouping()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Group Band field are too long for these destinations.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The must be no greater than 4 chars long.");

        foreach (string band in groupBands)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + band);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      # endregion

      # region group band description
      List<string> groupBandDescriptions = new List<string>();
      var queryGroupBandDescription =
        from DataRecord db in StaticVariable.CustomerDetailsDataRecord
        select new
        {
          db.CustomerGroupBandDescription,
          db.CustomerPrefixName,
          db.CustomerUsingGroupBands
        };

      foreach (var q in queryGroupBandDescription)
      {
        if (q.CustomerGroupBandDescription.Length > Constants.V5Tc2BandDescriptionLength &&
            StaticVariable.ExportNdsValue.Equals("TRUE") && q.CustomerUsingGroupBands.Equals("TRUE"))
        {
            groupBandDescriptions.Add(q.CustomerPrefixName + " --> " + q.CustomerGroupBandDescription);
        }
      }
      if (groupBandDescriptions.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckGrouping()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Group Band Description field are too long for these destinations.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The must be no greater than 20 chars long.");

        foreach (string s in groupBandDescriptions)
        {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + s);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      # endregion

      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckGrouping() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckGrouping() -- finished");
    }
    public static void CheckIntervalLengthsGreaterOrEqualToZero()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalLengthsGreaterOrEqualToZero() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIntervalLengthsGreaterOrEqualToZero() -- started");
      List<string> results = new List<string>();
      List<string> errors = new List<string>();
      int nValue = 0;
      const string defaultIntervalLength = "60";

      var queryIntervalLengthGreaterThanZero =
      (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
      where !dr.ChargingType.ToUpper().Equals("PULSE")
      select new
      {
        dr.CustomerInitialIntervalLength,
        dr.CustomerSubsequentIntervalLength
      }).Distinct();

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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckIntervalLengthsGreaterOrEqualToZero()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "One or more of the interval lengths is not an integer or is less than 1. ");
        foreach (string error in errors)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The illegal value is - " + error);
          var query =
            from qry in StaticVariable.CustomerDetailsDataRecord
            where
            qry.CustomerInitialIntervalLength.Equals(error) ||
            qry.CustomerSubsequentIntervalLength.Equals(error)
            select
            new
            {
              qry.CustomerPrefixName,
              qry.CustomerInitialIntervalLength,
              qry.CustomerSubsequentIntervalLength
            };

          foreach (var q in query)
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + q.CustomerPrefixName + " - ");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "initial interval Length    = " + q.CustomerInitialIntervalLength);
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "subsequent interval length = " + q.CustomerSubsequentIntervalLength);
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      var queryIntervalLengthValues =
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        where
        !dr.ChargingType.ToUpper().Equals("PULSE") &&
        (!dr.CustomerInitialIntervalLength.Equals(defaultIntervalLength) ||
        !dr.CustomerSubsequentIntervalLength.Equals(defaultIntervalLength))
        select
        new
        {
          dr.CustomerInitialIntervalLength,
          dr.CustomerSubsequentIntervalLength,
          dr.CustomerPrefixName
        };

      if (queryIntervalLengthValues.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine +
                                          "ValidateData::CheckIntervalLengthsGreaterOrEqualToZero()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The interval lengths listed below are not the default 60 seconds. ");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "This may be correct. ");
        foreach (var interval in queryIntervalLengthValues)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + interval.CustomerPrefixName + " - ");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "initial interval Length    = " + interval.CustomerInitialIntervalLength);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.FiveSpacesPadding + "subsequent interval length = " + interval.CustomerSubsequentIntervalLength);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalLengthsGreaterOrEqualToZero() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIntervalLengthsGreaterOrEqualToZero() -- finished");
    }
    public static void CheckUsingCustomerNames()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckUsingCustomerNames() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckUsingCustomerNames() -- started");

        var query =
              from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
              where
              !dr.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") &&
              !dr.CustomerUsingCustomerNames.ToUpper().Equals("FALSE")
              select new
              {
                dr.CustomerUsingCustomerNames,
                dr.CustomerPrefixName
              };

        if (query.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckUsingCustomerNames()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                            "Using Customer Names values must be TRUE or FALSE. The destinations below are incorrect.");
          foreach (var tok in query)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok.CustomerPrefixName + " : " + tok.CustomerUsingCustomerNames);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckUsingCustomerNames() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckUsingCustomerNames() -- finished");
    }
    public static void CheckMinimumIntervals()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumIntervals() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumIntervals() -- started");
        List<string> errList = new List<string>();

        var query =
          (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
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
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckMinimumIntervals()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "One of the interval lengths is not an integer. ");
          foreach (string token in errList)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + " --> " + token);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumIntervals() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumIntervals() -- finished");
    }
    public static void CheckMinimumDigits()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumDigits() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumDigits() -- started");
      List<string> errList = new List<string>();

      var query =
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new
        {
          dr.CustomerMinDigits,
          dr.CustomerPrefixName,
          dr.ChargingType
        };

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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckMinimumDigits()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "One of the minimum digits is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumDigits() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumDigits() -- finished");
    }
    public static void CheckCutOffDuration()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckCutOffDuration() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckCutOffDuration() -- started");
      List<string> errList = new List<string>();
      List<string> cutOffList = new List<string>();

      var query =
        (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
          select new
          {
            dr.CustomerCutOff1Cost,
            dr.CustomerCutOff2Duration
          }).Distinct();
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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckCutOffDuration()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "One of the Cut-Off values is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "incorrect value  --> " + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckCutOffDuration() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckCutOffDuration() -- finished");
    }
    public static void CheckMultiLevelEnabled()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMultiLevelEnabled() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMultiLevelEnabled() -- started");

      var query =
        (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new
        {
          dr.CustomerMultiLevelEnabled
        }).Distinct();
      foreach (var tok in query)
      {
        if (!(tok.CustomerMultiLevelEnabled.ToUpper().Equals("TRUE")) &&
            !(tok.CustomerMultiLevelEnabled.ToUpper().Equals("FALSE")))
        {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckMultiLevelEnabled()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                "Multi Level Enabled values must be TRUE or FALSE. ");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + " --> " + tok.CustomerMultiLevelEnabled);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMultiLevelEnabled() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMultiLevelEnabled() -- finished");
    }
    public static void CheckAllSchemes()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckAllSchemes() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckAllSchemes() -- started");

      var query =
        (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        select new
        {
          dr.CustomerAllSchemes
        }).Distinct();
      foreach (var tok in query)
      {
        if (!(tok.CustomerAllSchemes.ToUpper().Equals("TRUE")) && !(tok.CustomerAllSchemes.ToUpper().Equals("FALSE")))
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckAllSchemes()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "All Schemes values must be TRUE or FALSE. ");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + " --> " + tok.CustomerAllSchemes);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckAllSchemes() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckAllSchemes() -- finished");
    }
    public static void CheckDialTime()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDialTime() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDialTime() -- started");
        List<string> errList = new List<string>();

        var query =
          (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
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
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckDialTime()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "One of the Dial Time values is not an integer or is less than zero. ");
          foreach (string token in errList)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + " --> " + token);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDialTime() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDialTime() -- finished");
    }
    public static void CheckMinimumTime()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumTime() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumTime() -- started");
      List<string> errList = new List<string>();

      var query =
        (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckMinimumTime()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "One of the minimum digits is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + " --> " + token);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMinimumTime() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMinimumTime() -- finished");
    }
    public static void CheckIntervalsAtInitialCostGreaterOrEqualToZero()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalsAtInitialCostGreaterOrEqualToZero() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIntervalsAtInitialCostGreaterOrEqualToZero() -- started");
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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckIntervalsAtInitialCostGreaterOrEqualToZero()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Intervals At Initial Cost is not an integer or is less than zero. ");
        foreach (string token in errList)
        {
          var errQuery =
          from db in StaticVariable.CustomerDetailsDataRecord
          where db.CustomerIntervalsAtInitialCost.Equals(token)
          select new
          {
            db.CustomerPrefixName,
            db.CustomerIntervalsAtInitialCost
          };

          foreach (var error in errQuery)
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + error.CustomerPrefixName + " : " + error.CustomerIntervalsAtInitialCost);
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIntervalsAtInitialCostGreaterOrEqualToZero() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIntervalsAtInitialCostGreaterOrEqualToZero() -- finished");
    }
    public static void CheckTableNames()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTableNames() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTableNames() -- started");
      var queryTableName =
            from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
            where !dr.CustomerTableName.Contains(StaticVariable.CountryCodeValue) || !dr.CustomerTableName.Contains("_")
            orderby dr.CustomerTableName
            select new
            {
              dr.CustomerPrefixName,
              dr.CustomerTableName
            };

      if (queryTableName.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckTableNames()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "This table entry is incorrect : ");
        foreach (var table in queryTableName)
        {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + table.CustomerPrefixName + " : " + table.CustomerTableName);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      var queryTableNameUnique =
      (from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
      select dr.CustomerTableName).Distinct();

      StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckTableNames()");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Table Names used : ");
      foreach (var uniqueTable in queryTableNameUnique)
      {
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + uniqueTable);
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTableNames() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTableNames() -- finished");
    }
    public static void CheckDestinationTypesNames()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDestinationTypesNames() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDestinationTypesNames() -- started");
      var query =
        from DataRecord dr in StaticVariable.CustomerDetailsDataRecord
        where
        !dr.CustomerDestinationType.ToUpper().Equals("LOCAL") &&
        !dr.CustomerDestinationType.ToUpper().Equals("NATIONAL") &&
        !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL") &&
        !dr.CustomerDestinationType.ToUpper().Equals("INTERNATIONAL MOBILE") &&
        !dr.CustomerDestinationType.ToUpper().Equals("MOBILE") &&
        !dr.CustomerDestinationType.ToUpper().Equals("SERVICES") &&
        !dr.CustomerDestinationType.ToUpper().Equals("OTHER")
        select new
        {
          dr.CustomerDestinationType,
          dr.CustomerPrefixName
        };

      if (query.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckDestinationTypesNames()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Destination types that don't match the V5 default ones of 'Local', 'National', International', International Mobile', 'Mobile', 'Services' & 'Other'");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If the tariff is for V6 (TDI) then these default destination types are irrelevant");
        foreach (var q in query)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + q.CustomerPrefixName + " : " + q.CustomerDestinationType);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDestinationTypesNames() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDestinationTypesNames() -- finished");
    }
    private static void CheckForDuplicateBands()
  {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForDuplicateBands() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForDuplicateBands() -- started");
      List<string> tmpList = new List<string>();
      try
      {
        var query =
        (from DataRecord db in StaticVariable.CustomerDetailsDataRecord
          group db by db.StdBand into newgroup
          where newgroup.Count() > 1
          orderby newgroup.Key
          select newgroup).Distinct();

        foreach (var group in query)
        {
          foreach (var g in group)
          {
            tmpList.Add(g.StdBand + " --> " + g.CustomerPrefixName);
          }
        }
        tmpList = tmpList.Distinct().ToList();
        if (tmpList.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckForDuplicateBands()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error 1: There may be two Regexes with the same band.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error 2: There may be two entries for the client name in the input Xlsx file. E.g. 'Martinique' & 'French Antilles Martinique'");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error 3: Check the Header file for entries also in xlsx file.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error 4: If no duplicate band is shown then prefix name is being matched by two regexes." + Environment.NewLine);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Duplicate Bands:");
          foreach (string tok in tmpList)
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      catch (Exception e)
      {
        StaticVariable.ProgressDetails.Add("ValidateData::CheckForDuplicateBands()");
        StaticVariable.ProgressDetails.Add("Exception Message :: " + e.Message);
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForDuplicateBands() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForDuplicateBands() -- finished");
    }
    public static void CheckForMoreThanTwoRegExFiles()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForMoreThanTwoRegExFiles() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForMoreThanTwoRegExFiles() -- started");

      int count = CountNumberOfRegExFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);
      if (!count.Equals(1))
      {
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There can only be ONE RegEx file in the " + StaticVariable.DatasetsFolder);
        string[] regexes = Directory.GetFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);
        foreach (var regex in regexes)
        {
          if (regex.ToUpper().Contains("REGEX"))
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "- " + Path.GetFileName(regex));
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }

      count = CountNumberOfRegExFiles(StaticVariable.DatasetFolderToUse, Constants.TxtExtensionSearch);
      if (!count.Equals(1))
      {
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There can only be ONE RegEx file in the " + StaticVariable.DatasetFolderToUse);
        string[] files = Directory.GetFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);
        foreach (var regex in files)
        {
          if (regex.ToUpper().Contains("REGEX"))
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "- " + Path.GetFileName(regex));
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForMoreThanTwoRegExFiles() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForMoreThanTwoRegExFiles() -- finished");
    }
    private static int CountNumberOfRegExFiles(string folder, string findTextFiles)
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CountNumberOfRegExFiles() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CountNumberOfRegExFiles() -- started");
      int fileFound = 0;
      string[] files = Directory.GetFiles(folder, findTextFiles);
      foreach (var file in files)
      {
        if (file.ToUpper().Contains("REGEX"))
        {
            fileFound++;
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CountNumberOfRegExFiles() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CountNumberOfRegExFiles() -- finished");
      return fileFound;
    }
    private static void CheckForNonMatchingCustomerNames()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonMatchingCustomerNames() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForNonMatchingCustomerNames() -- started");
      List<string> tmpList = new List<string>();

      var queryNames =
        from db in StaticVariable.CustomerDetailsDataRecord
        orderby db.StdPrefixName
        where !db.StdPrefixName.ToUpper().Equals(db.CustomerPrefixName.ToUpper())
        select new
        {
          db.StdPrefixName,
          db.CustomerPrefixName
        };

      if (queryNames.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckForNonMatchingCustomerNames()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Std RegEx Names that don't match the Client Names exactly");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Standard Name".PadRight(44, '.') + " : XLSX Name exactly");  
        foreach (var names in queryNames)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + names.StdPrefixName.PadRight(44, '.') + " : " + names.CustomerPrefixName);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonMatchingCustomerNames() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForNonMatchingCustomerNames() -- finished");
    }
    private static void CheckForNonUniqueGroupBands()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonUniqueGroupBands() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') +
                                      "CheckForNonUniqueGroupBands() -- started");
      List<string> tmpList = new List<string>();

      var result =
      (from db in StaticVariable.CustomerDetailsDataRecord
      where
      (!db.CustomerGroupBand.ToUpper().Equals("NULL", StringComparison.CurrentCultureIgnoreCase) ||
      !db.CustomerGroupBandDescription.ToUpper().Equals("NULL", StringComparison.CurrentCultureIgnoreCase))
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
        group sr by sr.CustomerGroupBand
        into newGroup
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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckForNonUniqueGroupBands()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There are duplicate group bands.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                          "The bands entry may be the same but other fields may be different.");
        foreach (string dupe in tmpList)
        {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + dupe);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForNonUniqueGroupBands() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForNonUniqueGroupBands() -- finished");
    }
    public static void CheckSourceDestinationBandsPresentInPrefixBands()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSourceDestinationBandsPresentInPrefixBands() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckSourceDestinationBandsPresentInPrefixBands() -- started");
        List<string> errorList = new List<string>();
        const int sourceDestinationBand = 3;

        List<string> uniqueSourceDestinations = StaticVariable.SourceDestinationBands.Select(sdb => sdb.Split('\t')).Select(sourceDestinationAry => sourceDestinationAry[sourceDestinationBand]).ToList();
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
          if (bandsInUse.Any(band => tok.ToLower().Equals(band.ToLower())))
          {
              found = true;
          }
          if (!found)
          {
              errorList.Add(tok);
          }
        }
        if (errorList.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData:CheckSourceDestinationBandsPresentInPrefixBands()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There is a Source Destination band in header file that is not found in spreadsheet.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "You may be missing an spreadsheet entry for 'Local', 'National' or 'Regional'?");
          foreach (string band in errorList)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + band);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSourceDestinationBandsPresentInPrefixBands() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckSourceDestinationBandsPresentInPrefixBands() -- finished");
    }
    public static string AdjustRoundingValueForV6Twb(string value)
    {
        string result = string.Empty;

        if (value.ToLower().Equals("yes") || value.ToLower().Equals("y") || value.ToLower().Equals("1") ||
              value.ToLower().Equals("roundup") || value.ToLower().Equals("round up"))
        {
          result = "YES";
        }
        else if (value.ToLower().Equals("no") || value.ToLower().Equals("n") || value.ToLower().Equals("3") ||
                value.ToLower().Equals("no round") || value.ToLower().Equals("noround") ||
                value.ToLower().Equals("exact"))
        {
          result = "NO";
        }
        else
        {
          result = "NULL";
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::AdjustRoundingValueForV6Twb()");
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "The rounding value is incorrect - " + value);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        return result;
    }
    private static void CheckIfAllMatrixBandsUsed()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfAllMatrixBandsUsed() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIfAllMatrixBandsUsed() -- started");
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
        select new
        {
          bnd.StdBand,
          bnd.CustomerGroupBand
        }).Distinct();

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
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData:CheckIfAllMatrixBandsUsed()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The matrix band - " + hSet + " was not found");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfAllMatrixBandsUsed() -- finish");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIfAllMatrixBandsUsed() -- finish");
    }
    private static void CheckChargingType()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckChargingType() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckChargingType() -- started");
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
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckChargingType()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Charging Type is incorrect.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "It must be either Duration, Capped or Pulse, not: " + type);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        found = false;
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckChargingType() -- finish");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckChargingType() -- finish");
    }
    public static void CheckTariffPlanList()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTariffPlanList() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTariffPlanList() -- started");
      string name = string.Empty;
      string value = string.Empty;
      const int fiveCharsLong = 5;
      string carrierUnitPriceValue = string.Empty;

      if (StaticVariable.TariffPlan.Count.Equals(0))
      {
        StaticVariable.ProgressDetails.Add("ValidateData".PadRight(30, '.') + "CheckTariffPlanList()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TariffPlanList is empty");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      else
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "Tariff Plan:");
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
              StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                  "CheckTariffPlanList: Name \t Value must be tab seperated");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            switch (name.ToUpper())
            {
              case Constants.TariffPlanName:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Tariff Plan Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.TariffPlanNameValue = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.TariffPlanNameValue);
                  }
                  break;
              case Constants.OperatorName:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Operator Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
                  }
                  break;
              case Constants.ReleaseDate:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Release Date Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ReleaseDateValue = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.ReleaseDateValue);
                  }
                  break;
              case Constants.EffectiveFrom:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Effective From Date Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
                  }
                  break;
              case Constants.Country:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Country Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.CountryValue = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') +
                                                  " = " + StaticVariable.CountryValue);
                  }
                  break;
              case Constants.CountryCode:
                  if (string.IsNullOrEmpty(value) || !CheckIfInteger(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Country Code Value Column has an incorrect entry. It must have a integer value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + " : " + value);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.CountryCodeValue = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') +
                                                  " = " + StaticVariable.CountryCodeValue);
                  }
                  break;
              case Constants.CurrencyIsoCode:
                  if (string.IsNullOrEmpty(value) || !CheckIfInteger(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Currency (ISOCode) Value Column has an incorrect entry. It must be an integer value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + " : " + value);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') +
                                                  " = " + value);
                  }
                  break;
              case Constants.StartingPointTableName:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Starting Point Table Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') +
                                                  " = " + value);
                  }
                  break;
              case Constants.IsPrivate:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Is Private Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') +
                                                  " = " + value);
                  }
                  break;
              case Constants.Rate1:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                        "CheckTariffPlanList: Rate 1 Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.Rate1Name = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate1Name);
                  }
                  break;
              case Constants.Rate2:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Rate 2 Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.Rate2Name = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate2Name);
                  }
                  break;
              case Constants.Rate3:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Rate 3 Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.Rate3Name = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate3Name);
                  }
                  break;
              case Constants.Rate4:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Rate 4 Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.Rate4Name = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.Rate4Name);
                  }
                  break;
              case Constants.TariffReferenceNumber:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: T-ref Number Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  --> " + value + " ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    if (!value.Length.Equals(fiveCharsLong))
                    {
                        StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Tariff reference Number Value Column may have an invalid value.");
                        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue. It must be 5 chars long.");
                        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  --> " + value + " ?");
                        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                    }
                    else
                    {
                        StaticVariable.TariffReferenceNumberValue = value;
                        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.TariffReferenceNumberValue);
                    }
                  }
                  break;
              case Constants.Using:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Using Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
                  }
                  break;
              case Constants.Version:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Version Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.VersionValue = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.VersionValue);
                  }
                  break;
              case Constants.ExportNds:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Export NDS Value Column has no entry. It must have a value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    StaticVariable.ExportNdsValue = value;
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + StaticVariable.ExportNdsValue);
                  }
                  break;
              case Constants.CarrierUnitPrice:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList:Carrier Unit Price Value Column has no entry. It must have a double value.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If the carrier unit is not being used (i.e. no Pulse prices, delete carrier unit price from 'TariffPlan' header and the Pulse worksheet.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  if (!value.ToUpper().Equals("N/A"))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Carrier Unit Price is 'N/A' or blank. It must be a double.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If the carrier unit is not being used (i.e. no Pulse prices, delete carrier unit price from 'TariffPlan' header and the Pulse worksheet.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + value);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  if (!CheckIfDouble(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Carrier Unit Price is not correct. It must be a double.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + " : " + value);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  carrierUnitPriceValue = value;
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + carrierUnitPriceValue + " : If pulse is not being used, delete Carrier Unit Price in TWB header file & delete pulse work sheet.");
                  break;
              case Constants.Holiday:
                  if (string.IsNullOrEmpty(value))
                  {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Holiday Value Column has no entry. It must have at least one value. They are comma seperated");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + "  ?");
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  }
                  else
                  {
                    GetHolidaysIntoList(value);
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name.PadRight(26, ' ') + " = " + value);
                  }
                  break;
              default:
                  StaticVariable.ProgressDetails.Add("ValidateData::CheckTariffPlanList()");
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckTariffPlanList: Default Error: A Column has no entry. It must have a value.");
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Name\t\tValue");
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name + " ? ");
                  ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                  break;
            }
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTariffPlanList() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTariffPlanList() -- finished");

    }
    private static void GetHolidaysIntoList(string value)
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "GetHolidaysIntoList() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "GetHolidaysIntoList() -- started");
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
          StaticVariable.ProgressDetails.Add("ValidateData::GetHolidaysIntoList()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The holidays are not in the correct format. They must be like so: DD-Mmm-YYYY.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Check for additional white space. dates must be comma seperated");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "GetHolidaysIntoList() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "GetHolidaysIntoList() -- finished");
    }
    public static void CheckTableLinksList()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTableLinksList() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTableLinksList() -- started");
        const int headerPlusAtLeastOneEntry = 2;
        const int numberOfTableLinksColumns = 4;
        const int tableName = 0;
        const int prefix = 1;
        const int passPrefix = 2;
        const int destination = 3;

        if (StaticVariable.TableLinks.Count < headerPlusAtLeastOneEntry)
        {
          StaticVariable.ProgressDetails.Add("ValidateData::CheckTableLinksList()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Table Links List is empty");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        else
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "Table Links:");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Table Name".PadRight(23, ' ') +
                                      "Prefix".PadRight(9, ' ') + "Pass Prefix".PadRight(19, ' ') + "Destination");
          foreach (string tok in StaticVariable.TableLinks)
          {
              string[] aryLine = tok.Split('\t');
              if (!aryLine.Length.Equals(numberOfTableLinksColumns))
              {
                StaticVariable.ProgressDetails.Add("ValidateData::CheckTableLinksList()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                    "Table Links has an incorrect entry. There should be 4 columns");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok);
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              for (int i = 0; i < numberOfTableLinksColumns; i++)
              {
                if (string.IsNullOrEmpty(aryLine[i]))
                {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTableLinksList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                    "Table Links has an incorrect entry. One of the columns is empty");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + aryLine[i]);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                }
                if (i.Equals(passPrefix))
                {
                    if (aryLine[passPrefix].ToUpper().Equals("NO") ||
                      aryLine[passPrefix].ToUpper().Equals("YES"))
                      continue;
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTableLinksList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                    "Table Links has an incorrect entry in the Pass Prefix column. It should be either 'TRUE' or 'FALSE'");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + aryLine[passPrefix]);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                }
              }
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + aryLine[tableName].PadRight(23, ' ') +
                                            aryLine[prefix].PadRight(9, ' ') +
                                            aryLine[passPrefix].PadRight(19, ' ') +
                                            aryLine[destination]);
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTableLinksList() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTableLinksList() -- finished");
    }
    public static void CheckTimeSchemesList()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTimeSchemesList() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTableLinksList() -- finished");
        int numberOfTimeSchemes = 0;
        const int numberOfColumns = 3;
        const int schemeName = 0;
        const int holidaysRelevant = 1;
        const int defaultRate = 2;
        if (StaticVariable.TimeSchemes.Count.Equals(0))
        {
          StaticVariable.ProgressDetails.Add("ValidateData::CheckTimeSchemesList()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckTimeSchemesList is empty");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        else
        {            
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "Time Schemes:");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Time Scheme Name\tHolidays Relevant\tDefault Rate");
          foreach (string tok in StaticVariable.TimeSchemes)
          {
              string[] aryLine = tok.Split('\t');
              if (!aryLine.Length.Equals(numberOfColumns))
              {
                StaticVariable.ProgressDetails.Add("ValidateData::CheckTimeSchemesList()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Time Schemes has an incorrect entry. There should be 3 columns");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok);
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              for (int i = 0; i < numberOfColumns; i++)
              {
                if (string.IsNullOrEmpty(aryLine[i]))
                {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckTimeSchemesList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Time Schemes has an incorrect entry. One of the columns is empty");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + aryLine[i]);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                }
              }
              StaticVariable.TimeSchemesNames.Add(aryLine[schemeName]);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + aryLine[schemeName].PadRight(19, ' ') + aryLine[holidaysRelevant].PadRight(24, ' ') + aryLine[defaultRate]);
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTimeSchemesList() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTimeSchemesList() -- finished");
    }
    public static void CheckTimeSchemeExceptionsList()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTimeSchemeExceptionsList() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTimeSchemeExceptionsList() -- started");
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "Time Schemes Exceptions:");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Time Scheme Name".PadRight(19, ' ') + "Day".PadRight(5, ' ') + "Start".PadRight(11, ' ') + "Finish".PadRight(11, ' ') + "Rate");
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
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + lines[schemeName].PadRight(19, ' ') + lines[day].PadRight(5, ' ') + lines[start].PadRight(11, ' ') + lines[finish].PadRight(11, ' ') + lines[rate]);
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
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "The time scheme " + name + " was not defined in Time Scheme Exceptions. It does not need to be if only one rate (24/7) exists.");
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckTimeSchemeExceptionsList() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckTimeSchemeExceptionsList() -- finished");
    }
    public static void CheckSpellingList()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSpellingList() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckSpellingList() -- started");
        string sValue = string.Empty;
        string sName = string.Empty;
        if (StaticVariable.TariffPlan.Count.Equals(0))
        {
          StaticVariable.ProgressDetails.Add("ValidateData::CheckSpellingList()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "CheckSpellingList is empty");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        else
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "Spelling:");
          foreach (string tok in StaticVariable.Spelling)
          {
              string[] aryToken = tok.Split('=');
              sValue = aryToken[1];
              sName = aryToken[0];
              if (string.IsNullOrEmpty(sValue))
              {
                StaticVariable.ProgressDetails.Add("ValidateData::CheckSpellingList()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TITLE=SPELLING in Header files has a missing value for ");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + sName);
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              if (sName.ToUpper().Equals(Constants.InternationalMobileSpelling))
              {
                StaticVariable.InternationalMobileSpellingValue = sValue;
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "International Mobile Spelling".PadRight(29, ' ') + " = " + StaticVariable.InternationalMobileSpellingValue);
              }
              else if (sName.ToUpper().Equals(Constants.InternationalTableSpelling))
              {
                StaticVariable.InternationalTableSpellingValue = sValue;
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "International Table Spelling".PadRight(29, ' ') + " = " + StaticVariable.InternationalTableSpellingValue);
              }
              else if (sName.ToUpper().Equals(Constants.NationalTableSpelling))
              {
                StaticVariable.NationalTableSpellingValue = sValue;
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "National Table Spelling".PadRight(29, ' ') + " = " + StaticVariable.NationalTableSpellingValue);
              }
              else
              {
                StaticVariable.ProgressDetails.Add("ValidateData::CheckSpellingList()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "CheckSpellingList has an extra entry");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + sName + " = " + sValue);
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSpellingList() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckSpellingList() -- finished");
    }
    public static void CheckSourceDestinationsBandList()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSourceDestinationsBandList() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckSourceDestinationsBandList() -- started");
        const int twbBandLengthLimit = 100; // current TWB (V6) limit
        const int numberOfColumns = 4;
        const int band = 3;
        const int table = 0;
        if (StaticVariable.SourceDestinationBands.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckSourceDestinationsBandList()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                      "SourceDestinationsBands: A Matrix is being used.");
          foreach (string tok in StaticVariable.SourceDestinationBands)
          {
              string[] matrixTokens = tok.Split('\t');
              if (!matrixTokens.Length.Equals(numberOfColumns))
              {
                StaticVariable.ProgressDetails.Add("ValidateData::CheckSourceDestinationsBandList()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Source Destination (Matrix) columns are incorrect.");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There must be " + numberOfColumns + " columns");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              if (matrixTokens[band].Length > twbBandLengthLimit)
              {
                StaticVariable.ProgressDetails.Add("ValidateData::CheckSourceDestinationsBandList()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Source Destination (Matrix) band length is greater than " + twbBandLengthLimit + ".");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "It must be no greater than " + twbBandLengthLimit);
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              if (StaticVariable.ExportNdsValue.ToUpper().Equals("TRUE") &&
                matrixTokens[band].Length > Constants.V5Tc2BandLengthLimit)
              {
                StaticVariable.ProgressDetails.Add("ValidateData::CheckSourceDestinationsBandList()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The NDS band length is over " + Constants.V5Tc2BandLengthLimit + " characters limit.");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
              }
              foreach (string token in matrixTokens)
              {
                if (string.IsNullOrEmpty(token))
                {
                    StaticVariable.ProgressDetails.Add("ValidateData::CheckSourceDestinationsBandList()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "SourceDestinationsBands has an inncorrect number of columns. There should be " + numberOfColumns + " columns.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Table Name\tSource\tDestination\tBand");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                }
              }
              if (matrixTokens[table].ToUpper().Equals(StaticVariable.NationalTableSpellingValue.ToUpper())) continue;
              StaticVariable.ProgressDetails.Add("ValidateData::CheckSourceDestinationsBandList()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "SourceDestinationsBands: The table name is not the national name. ");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckSourceDestinationsBandList() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckSourceDestinationsBandList() -- finished");
    }
    public static void CheckForStdIntAndBandsFile()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForStdIntAndBandsFile() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForStdIntAndBandsFile() -- started");
        bool found = false;
        string[] ary = Directory.GetFiles(StaticVariable.DatasetsFolder, Constants.TxtExtensionSearch);
        foreach (string tok in ary)
        {
          try
          {
              if (!Path.GetFileName(tok).ToUpper().Equals(Constants.StdIntAndBands.ToUpper()))
                continue;
              found = true;
              break;
          }
          catch (Exception e)
          {
              StaticVariable.ProgressDetails.Add("ValidateData::CheckForStdIntAndBandsFile");
              StaticVariable.ProgressDetails.Add(e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        if (!found)
        {
          StaticVariable.ProgressDetails.Add("ValidateData::CheckForStdIntAndBandsFile");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There is no Std International Bands file in the dataset folder");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "It needs to be called - Std_Int_Names_Bands.txt");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForStdIntAndBandsFile() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForStdIntAndBandsFile() -- finished");
    }
    private static void CheckForMissingDefaultEntries()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForMissingDefaultEntries() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckForMissingDefaultEntries() -- started");
        List<string> tables = new List<string>();
        bool found = false;

        var queryDefaultPrefixes =
              from db in StaticVariable.PrefixNumbersRecord
              where db.PrefixNumber.Equals("?")
              orderby db.TableName
              select db;

        var numberOFTables =
        (from dd in StaticVariable.CustomerDetailsDataRecord
        select dd.CustomerTableName).Distinct();

        foreach (var tab in numberOFTables)
        {
          if (queryDefaultPrefixes.Any(table => table.TableName.ToUpper().Equals(tab.ToUpper())))
          {
              found = true;
          }
          if (found)
              continue;
          tables.Add(tab);
          found = false;
        }
        if (tables.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckForMissingDefaultEntries");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There amy be a default prefix missing");
          foreach (var item in tables)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + item);
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckForMissingDefaultEntries() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') +
                                        "CheckForMissingDefaultEntries() -- finished");
    }
    public static bool CheckForPulseWorksheet()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "GetWorksheetsToBeUsed() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "GetWorksheetsToBeUsed() -- started");
        List<string> workSheetsNotUsed = new List<string>();
        List<string> WorkSheetsUsed = new List<string>();
        bool found = false;
        string[] worksheetsTypes = { Constants.Duration, Constants.Capped, Constants.Pulse };
        SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(StaticVariable.InputFile);

        foreach (string wksheet in worksheetsTypes)
        {
          try
          {
              SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[wksheet];
              SpreadsheetGear.IRange cells = worksheet.Cells;
              WorkSheetsUsed.Add(wksheet);
          }
          catch (Exception)
          {
              workSheetsNotUsed.Add(wksheet);
          }
        }
        if (WorkSheetsUsed.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::GetWorksheetsToBeUsed");
          foreach (var item in WorkSheetsUsed)
          {
              if (item.ToUpper().Equals(Constants.Pulse.ToUpper()))
              {
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + " A pulse worksheet has been detected. If it is not being used, delete worksheet.");
                found = true;
              }
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "GetWorksheetsToBeUsed() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "GetWorksheetsToBeUsed() -- finished");
        return found;
    }
    public static void CheckforAllTariffPlanEntries(string value)
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckforAllTariffPlanEntries() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckforAllTariffPlanEntries() -- started");
        Dictionary<string, string> checkEntries = new Dictionary<string, string>();
        checkEntries.Add(Constants.TariffPlanName.ToUpper(), "present");
        checkEntries.Add(Constants.OperatorName.ToUpper(), "present");
        checkEntries.Add(Constants.ReleaseDate.ToUpper(), "present");
        checkEntries.Add(Constants.EffectiveFrom.ToUpper(), "present");
        checkEntries.Add(Constants.Country.ToUpper(), "present");
        checkEntries.Add(Constants.CountryCode.ToUpper(), "present");
        checkEntries.Add(Constants.CurrencyIsoCode.ToUpper(), "present");
        checkEntries.Add(Constants.StartingPointTableName.ToUpper(), "present");
        checkEntries.Add(Constants.IsPrivate.ToUpper(), "present");
        checkEntries.Add(Constants.Rate1.ToUpper(), "present");
        checkEntries.Add(Constants.Rate2.ToUpper(), "present");
        checkEntries.Add(Constants.Rate3.ToUpper(), "present");
        checkEntries.Add(Constants.Rate4.ToUpper(), "present");
        checkEntries.Add(Constants.Using.ToUpper(), "present");
        checkEntries.Add(Constants.TariffReferenceNumber.ToUpper(), "present");
        checkEntries.Add(Constants.Version.ToUpper(), "present");
        checkEntries.Add(Constants.ExportNds.ToUpper(), "present");
        //checkEntries.Add(Constants.CarrierUnitPrice.ToUpper(), "present");
        checkEntries.Add(Constants.Holiday.ToUpper(), "present");

        try
        {
          var entries = value.Split('=');
          var entry = checkEntries[entries[0].ToUpper()];
        }
        catch (Exception)
        {
          if (value.ToUpper().Contains(Constants.CarrierUnitPrice.ToUpper()))
          {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckforAllTariffPlanEntries()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Is a pulse rate being used?. If not delete the 'Carrier unit Price' in the header file and the pulse worksheet.");
          }
          else
          {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckforAllTariffPlanEntries()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The entry - " + value + " is missing in the Tariff Plan header.");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckforAllTariffPlanEntries() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckforAllTariffPlanEntries() -- finished");
    }
    public static void CheckDestinationsAssignedMultipleBands()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "DestinationsAssignedMultipleBands() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "DestinationsAssignedMultipleBands() -- started");
        Dictionary<string, string> errorNames = new Dictionary<string, string>();

        var queryCustomerDetails =
              (from dr in StaticVariable.CustomerDetailsDataRecord
              select new
              {
                  dr.StdPrefixName,
                  dr.CustomerPrefixName,
                  dr.StdBand,
                  dr.CustomerUsingCustomerNames
              })
              .Distinct();

        foreach (var detail in queryCustomerDetails)
        {
          try
          {
              errorNames.Add(detail.StdBand.ToUpper(), "Custname - " + detail.CustomerPrefixName + ",\t stdName - " + detail.StdPrefixName + ",\t band - " + detail.StdBand + ",\t UsingCustName - ");
          }
          catch (Exception e)
          {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::DestinationsAssignedMultipleBands()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problems adding values to dictionary.");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message + " - " + detail.CustomerPrefixName);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If 'using customer name' is true, an entry like 'peru NGN PRS' will be assigned two different bands but the customer name will be added twice into the dictionary.");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Fix: when using standard prefixes: In the " + StaticVariable.XlsxFileName + " file split the entry 'peru NGN PRS' into two seperate entries - 'peru NGN' and 'peru PRS'.");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Fix: when using client supplied prefixes: In the regex, remove one of the regex matches.");
              StaticVariable.ProgressDetails.Add(Environment.NewLine + Constants.FiveSpacesPadding + "Existing dictionary entry : ");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + errorNames[detail.CustomerPrefixName.ToUpper()]);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Duplicate entry :");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Custname - " + detail.CustomerPrefixName + ",\t StdName - " + detail.StdPrefixName + ",\t band - " + detail.StdBand + ",\t UsingCustName - ");
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "DestinationsAssignedMultipleBands() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "DestinationsAssignedMultipleBands() -- finished");
    }
    public static void CheckGroupBands()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckGroupBands() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckGroupBands() -- started");
        bool durationUsed = false;

        try
        {
          var durationQuery =
          (from groups in StaticVariable.CustomerDetailsDataRecord
            where groups.CustomerUsingGroupBands.ToUpper().Equals("TRUE") && groups.ChargingType.ToUpper().Equals("DURATION")
            select new
            {
              groups.ChargingType,
              groups.CustomerGroupBand,
              groups.CustomerGroupBandDescription,
              groups.CustomerMinCharge,
              groups.CustomerConnectionCost,
              groups.CustomerFirstInitialRate,
              groups.CustomerFirstSubseqRate,
              groups.CustomerSecondInitialRate,
              groups.CustomerSecondSubseqRate,
              groups.CustomerThirdInitialRate,
              groups.CustomerThirdSubseqRate,
              groups.CustomerFourthInitialRate,
              groups.CustomerFourthSubseqRate
            }).Distinct();

          if (durationQuery.Any())
          {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckGroupBands()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "group bands used.........");
            durationUsed = true;
            foreach (var dg in durationQuery)
            {                               
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Charging Type: " + dg.ChargingType);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Band:".PadRight(15, '.') + '\x0020' + dg.CustomerGroupBand.PadRight(10, '\x0020') + " Band Description:".PadRight(20, '.') + '\x0020' + dg.CustomerGroupBandDescription);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Min Cost:".PadRight(15, '.') + '\x0020' + dg.CustomerMinCharge.PadRight(10, '\x0020') + " Connection Cost:".PadRight(20, '.') + '\x0020' + dg.CustomerConnectionCost);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Rate 1 Initial:".PadRight(15, '.') + '\x0020' + dg.CustomerFirstInitialRate.PadRight(10, '\x0020') + " Rate 1 Subseq:".PadRight(20, '.') + '\x0020' + dg.CustomerFirstSubseqRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Rate 2 initial:".PadRight(15, '.') + '\x0020' + dg.CustomerSecondInitialRate.PadRight(10, '\x0020') + " Rate 2 Subseq:".PadRight(20, '.') + '\x0020' + dg.CustomerSecondSubseqRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Rate 3 Initial:".PadRight(15, '.') + '\x0020' + dg.CustomerThirdInitialRate.PadRight(10, '\x0020') + " Rate 3 Subseq:".PadRight(20, '.') + '\x0020' + dg.CustomerThirdSubseqRate);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Rate 4 Initial:".PadRight(15, '.') + '\x0020' + dg.CustomerFourthInitialRate.PadRight(10, '\x0020') + " Rate 4 Subseq:".PadRight(20, '.') + '\x0020' + dg.CustomerFourthSubseqRate + Environment.NewLine);
            }
          }            
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add("ValidateData::CheckGroupBands()");
          StaticVariable.ProgressDetails.Add("Error in CheckGroupBands() - duration");
          StaticVariable.ProgressDetails.Add(e.Message);
        }
        try
        {
          var cappedQuery =
          (from groups in StaticVariable.CustomerDetailsDataRecord
            where groups.CustomerUsingGroupBands.ToUpper().Equals("TRUE") && groups.ChargingType.ToUpper().Equals("CAPPED")
            select groups).Distinct();
          foreach (var cq in cappedQuery)
          {
            if (!durationUsed)                 
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckGroupBands()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "group bands used.........");
            }
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Charging Type: " + cq.ChargingType);
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Band:".PadRight(15, '.') + '\x0020' + cq.CustomerGroupBand.PadRight(10, '\x0020') + " Band Description:".PadRight(20, '.') + '\x0020' + cq.CustomerGroupBandDescription);
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Min Cost:".PadRight(15, '.') + '\x0020' + cq.CustomerMinCharge.PadRight(10, '\x0020') + " Connection Cost:".PadRight(20, '.') + '\x0020' + cq.CustomerConnectionCost);
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Standard:".PadRight(15, '.') + '\x0020' + cq.CustomerFirstInitialRate.PadRight(10, '\x0020') + " Cap Price:".PadRight(20, '.') + '\x0020' + cq.CustomerFirstSubseqRate);
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Cap Time:".PadRight(15, '.') + '\x0020' + cq.CustomerSecondInitialRate.PadRight(10, '\x0020'));
            StaticVariable.ProgressDetails.Add("\n");
          }
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add("ValidateData::CheckGroupBands()");
          StaticVariable.ProgressDetails.Add("Error in CheckGroupBands() - capped");
          StaticVariable.ProgressDetails.Add(e.Message);
        }
        try
        {
          var pulseQuery =
          (from groups in StaticVariable.CustomerDetailsDataRecord
            where groups.CustomerUsingGroupBands.ToUpper().Equals("TRUE") && groups.ChargingType.ToUpper().Equals("PULSE")
            select groups).Distinct();
          foreach (var pq in pulseQuery)
          {
              if (!durationUsed)              
              {
                StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckGroupBands()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "group bands used.........");
              }
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Charging Type: " + pq.ChargingType);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Band:".PadRight(15, '.') + '\x0020' + pq.CustomerGroupBand.PadRight(10, '\x0020') + " Band Description:".PadRight(20, '.') + '\x0020' + pq.CustomerGroupBandDescription);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Min Cost:".PadRight(15, '.') + '\x0020' + pq.CustomerMinCharge.PadRight(10, '\x0020') + " Connection Cost:".PadRight(20, '.') + '\x0020' + pq.CustomerConnectionCost);
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Pulse Length:".PadRight(15, '.') + '\x0020' + pq.CustomerFirstInitialRate.PadRight(10, '\x0020') + " Pulse Unit:".PadRight(20, '.') + '\x0020' + pq.CustomerFirstSubseqRate);
              StaticVariable.ProgressDetails.Add("\n");
          }
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add("ValidateData::CheckGroupBands()");
          StaticVariable.ProgressDetails.Add("Error in CheckGroupBands() - pulse");
          StaticVariable.ProgressDetails.Add(e.Message);
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckGroupBands() -- finish");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckGroupBands() -- finish");
    }
    public static void CheckIfMinCostAnd4ThRateSamePrice()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "DisplayMinCostV4thRateSamePrice() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "DisplayMinCostV4thRateSamePrice() -- started");
        List<string> matches = new List<string>();
        var query =
          from peakAndMinCost in StaticVariable.CustomerDetailsDataRecord
          where peakAndMinCost.CustomerFourthSubseqRate.Equals(peakAndMinCost.CustomerMinCharge)
          select new
          {
              peakAndMinCost.StdPrefixName,
              peakAndMinCost.CustomerFourthSubseqRate,
              peakAndMinCost.CustomerMinCharge,
              peakAndMinCost.CustomerPrefixName
          };

        foreach (var pk in query)
        {
          if (pk.CustomerMinCharge.Equals(pk.CustomerFourthSubseqRate) && Convert.ToDouble(pk.CustomerMinCharge) > 0.0)
          {
              ValidateData.CheckIfDouble(pk.CustomerMinCharge);
              matches.Add(Constants.FiveSpacesPadding + pk.CustomerPrefixName);
          }
        }
        if (matches.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData:DisplayMinCostV4thRateSamePrice()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "4th rate is the same price as Minimum Cost.");
          foreach (var m in matches)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + m);
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "DisplayMinCostV4thRateSamePrice() -- finish");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "DisplayMinCostV4thRateSamePrice() -- finish");
    }
    public static void FindMissingInternationalCountries()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "FindMissingInternationalCountries() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "FindMissingInternationalCountries() -- started");
        List<string> tmpList = new List<string>();
        List<string> noPricesForCountriesList = new List<string>();
        List<string> bandsNotFoundList = new List<string>();
        string custband = string.Empty;
        const string zeroPrices = "0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000";
        string otherInfo = "FALSE\tnull\tnull\t" + StaticVariable.CountryCodeValue + "_International\tInternational\tRounding\tTimeScheme\tFALSE\t60\t60\t0\t1\t1\t1\tTRUE\tFALSE\t0\t0\t0\tCharging Type";

        CreateStdIntBandsDataRecord();
        var intQuery =
          from drm in StaticVariable.CustomerDetailsDataRecord
          where drm.CustomerTableName.ToUpper().Equals(StaticVariable.InternationalTableSpellingValue.ToUpper())
          select new
          {
              drm.StdPrefixName,
              drm.StdBand,
              drm.CustomerPrefixName
          };

        var customerCountryList = intQuery.Select(q => q.StdBand + "\t" + q.StdPrefixName + "\t" + q.CustomerPrefixName).ToList();

        foreach (StandardInternationalBandsDataRecord sib in StaticVariable.StandardInternationalBands)
        {
          bool bFound = false;
          foreach (string tok in customerCountryList)
          {
              string[] ary = tok.Split('\t');
              custband = ary[0];
              if (sib.SBand.ToUpper().Equals(custband.ToUpper()))
              {
                bFound = true;
                break;
              }
          }
          if (!bFound)
          {
              bandsNotFoundList.Add(sib.SBand + "\t" + sib.SPrefixName + "\t" + sib.SCountryCode);
          }
        }
        foreach (string token in bandsNotFoundList)
        {
          string[] aryBandsNotFound = token.Split('\t');
          custband = aryBandsNotFound[0];
          string custName = aryBandsNotFound[1];
          string stdCountryCode = aryBandsNotFound[2];
          const string mobileBand = "M";

          if (custband.EndsWith(mobileBand) && custband.Length > 2)
          {
              string tmpCustBand = custband.Substring(0, custband.Length - 1);
              var mobileQuery =
              from drm in StaticVariable.CustomerDetailsDataRecord
              where tmpCustBand.ToUpper().Equals(drm.StdBand.ToUpper())
              select drm;
              foreach (var dr in mobileQuery)
              {
                string name = dr.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? dr.CustomerPrefixName : dr.StdPrefixName;
                string groupBand = dr.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? dr.CustomerGroupBand + mobileBand : dr.StdBand + mobileBand;
                string groupBandDescription = dr.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? dr.CustomerGroupBandDescription + mobileBand : dr.StdPrefixDescription + mobileBand;
                tmpList.Add(name + " " + StaticVariable.InternationalMobileSpellingValue + "\t" + dr.CustomerFirstInitialRate + "\t" +
                    dr.CustomerFirstSubseqRate + "\t" + dr.CustomerSecondInitialRate + "\t" + dr.CustomerSecondSubseqRate + "\t" +
                    dr.CustomerThirdInitialRate + "\t" + dr.CustomerThirdSubseqRate + "\t" + dr.CustomerFourthInitialRate + "\t" +
                    dr.CustomerFourthSubseqRate + "\t" + dr.CustomerMinCharge + "\t" + dr.CustomerConnectionCost + "\t" +
                    "FALSE\t" + groupBand + "\t" + groupBandDescription + "\t" + ValidateData.CapitaliseWord(dr.CustomerTableName) + "\t" +
                    ValidateData.CapitaliseWord(dr.CustomerDestinationType + " " + StaticVariable.InternationalMobileSpellingValue) + "\t" +
                    dr.CustomerRounding + "\t" + ValidateData.CapitaliseWord(dr.CustomerTimeScheme) + "\t" +
                    dr.CustomerUsingCustomerNames + "\t" + dr.CustomerInitialIntervalLength + "\t" + dr.CustomerSubsequentIntervalLength + "\t" +
                    dr.CustomerMinimumIntervals + "\t" + dr.CustomerIntervalsAtInitialCost + "\t" + dr.CustomerMinimumTime + "\t" + dr.CustomerDialTime + "\t" +
                    dr.CustomerAllSchemes + "\t" + dr.CustomerMultiLevelEnabled + "\t" + dr.CustomerMinDigits + "\t" +
                    dr.CustomerCutOff1Cost + "\t" + dr.CustomerCutOff2Duration + "\t" + dr.ChargingType
                    );
              }
          }
          else if (!stdCountryCode.Equals(StaticVariable.CountryCodeValue) && !stdCountryCode.ToUpper().Equals("N/A"))
          {
              noPricesForCountriesList.Add(ValidateData.CapitaliseWord(custName) + "\t" + zeroPrices + "\t" + otherInfo);
          }
        }
        if (tmpList.Any())
        {
          tmpList.Sort();
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData:FindMissingInternationalCountries()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Int mobile prices missing for these destinations, so given landine prices.");
          foreach (string sEntry in tmpList)
          {
              StaticVariable.ProgressDetails.Add(sEntry);
          }
        }

        if (noPricesForCountriesList.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData:FindMissingInternationalCountries()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "No prices supplied for these countries. They have been given a zero price.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If a BT / Eir / KPN price is available for these destinations then use that price unless told to zero cost these destinations.");
          noPricesForCountriesList.Sort();
          foreach (string sZeroEntry in noPricesForCountriesList)
          {
              StaticVariable.ProgressDetails.Add(sZeroEntry);
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "FindMissingInternationalCountries() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "FindMissingInternationalCountries() -- finished");
    }
    public static void CheckDestinationsAssignedIncorrectTable()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDestinationsAssignedIncorrectTable() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDestinationsAssignedIncorrectTable() -- started");

        var query =
          (from drm in StaticVariable.CustomerDetailsDataRecord
          join pn in StaticVariable.PrefixNumbersRecord on drm.StdPrefixName.ToUpper() equals pn.stdPrefixName.ToUpper()
          where !drm.CustomerTableName.ToUpper().Equals(pn.TableName.ToUpper())
          select new
          {
              pn.TableName,
              pn.PrefixName,
              drm.CustomerTableName,
              drm.CustomerPrefixName
          }).Distinct();

        if (query.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckDestinationsAssignedIncorrectTable()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Prefix Name and prefix are assigned different tables.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Perhaps check the RegExMatchedList for incorrect regex matching." + Environment.NewLine);
          foreach (var entry in query)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + ("Prefix Table: " + entry.PrefixName + " --> " + entry.TableName).PadRight(70) + "XLSX File: " + entry.CustomerPrefixName + " --> " + entry.CustomerTableName);
          }
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckDestinationsAssignedIncorrectTable() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckDestinationsAssignedIncorrectTable() -- finished");
    }
    private static void CreateStdIntBandsDataRecord()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CreateStdIntBandsDataRecord() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CreateStdIntBandsDataRecord() -- started");
        try
        {
          using (StreamReader oSr = new StreamReader(File.OpenRead(StaticVariable.DatasetsFolder + @"\" + Constants.StdIntAndBands)))
          {
              while (!oSr.EndOfStream)
              {
                string sLine = oSr.ReadLine();
                if (!string.IsNullOrEmpty(sLine) && !sLine.StartsWith(";"))
                {
                    StaticVariable.StandardInternationalBands.Add(new StandardInternationalBandsDataRecord(sLine));
                }
              }
              oSr.Close();
          }
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CreateStdIntBandsDataRecord()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "problem creating the Std Int Bands Data Record List");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "CreateStdIntBandsDataRecord() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CreateStdIntBandsDataRecord() -- finished");
    }
    public static void DisplayMissingDetails()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "DisplayMissingDetails() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "DisplayMissingDetails() -- started");
        CheckDestinationsAssignedIncorrectTable();
        CheckForNonMatchingCustomerNames();
        AddMainlandPricesToDependentCountries();
        FindMissingInternationalCountries();
        Console.WriteLine("ValidateData".PadRight(30, '.') + "DisplayMissingDetails() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "DisplayMissingDetails() -- finished");
    }
    private static void AddMainlandPricesToDependentCountries()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "AddMainlandPricesToDependentCountries() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "AddMainlandPricesToDependentCountries() -- started");
        List<string> tmpList = new List<string>();
        List<string> deptmpList = new List<string>();
        List<string> tmpMissingList = new List<string>();
        Dictionary<string, string> dependentCountriresDict = new Dictionary<string, string>();
        dependentCountriresDict.Add("CW", "Curacao\tAN");
        dependentCountriresDict.Add("CWM", "Curacao Mobile\tANM");
        dependentCountriresDict.Add("CYN", "Cyprus North\tTR");
        dependentCountriresDict.Add("CYNM", "Cyprus North Mobile\tTRM");
        dependentCountriresDict.Add("SX", "Sint Maarten\tAN");
        dependentCountriresDict.Add("SXM", "Sint Maarten Mobile\tANM");
        dependentCountriresDict.Add("ROD", "Rodriguez Island Mauritius\tMU");
        dependentCountriresDict.Add("RODM", "Rodriguez Isl Mauritius Mobile\tMUM");
        dependentCountriresDict.Add("KRF", "Khabarovsk Russian Federation\tRU");
        dependentCountriresDict.Add("NRF", "Nakhodka Russian Federation\tRU");
        dependentCountriresDict.Add("SRF", "Sakhalin Russian Federation\tRU");
        dependentCountriresDict.Add("TAT", "Tatarstan Federation\tRU");
        dependentCountriresDict.Add("SDS", "South Sudan\tSD");
        dependentCountriresDict.Add("SDSM", "South Sudan Mobile\tSDM");
        dependentCountriresDict.Add("CX", "Christmas Island\tAU");
        dependentCountriresDict.Add("CC", "Cocos Island\tAU");
        dependentCountriresDict.Add("XK", "Kosovo\tRS");
        dependentCountriresDict.Add("XKM", "Kosovo Mobile\tRSM");
        dependentCountriresDict.Add("VA", "Vatican City\tIT");

        var getBandsOnlyQuery =
          from drm in StaticVariable.CustomerDetailsDataRecord
          select drm.StdBand;

        foreach (string dd in dependentCountriresDict.Keys)
        {
          deptmpList.Add(dd.ToUpper());
        }

        var dependentCountriesRequired = deptmpList.Except(getBandsOnlyQuery);

        foreach (var mq in dependentCountriesRequired)
        {
          tmpMissingList.Add(dependentCountriresDict[mq]);
        }

        foreach (string token in tmpMissingList)
        {
          string[] dependantCountryAry = token.Split('\t');
          string parentCode = dependantCountryAry[1];

          var intQuery =
              from drm in StaticVariable.CustomerDetailsDataRecord
              where drm.StdBand.ToUpper().Equals(parentCode.ToUpper())
              select drm;

          foreach (var tok in intQuery)
          {
              tmpList.Add(dependantCountryAry[0] + "\t" + tok.CustomerFirstInitialRate + "\t" + tok.CustomerFirstSubseqRate + "\t" +
              tok.CustomerSecondInitialRate + "\t" + tok.CustomerSecondSubseqRate + "\t" + tok.CustomerThirdInitialRate + "\t" +
              tok.CustomerThirdSubseqRate + "\t" + tok.CustomerFourthInitialRate + "\t" + tok.CustomerFourthSubseqRate + "\t" +
              tok.CustomerMinCharge + "\t" + tok.CustomerConnectionCost + "\t" + tok.CustomerUsingGroupBands + "\t" +
              tok.CustomerGroupBand + "\t" + tok.CustomerGroupBandDescription + "\t" + CapitaliseWord(tok.CustomerTableName) + "\t" +
              CapitaliseWord(tok.CustomerDestinationType) + "\t" + tok.CustomerRounding + "\t" + CapitaliseWord(tok.CustomerTimeScheme) + "\t" +
              tok.CustomerUsingCustomerNames + "\t" + tok.CustomerInitialIntervalLength + "\t" + tok.CustomerSubsequentIntervalLength + "\t" +
              tok.CustomerMinimumIntervals + "\t" + tok.CustomerIntervalsAtInitialCost + "\t" + tok.CustomerMinimumTime + "\t" +
              tok.CustomerDialTime + "\t" + tok.CustomerAllSchemes + "\t" + tok.CustomerMultiLevelEnabled + "\t" +
              tok.CustomerMinDigits + "\t" + tok.CustomerCutOff1Cost + "\t" + tok.CustomerCutOff2Duration + "\t" +
              tok.ChargingType);
          }
        }
        if (tmpList.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData:AddMainlandPricesToDependentCountries()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Countries given Mainland prices");
          foreach (string price in tmpList)
          {
              StaticVariable.ProgressDetails.Add(price);
          }
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "AddMainlandPricesToDependentCountries() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "AddMainlandPricesToDependentCountries() -- finished");
    }
    public static HashSet<string> GetSourceAndDestinationNames()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "GetSourceAndDestinationNames() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "GetSourceAndDestinationNames() -- started");
        HashSet<string> sourceDestinationNames = new HashSet<string>();
        const int source = 1;
        const int destination = 2;

        foreach (var placename in StaticVariable.SourceDestinationBands)
        {
          string[] place = placename.Split('\t');
          sourceDestinationNames.Add(place[source].ToUpper());
          sourceDestinationNames.Add(place[destination].ToUpper());
        }
        Console.WriteLine("ValidateData".PadRight(30, '.') + "GetSourceAndDestinationNames() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "GetSourceAndDestinationNames() -- finished");
        return sourceDestinationNames;
    }    
    private static HashSet<string> CreateCountryExceptionsHashset()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CreateCountryExceptions() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CreateCountryExceptions() -- started");
      HashSet<string> prefixNameExceptions = new HashSet<string>();
      string[] countries = Constants.SpecialCountries.Split(',');
      foreach(var country in countries)
      {
        prefixNameExceptions.Add(country.Trim());
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CreateCountryExceptions() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CreateCountryExceptions() -- finished");
      return prefixNameExceptions;
    }
    private static void CheckMaximumPrices()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMaximumPrices() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMaximumPrices() -- started");
      StringBuilder sb = new StringBuilder();
      var firstInitialRatequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerFirstInitialRate)).Distinct();     
      sb.Append(Constants.FiveSpacesPadding + "Rate1 Initial = " + firstInitialRatequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var firstSubseqRatequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerFirstSubseqRate)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Rate1 Subsequent = " + firstSubseqRatequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var secondInitialRateequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerSecondInitialRate)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Rate2 Initial = " + secondInitialRateequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var secondSubseqRatequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerSecondSubseqRate)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Rate2 Subsequent = " + secondSubseqRatequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var thirdInitialRatequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerThirdInitialRate)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Rate3 Initial = " + thirdInitialRatequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var thirdSubseqRatequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerThirdSubseqRate)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Rate3 Subsequent = " + thirdSubseqRatequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var fourthInitialRatequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerFourthInitialRate)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Rate4 Initial = " + fourthInitialRatequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var fourthSubseqRatequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerFourthSubseqRate)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Rate4 Subsequent = " + fourthSubseqRatequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var minChargequery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerMinCharge)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Minimum Cost = " + minChargequery.ToList().Max(m => double.Parse(m)) + Environment.NewLine);

      var connectionCostquery =
        (from db in StaticVariable.CustomerDetailsDataRecord
        select SetToFourDecimalPlaces(db.CustomerConnectionCost)).Distinct();
      sb.Append(Constants.FiveSpacesPadding + "Connection Cost = " + connectionCostquery.ToList().Max(m => double.Parse(m))); 
       
      StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckMaximumPrices()");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The maximum prices are:"); 
      StaticVariable.ProgressDetails.Add(sb.ToString());             
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckMaximumPrices() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckMaximumPrices() -- finished");
    }
    private static void CheckIfInitialIntervalSameAsSubsequentInterval()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfInitialIntervalSameAsSubsequentInterval() -- started");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIfInitialIntervalSameAsSubsequentInterval() -- started");      
      var rateQuery =
        from db in StaticVariable.CustomerDetailsDataRecord
        where !db.CustomerFirstInitialRate.Equals(db.CustomerFirstSubseqRate) ||
              !db.CustomerSecondInitialRate.Equals(db.CustomerSecondSubseqRate) ||
              !db.CustomerThirdInitialRate.Equals(db.CustomerThirdSubseqRate) ||
              !db.CustomerFourthInitialRate.Equals(db.CustomerFourthSubseqRate)
        select db.CustomerPrefixName;

      if(rateQuery.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ValidateData::CheckIfInitialIntervalSameAsSubsequentInterval()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Initial and Subsequent prices are different:");
        foreach(var name in rateQuery)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + name);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "CheckIfInitialIntervalSameAsSubsequentInterval() -- finished");
      StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "CheckIfInitialIntervalSameAsSubsequentInterval() -- finished");
    }
    public static void TestMethod()
    {
        Console.WriteLine("ValidateData".PadRight(30, '.') + "TestMethod() -- started");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "TestMethod() -- started");
        StaticVariable.ProgressDetails.Add("ValidateData::TextMethod()");
        List<string> nationalPrefixes =  new List<string>();
        const int tableName = 0;
        const int prefix = 1;
        const int prefixName = 2;

        foreach(var name in StaticVariable.PrefixNumbersFromIniFiles)
        {
          string[] nationalOnly = name.Split('\t');
          if(nationalOnly[tableName].ToUpper().Equals(StaticVariable.NationalTableSpellingValue.ToUpper()) && !nationalOnly[prefix].Equals("?"))
          {
              nationalPrefixes.Add(nationalOnly[prefixName] + "," + nationalOnly[prefix]);
          }
        }
        nationalPrefixes.Sort();
        foreach(var item in nationalPrefixes)
        {
          StaticVariable.ProgressDetails.Add(item);
        }         
        Console.WriteLine("ValidateData".PadRight(30, '.') + "TestMethod() -- finished");
        StaticVariable.ConsoleOutput.Add("ValidateData".PadRight(30, '.') + "TestMethod() -- finished");
    }
  }
}
