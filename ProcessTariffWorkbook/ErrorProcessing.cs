// <copyright file="ErrorProcessing.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 

using System.Collections.Generic;
using System.Linq;

namespace ProcessTariffWorkbook
{
  using System;
  using System.Diagnostics.CodeAnalysis;
  using System.IO;
  using System.Text;  
  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600:ElementsMustBeDocumented", Justification = "Suppress description for each element")]
  public static class ErrorProcessing
  {
    public static void OutputToErrorLog()
    {
      //// Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "OutputToErrorLog()");
      StaticVariable.ErrorLogFile = StaticVariable.DirectoryName + @"\" + Constants.ErrorLog;
      try
      {
          if (File.Exists(StaticVariable.ErrorLogFile))
          {
              File.Delete(StaticVariable.ErrorLogFile);
          }
      }
      catch (Exception e)
      {
          Console.WriteLine("ErrorProcessing::CreateErrorLog()");
          Console.WriteLine(Constants.FiveSpacesPadding + "Error Log could not be deleted");
          Console.WriteLine(Constants.FiveSpacesPadding + e.Message);
          StopProcessDueToFatalError();
      }

      try
      {
          using (StreamWriter oSw = new StreamWriter(File.OpenWrite(StaticVariable.ErrorLogFile), Encoding.Unicode))
          {
              foreach (string token in StaticVariable.Errors)
              {
                  oSw.WriteLine(token);
              }

              oSw.Close();
          }
      }
      catch (IOException io)
      {
          Console.WriteLine("ErrorProcessing::OutputToErrorLog() -- io exception");
          Console.WriteLine(Constants.FiveSpacesPadding + "Error Log could not be opened");
          Console.WriteLine(Constants.FiveSpacesPadding + io.Message);
          StopProcessDueToFatalError();
      }
      catch (Exception e)
      {
          Console.WriteLine("ErrorProcessing::OutputToErrorLog() -- general exception");
          Console.WriteLine(Constants.FiveSpacesPadding + e.Message);
          StopProcessDueToFatalError();
      }
      //// Console.WriteLine("Errors".PadRight(30, '.') + "OutputToErrorLog()-- finished");
    }    
    public static void StopProcessDueToFatalError()
    {
      //// Console.WriteLine("ErrorProcessing::StopProcessDueToFatalError()");
      Console.WriteLine(Environment.NewLine + Constants.FiveSpacesPadding + "Process stopped due to error. See Console error");
      Console.ReadKey(true);
      Environment.Exit(Constants.KillProgram);
    }    
    public static void StopProcessDueToFatalErrorOutputToLog()
    {
      //// Console.WriteLine("ErrorProcessing::StopProcessDueToFatalErrorOutputToLog");
      OutputToErrorLog();
      Console.WriteLine(Environment.NewLine + Constants.FiveSpacesPadding + "Process stopped due to error. See Error Log");
      Console.ReadKey(true);
      Environment.Exit(Constants.KillProgram);
    }        
    public static void AddRequiredDataDetailsToErrorLog()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "AddRequiredDataDetailsToErrorLog() -- starting");
      StaticVariable.Errors.Add("Default Headers details listed below.." + Environment.NewLine);
      foreach (string tok in StaticVariable.TwbHeader)
      {
         StaticVariable.Errors.Add(tok);
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "AddRequiredDataDetailsToErrorLog() -- finishing");
    }           
    public static void CreateIntermediateLog()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "CreateIntermediateLog()-- started");
      try
      {
        StaticVariable.IntermediateLog = StaticVariable.DirectoryName + @"\" + Constants.IntermediateLog;
        if (File.Exists(StaticVariable.IntermediateLog))
        {
           File.Delete(StaticVariable.IntermediateLog);
        }

        File.Create(StaticVariable.IntermediateLog);
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ErrorProcessing::CreateIntermediateLog()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.IntermediateLog + ": Problem creating this sheet");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "CreateIntermediateLog()-- finished");
    }    
    public static void OutputConsoleLog()
    {
      string consoleOutputLogFile = StaticVariable.DirectoryName + @"\ConsoleOutput.log";
      try
      {
        if (File.Exists(consoleOutputLogFile))
        {
            File.Delete(consoleOutputLogFile);
        }
          //// File.Create(consoleOutputLogFile);        
      }
      catch (Exception e)
      {
        Console.WriteLine("ErrorProcessing::OutputConsoleLog");
        Console.WriteLine(Constants.FiveSpacesPadding + "Error Log could not be deleted");
        Console.WriteLine(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalError();
      }

      try
      {
        using (StreamWriter oSw = new StreamWriter(File.OpenWrite(consoleOutputLogFile), Encoding.Unicode))
        {
          foreach (string token in StaticVariable.ConsoleOutput)
          {
              oSw.WriteLine(token);
          }
          oSw.Close();
        }
      }
      catch (IOException io)
      {
        Console.WriteLine("ErrorProcessing::OutputConsoleLog() -- io exception");
        Console.WriteLine(Constants.FiveSpacesPadding + "Console Log could not be opened");
        Console.WriteLine(Constants.FiveSpacesPadding + io.Message);
        StopProcessDueToFatalError();
      }
      catch (Exception e)
      {
        Console.WriteLine("ErrorProcessing::OutputConsoleLog() -- general exception");
        Console.WriteLine(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalError();
      }
    }
    public static void AddMainlandPricesToDependentCountries()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "AddMainlandPricesToDependentCountries()-- started");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "AddMainlandPricesToDependentCountries()");
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
      dependentCountriresDict.Add("VA", "Vatican City\tIT"); //CustomerDetailsDataRecord

      var getBandsOnlyQuery =
        from drm in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
        select drm.StdBand;

      foreach (string dd in dependentCountriresDict.Keys)
      {
        deptmpList.Add(dd.ToUpper());
      }

      var dependentCountriesRequired = deptmpList.Except(getBandsOnlyQuery); //get only those dependent countries not in input xls

      foreach (var mq in dependentCountriesRequired)
      {
        tmpMissingList.Add(dependentCountriresDict[mq]);
      }

      foreach (string token in tmpMissingList)
      {
        string[] dependantCountryAry = token.Split('\t');
        string parentCode = dependantCountryAry[1];

        var intQuery =
          from drm in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
          where drm.StdBand.ToUpper().Equals(parentCode.ToUpper())
          select drm;

        foreach (var tok in intQuery)
        {
          tmpList.Add(dependantCountryAry[0] + "\t" + tok.CustomerFirstInitialRate + "\t" + tok.CustomerFirstSubseqRate + "\t" +
            tok.CustomerSecondInitialRate + "\t" + tok.CustomerSecondSubseqRate + "\t" + tok.CustomerThirdInitialRate + "\t" +
            tok.CustomerThirdSubseqRate + "\t" + tok.CustomerFourthInitialRate + "\t" + tok.CustomerFourthSubseqRate + "\t" +
            tok.CustomerMinCharge + "\t" + tok.CustomerConnectionCost + "\t" + tok.CustomerUsingGroupBands + "\t" +
            tok.CustomerGroupBand + "\t" + tok.CustomerGroupBandDescription + "\t" + ValidateData.CapitaliseWord(tok.CustomerTableName) + "\t" +
            ValidateData.CapitaliseWord(tok.CustomerDestinationType) + "\t" + tok.CustomerRounding + "\t" + ValidateData.CapitaliseWord(tok.CustomerTimeScheme) + "\t" +
            tok.CustomerUsingCustomerNames + "\t" + tok.CustomerInitialIntervalLength + "\t" + tok.CustomerSubsequentIntervalLength + "\t" +
            tok.CustomerMinimumIntervals + "\t" + tok.CustomerIntervalsAtInitialCost + "\t" + tok.CustomerMinimumTime + "\t" +
            tok.CustomerDialTime + "\t" + tok.CustomerAllSchemes + "\t" + tok.CustomerMultiLevelEnabled + "\t" +
            tok.CustomerMinDigits + "\t" + tok.CustomerCutOff1Cost + "\t" + tok.CustomerCutOff2Duration + "\t" +
            tok.ChargingType);
        }
      }
      if (tmpList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ErrorProcessing:AddMainlandPricesToDependentCountries()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Countries given Mainland prices");
        foreach (string price in tmpList)
        {
          StaticVariable.Errors.Add(price);
        }
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "AddMainlandPricesToDependentCountries()-- finished");
    }
    public static void FindMissingInternationalCountries()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "FindMissingInternationalCountries() -- started");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "FindMissingInternationalCountries()");
      List<string> tmpList = new List<string>();
      List<string> noPricesForCountriesList = new List<string>();
      List<string> bandsNotFoundList = new List<string>();
      string custband = string.Empty;
      const string zeroPrices = "0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000\t0.0000";
      string otherInfo = "FALSE\tnull\tnull\t" + StaticVariable.CountryCode + "_International\tInternational\tRounding\tTimeScheme\tFALSE\t60\t60\t0\t1\t1\t1\tTRUE\tFALSE\t0\t0\t0\tCharging Type";

      CreateStdIntBandsDataRecord();
      //from regexmatched get all int destinations into tmp list
      var intQuery =
        from drm in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
        where drm.CustomerTableName.ToUpper().Equals(StaticVariable.InternationalTableSpelling.ToUpper())
        select new { drm.StdPrefixName, drm.StdBand, drm.CustomerPrefixName };

      var customerCountryList = intQuery.Select(q => q.StdBand + "\t" + q.StdPrefixName + "\t" + q.CustomerPrefixName).ToList();
      //check if all int bands are in input list.
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
      //find out if a fixed band exists. yes give it landline price.
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
            from drm in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
            where tmpCustBand.ToUpper().Equals(drm.StdBand.ToUpper())
            select drm;
          foreach (var dr in mobileQuery)
          {
            string name = dr.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? dr.CustomerPrefixName : dr.StdPrefixName;
            string groupBand = dr.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? dr.CustomerGroupBand + mobileBand : dr.StdBand + mobileBand;
            string groupBandDescription = dr.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? dr.CustomerGroupBandDescription + mobileBand : dr.StdPrefixDescription + mobileBand;
            tmpList.Add(name + " " + StaticVariable.IntMobileSpelling + "\t" + dr.CustomerFirstInitialRate + "\t" +
              dr.CustomerFirstSubseqRate + "\t" + dr.CustomerSecondInitialRate + "\t" + dr.CustomerSecondSubseqRate + "\t" +
              dr.CustomerThirdInitialRate + "\t" + dr.CustomerThirdSubseqRate + "\t" + dr.CustomerFourthInitialRate + "\t" +
              dr.CustomerFourthSubseqRate + "\t" + dr.CustomerMinCharge + "\t" + dr.CustomerConnectionCost + "\t" +
              "FALSE\t" + groupBand + "\t" + groupBandDescription + "\t" + ValidateData.CapitaliseWord(dr.CustomerTableName) + "\t" +
              ValidateData.CapitaliseWord(dr.CustomerDestinationType + " " + StaticVariable.IntMobileSpelling) + "\t" + 
              dr.CustomerRounding + "\t" + ValidateData.CapitaliseWord(dr.CustomerTimeScheme) + "\t" +
              dr.CustomerUsingCustomerNames + "\t" + dr.CustomerInitialIntervalLength + "\t" + dr.CustomerSubsequentIntervalLength + "\t" + 
              dr.CustomerMinimumIntervals + "\t" + dr.CustomerIntervalsAtInitialCost + "\t" + dr.CustomerMinimumTime + "\t" + dr.CustomerDialTime + "\t" +
              dr.CustomerAllSchemes + "\t" + dr.CustomerMultiLevelEnabled + "\t" + dr.CustomerMinDigits + "\t" +
              dr.CustomerCutOff1Cost + "\t" + dr.CustomerCutOff2Duration + "\t" + dr.ChargingType
              );
          }
        }
        else if (!stdCountryCode.Equals(StaticVariable.CountryCode) && !stdCountryCode.ToUpper().Equals("N/A"))
        {
          noPricesForCountriesList.Add(ValidateData.CapitaliseWord(custName) + "\t" + zeroPrices + "\t" + otherInfo);
        }
      }
      if (tmpList.Any())
      {
        tmpList.Sort();
        StaticVariable.Errors.Add(Environment.NewLine + "ErrorProcessing:FindMissingInternationalCountries()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Int mobile prices missing for these destinations, so given landine prices.");
        foreach (string sEntry in tmpList)
        {
          StaticVariable.Errors.Add(sEntry);
        }
      }

      if (noPricesForCountriesList.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ErrorProcessing:FindMissingInternationalCountries()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "No prices supplied for these countries. They have been given a zero price.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "If a BT / Eir / KPN price is available for these destinations then use that price unless told to zero cost these destinations.");
        noPricesForCountriesList.Sort();
        foreach (string sZeroEntry in noPricesForCountriesList)
        {
          StaticVariable.Errors.Add(sZeroEntry);
        }
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "FindMissingInternationalCountries() -- finished");
    }
    public static void DestinationsAssignedIncorrectTable()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "DestinationsAssignedIncorrectTable()-- started");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "DestinationsAssignedIncorrectTable()");

      var query =
        from drm in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
        join pn in StaticVariable.PrefixNumbersRecord on drm.StdPrefixName.ToUpper() equals pn.StandardPrefixName.ToUpper()
        where !drm.CustomerTableName.ToUpper().Equals(pn.TableName.ToUpper())
        select new { pn.TableName, pn.StandardPrefixName, drm.CustomerTableName, drm.CustomerPrefixName };

      if (query.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ErrorProcessing::DestinationsAssignedIncorrectTable()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Prefix Name and prefix are assigned different tables.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Perhaps check the RegExMatchedList for incorrect regex matching." + Environment.NewLine);
        foreach (var entry in query)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + ("Prefix Table: " + entry.StandardPrefixName + " --> " + entry.TableName).PadRight(70) + "XLSX File: " + entry.CustomerPrefixName + " --> " + entry.CustomerTableName);
        }
        StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "DestinationsAssignedIncorrectTable()-- finished");
    }
    public static void DestinationsWithoutPrefixes()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "DestinationsWithoutPrefixes()-- started");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "DestinationsWithoutPrefixes()");
      Dictionary<string, string> customerNames = new Dictionary<string, string>();
      List<string> stdNames = new List<string>();

      var queryCustomerDetails =
        (from dr in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
         select new { dr.StdPrefixName, dr.CustomerPrefixName }).Distinct();

      foreach (var variable in queryCustomerDetails)
      {
        try
        {
          customerNames.Add(variable.StdPrefixName.ToUpper(), variable.CustomerPrefixName);
        }
        catch (Exception e)
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ErrorProcessing::DestinationsWithoutPrefixes()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Problems adding values to dictionary");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
          StopProcessDueToFatalErrorOutputToLog();
        }        
        stdNames.Add(variable.StdPrefixName.ToUpper());
      }

      var queryPrefixes =
        (from pn in StaticVariable.PrefixNumbersRecord
         select pn.StandardPrefixName.ToUpper()).Distinct();

      var missingPrefixes = stdNames.Except(queryPrefixes);

      if (missingPrefixes.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ErrorProcessing::DestinationsWithoutPrefixes()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "No Prefix Found:");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The std  or customer prefix name (if 'using customer name' is Yes) may not match the name in the appropriate prefix table or else the prefix may not exist in that table.");
        foreach (var entry in missingPrefixes)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Customer Name: " + customerNames[entry].PadRight(41, ' ') + ".".PadRight(20, '.') + " : Standard Name: " + ValidateData.CapitaliseWord(entry));
        }
        //StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "DestinationsWithoutPrefixes()-- finished");
    }
    private static void CreateStdIntBandsDataRecord()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "CreateStdIntBandsDataRecord() -- started");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "CreateStdIntBandsDataRecord()");
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
        StaticVariable.Errors.Add(Environment.NewLine + "ErrorProcessing::CreateStdIntBandsDataRecord()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "problem creating the Std Int Bands Data Record List");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "CreateStdIntBandsDataRecord() -- finished");
    }
    public static void WriteToIntermediateLog()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "RearrangeToIntermediateLog()-- started");
      StaticVariable.ConsoleOutput.Add(Environment.NewLine + "ErrorProcessing".PadRight(30, '.') + "RearrangeToIntermediateLog()");
      var allDetails =
        from db in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
        select db;
      try
      {
        using (StreamWriter oSw = new StreamWriter(File.OpenWrite(StaticVariable.IntermediateLog), Encoding.Unicode))
        {
          oSw.WriteLine("Standard Band\tStandard Name\tCustomer Name\tGroup Band\tGroup Band Description\tTable Name\tDestination Type\tStandard Description\tTime Scheme\tRounding\tRates");
          foreach (var ad in allDetails)
          {
            oSw.WriteLine(ad.StdBand.PadRight(4, ' ') + "\t" + ad.StdPrefixName.PadRight(40, ' ') + "\t" + ad.CustomerPrefixName.PadRight(40, ' ') + "\t" +
                    ad.CustomerGroupBand.PadRight(4, ' ') + "\t" + ad.CustomerGroupBandDescription.PadRight(20, ' ') + "\t" + ad.CustomerTableName.PadRight(20, ' ') + "\t" +
                    ad.CustomerDestinationType.PadRight(20, ' ') + "\t" + ad.StdPrefixDescription.PadRight(20, ' ') + "\t" + ad.CustomerTimeScheme + "\t" +
                    ad.CustomerRounding + "\t" + ad.CustomerFirstInitialRate + "\t" + ad.CustomerFirstSubseqRate + "\t" + ad.CustomerSecondInitialRate + "\t" +
                    ad.CustomerSecondSubseqRate + "\t" + ad.CustomerThirdInitialRate + "\t" + ad.CustomerThirdSubseqRate + "\t" + ad.CustomerFourthInitialRate + "\t" +
                    ad.CustomerFourthSubseqRate + "\t" + ad.CustomerMinCharge + "\t" + ad.CustomerConnectionCost + "\t" + ad.CustomerInitialIntervalLength + "\t" +
                    ad.CustomerSubsequentIntervalLength + "\t" + ad.CustomerMinimumIntervals + "\t" + ad.CustomerIntervalsAtInitialCost + "\t" +
                    ad.CustomerMinimumTime + "\t" + ad.CustomerMinDigits + "\t" + ad.CustomerUsingCustomerNames + "\t" + ad.CustomerUsingGroupBands + "\t" +
                    ad.CustomerMultiLevelEnabled + "\t" + ad.CustomerCutOff1Cost + "\t" + ad.CustomerCutOff2Duration + "\t" + ad.ChargingType);
          }
          oSw.Close();
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ErrorProcessing::RearrangeToIntermediateLog()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.IntermediateLog + ": Problem writing to Intermediate File");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "RearrangeToIntermediateLog()-- finished");
    }
    public static void WriteOutGroupBandsToErrorLog()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "WriteOutGroupBands() -- started");
      StaticVariable.Errors.Add(Environment.NewLine + "ValidateData:WriteOutGroupBands()");
      StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "group bands used........." + Environment.NewLine);

      try
      {
        var durationQuery =
        (from groups in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
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

        foreach (var dg in durationQuery)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Charging Type: " + dg.ChargingType);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Band:".PadRight(15, '.') + '\x0020' + dg.CustomerGroupBand.PadRight(10, '\x0020') + " Band Description:".PadRight(20, '.') + '\x0020' + dg.CustomerGroupBandDescription);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Min Cost:".PadRight(15, '.') + '\x0020' + dg.CustomerMinCharge.PadRight(10, '\x0020') + " Connection Cost:".PadRight(20, '.') + '\x0020' + dg.CustomerConnectionCost);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Cheap_1:".PadRight(15, '.') + '\x0020' + dg.CustomerFirstInitialRate.PadRight(10, '\x0020') + " Cheap_2:".PadRight(20, '.') + '\x0020' + dg.CustomerFirstSubseqRate);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Standard_1:".PadRight(15, '.') + '\x0020' + dg.CustomerSecondInitialRate.PadRight(10, '\x0020') + " Standard_2:".PadRight(20, '.') + '\x0020' + dg.CustomerSecondSubseqRate);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Economy_1:".PadRight(15, '.') + '\x0020' + dg.CustomerThirdInitialRate.PadRight(10, '\x0020') + " Economy_2:".PadRight(20, '.') + '\x0020' + dg.CustomerThirdSubseqRate);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Peak_1:".PadRight(15, '.') + '\x0020' + dg.CustomerFourthInitialRate.PadRight(10, '\x0020') + " Peak_2:".PadRight(20, '.') + '\x0020' + dg.CustomerFourthSubseqRate);
          StaticVariable.Errors.Add("\n");
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ValidateData: WriteOutGroupBands()");
        StaticVariable.Errors.Add("Error in WriteOutGroupBands() - duration");
        StaticVariable.Errors.Add(e.Message);
      }
      try
      {
        var cappedQuery =
        (from groups in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
         where groups.CustomerUsingGroupBands.ToUpper().Equals("TRUE") && groups.ChargingType.ToUpper().Equals("CAPPED")
         select groups).Distinct();
        foreach (var cq in cappedQuery)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Charging Type: " + cq.ChargingType);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Band:".PadRight(15, '.') + '\x0020' + cq.CustomerGroupBand.PadRight(10, '\x0020') + " Band Description:".PadRight(20, '.') + '\x0020' + cq.CustomerGroupBandDescription);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Min Cost:".PadRight(15, '.') + '\x0020' + cq.CustomerMinCharge.PadRight(10, '\x0020') + " Connection Cost:".PadRight(20, '.') + '\x0020' + cq.CustomerConnectionCost);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Standard:".PadRight(15, '.') + '\x0020' + cq.CustomerFirstInitialRate.PadRight(10, '\x0020') + " Cap Price:".PadRight(20, '.') + '\x0020' + cq.CustomerFirstSubseqRate);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Cap Time:".PadRight(15, '.') + '\x0020' + cq.CustomerSecondInitialRate.PadRight(10, '\x0020'));
          StaticVariable.Errors.Add("\n");
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ValidateData: WriteOutGroupBands()");
        StaticVariable.Errors.Add("Error in WriteOutGroupBands() - capped");
        StaticVariable.Errors.Add(e.Message);
      }
      try
      {
        var pulseQuery =
        (from groups in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
         where groups.CustomerUsingGroupBands.ToUpper().Equals("TRUE") && groups.ChargingType.ToUpper().Equals("PULSE")
         select groups).Distinct();
        foreach (var pq in pulseQuery)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Charging Type: " + pq.ChargingType);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Band:".PadRight(15, '.') + '\x0020' + pq.CustomerGroupBand.PadRight(10, '\x0020') + " Band Description:".PadRight(20, '.') + '\x0020' + pq.CustomerGroupBandDescription);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Min Cost:".PadRight(15, '.') + '\x0020' + pq.CustomerMinCharge.PadRight(10, '\x0020') + " Connection Cost:".PadRight(20, '.') + '\x0020' + pq.CustomerConnectionCost);
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Pulse Length:".PadRight(15, '.') + '\x0020' + pq.CustomerFirstInitialRate.PadRight(10, '\x0020') + " Pulse Unit:".PadRight(20, '.') + '\x0020' + pq.CustomerFirstSubseqRate);
          StaticVariable.Errors.Add("\n");
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ValidateData: WriteOutGroupBands()");
        StaticVariable.Errors.Add("Error in WriteOutGroupBands() - pulse");
        StaticVariable.Errors.Add(e.Message);
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "WriteOutGroupBands() -- finish");
    }
    public static void WriteToErrorlogIfMinCostAnd4ThRateSamePrice()
    {
      Console.WriteLine("ValidateData".PadRight(30, '.') + "DisplayMinCostV4thRateSamePrice() -- started");
      List<string> matches = new List<string>();
      var query =
        from peakAndMinCost in StaticVariable./*DestinationsMatchedByRegExDataRecord*/CustomerDetailsDataRecord
        where peakAndMinCost.CustomerFourthSubseqRate.Equals(peakAndMinCost.CustomerMinCharge)
        select new { peakAndMinCost.StdPrefixName, peakAndMinCost.CustomerFourthSubseqRate, peakAndMinCost.CustomerMinCharge, peakAndMinCost.CustomerPrefixName };

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
        StaticVariable.Errors.Add(Environment.NewLine + "ValidateData:DisplayMinCostV4thRateSamePrice()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "4th rate is the same price as Minimum Cost.");
        foreach (var m in matches)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + m);
        }
      }
      Console.WriteLine("ValidateData".PadRight(30, '.') + "DisplayMinCostV4thRateSamePrice() -- finish");
    }
  }
}
