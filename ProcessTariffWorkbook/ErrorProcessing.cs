﻿// <copyright file="ErrorProcessing.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 

using System.Collections.Generic;
using System.Linq;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text; 

namespace ProcessTariffWorkbook
{
  
  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600:ElementsMustBeDocumented", Justification = "Suppress description for each element")]
  public static class ErrorProcessing
  {
    public static void OutputToLogs(List<string> log, string logName )
    {      
      try
      {
        if (File.Exists(logName))
        {
          File.Delete(logName);
        }
      }
      catch (Exception e)
      {
        Console.WriteLine("ErrorProcessing::CreateErrorLog()");
        Console.WriteLine(Constants.FiveSpacesPadding + logName + " Log could not be deleted");
        Console.WriteLine(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalError();
      }

      try
      {
        using (StreamWriter oSw = new StreamWriter(File.OpenWrite(logName), Encoding.Unicode))
        {
          foreach (string token in log)
          {
            oSw.WriteLine(token);
          }

          oSw.Close();
        }
      }
      catch (IOException io)
      {
        Console.WriteLine("ErrorProcessing::OutputToProcessLog() -- io exception");
        Console.WriteLine(Constants.FiveSpacesPadding + "Error Log could not be opened");
        Console.WriteLine(Constants.FiveSpacesPadding + io.Message);
        StopProcessDueToFatalError();
      }
      catch (Exception e)
      {
        Console.WriteLine("ErrorProcessing::OutputToProcessLog() -- general exception");
        Console.WriteLine(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalError();
      }            
    }
    public static void StopProcessDueToFatalError()
    {      
      Console.WriteLine(Environment.NewLine + Constants.FiveSpacesPadding + "Process stopped due to error. See Console error");
      Console.ReadKey(true);
      Environment.Exit(Constants.KillProgram);
    }    
    public static void StopProcessDueToFatalErrorOutputToLog()
    {      
      OutputToLogs(StaticVariable.ProgressDetails, StaticVariable.DirectoryName + @"\" + Constants.ProgressLog);
      OutputToLogs(StaticVariable.ConsoleOutput, StaticVariable.DirectoryName + @"\" + Constants.ConsoleErrorLog);     
      Console.WriteLine(Environment.NewLine + Constants.FiveSpacesPadding + "Process stopped due to error. See Error Log");
      Console.ReadKey(true);
      Environment.Exit(Constants.KillProgram);
    }        
    public static void AddRequiredDataDetailsToErrorLog()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "AddRequiredDataDetailsToErrorLog() -- started");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "AddRequiredDataDetailsToErrorLog() -- started");
      StaticVariable.ProgressDetails.Add("Default Headers details listed below.." + Environment.NewLine);      
      foreach (string tok in StaticVariable.ProgressDetails)
      {
         StaticVariable.ProgressDetails.Add(tok);
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "AddRequiredDataDetailsToErrorLog() -- finishing");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "AddRequiredDataDetailsToErrorLog() -- finishing");
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
    public static void CreateAndWriteToRegExMatchedLog()
    {
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "WriteToRegExMatchedLog() -- started");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "WriteToRegExMatchedLog() -- started");
      const string regExMatchedLog = "RegExMatchedLog.txt";
      string regExMatchedLogValue = StaticVariable.DirectoryName + @"\" + regExMatchedLog;
      var allDetails =
        from db in StaticVariable.CustomerDetailsDataRecord
        select db;
      try
      {
        using (StreamWriter oSw = new StreamWriter(File.OpenWrite(regExMatchedLogValue), Encoding.Unicode))
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
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ErrorProcessing::WriteToRegExMatchedLog()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + regExMatchedLogValue + ": Problem writing to RegExMatchedLog File");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ErrorProcessing".PadRight(30, '.') + "WriteToRegExMatchedLog() -- finished");
      StaticVariable.ConsoleOutput.Add("ErrorProcessing".PadRight(30, '.') + "WriteToRegExMatchedLog() -- finished");
    }     
    public static void TestWriteToPrefixNumbersSheet2()
    {
      Console.WriteLine("RearrangeToSheetsAndLogs".PadRight(30, '.') + "WriteToPrefixNumbersSheet2()-- started");
      StaticVariable.ConsoleOutput.Add("RearrangeToSheetsAndLogs".PadRight(30, '.') + "WriteToPrefixNumbersSheet2()");
     // string sSheet = StaticVariable.FinalDirectory+ @"\" + StaticVariable.twb + sPrefixNumbersSheetCONST;      
      string sName = string.Empty;
      string sNumber = string.Empty;
      string sTmpTable = string.Empty;
      List<string> IncorrectTableList = new List<string>();
      List<string> tmpMissingPrefixList = new List<string>();
      List<string> tmpList = new List<string>();
      bool bFound = false;
            
          foreach (DataRecord drm in StaticVariable.CustomerDetailsDataRecord)
          {
            bFound = false;
            foreach (PrefixNumbersDataRecord pn in StaticVariable.PrefixNumbersRecord)
            {
              try
              {
                if (drm.StdPrefixName.ToUpper().Equals(pn.PrefixName.ToUpper()) || drm.CustomerPrefixName.ToUpper().Equals(pn.PrefixName.ToUpper()))
                {
                  bFound = true;
                  sTmpTable = pn.TableName;
                  sNumber = pn.PrefixNumber;
                  sName = pn.stdPrefixName;
                  if (drm.CustomerUsingCustomerNames.ToUpper().Equals("TRUE"))
                  {
                    tmpList.Add(pn.TableName + "\t" + pn.PrefixNumber + "\t" + ValidateData.CapitaliseWord(drm.CustomerPrefixName));
                  }
                  else
                  {
                    tmpList.Add(pn.TableName + "\t" + pn.PrefixNumber + "\t" + ValidateData.CapitaliseWord(drm.StdPrefixName));
                  }
                }
              }
              catch (Exception e)
              {
                StaticVariable.ProgressDetails.Add("RearrangeToSheetsAndLogs::WriteToPrefixNumbersSheet2()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + ": Problem with writing the prefix file");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
                StopProcessDueToFatalErrorOutputToLog();
              }
            }
            if (bFound)
            {
              if (!drm.CustomerTableName.ToUpper().Equals(sTmpTable.ToUpper())) // Now check the table names are the same.
              {
                IncorrectTableList.Add(sTmpTable + " : " + drm.CustomerPrefixName + " - " + sNumber);
              }
            }
            else
            {
              tmpMissingPrefixList.Add(drm.CustomerTableName + " : Customer Name : " + drm.CustomerPrefixName.PadRight(41, ' ') + " --> Standard Name : " + drm.StdPrefixName);
            }
          }          
     
      if (tmpMissingPrefixList.Count > 0)
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeToSheetsAndLogs::WriteToPrefixNumbersSheet2()");
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "No Prefix Found:");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The std  or customer prefix name (if 'using customer name' is: No / Yes) may not match the name in the appropriate prefix table or else the prefix may not exist in that table.");
        tmpMissingPrefixList.Sort();
        foreach (string error in tmpMissingPrefixList)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + error);
        }
      }
      if (IncorrectTableList.Count > 0)
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeToSheetsAndLogs::WriteToPrefixNumbersSheet2()");
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "Destinations assigned a different table to the table where the prefix is.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Check the RegExMatchedList for incorrect regex matching.");
        //StaticVariable.ProgressDetails.Add(@"StaticVariable.ProgressDetails.Add("Run the debug line:" + Enviroment.Newline + sRegExBand \t sRegexStandardName \t sRegExDescription \t + tok), in ParseInputFile::RegExMatchInputList");
        IncorrectTableList.Sort();
        foreach (string error in IncorrectTableList)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + error);
        }
        StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("RearrangeToSheetsAndLogs".PadRight(30, '.') + "WriteToPrefixNumbersSheet2()-- finished");
    }
  }
}
