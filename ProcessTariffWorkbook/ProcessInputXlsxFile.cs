// <copyright file="  class ProcessInputXlsxFile.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Timers;

namespace ProcessTariffWorkbook
{
  public class ProcessInputXlsxFile
  {
    public static void ParseInputXlsxFileIntoCustomerDetailsRecord()
    {
      ReadXlsxFileIntoList();
      MergeDefaultPricesListWithInputFileList();
      AddToCustomerDetailsDataRecordList(StaticVariable.InputXlsxFileDetails);      
    }
    private static void ReadXlsxFileIntoList()
    {      
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList() -- started");
      StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessInputXlsxFile::ReadXLSXFileIntoList()");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Any line containing 'DefaultXX' will be ignored, as will all headers");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If 'Capped' or 'Pulse' rates are not being used, delete the worksheet.");
      string[] worksheetsTypes = {Constants.Duration, Constants.Capped, Constants.Pulse};      
      List<string> workSheetsNotUsed = new List<string>();      
      List<string> discardedLines = new List<string>();
      List<string> workSheetsUsed = new List<string>();
      SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(StaticVariable.InputFile);

      foreach (string wksheet in worksheetsTypes) 
      {
        try
        {
          SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[wksheet];
          SpreadsheetGear.IRange cells = worksheet.Cells;
          workSheetsUsed.Add(wksheet);         
        }
        catch (Exception)
        {
          workSheetsNotUsed.Add(wksheet);
        }        
      }

      foreach (string wksheet in workSheetsUsed) 
      {                            
        SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[wksheet];        
        SpreadsheetGear.IRange cells = worksheet.Cells;
        var currentColumn = 0;
        for (currentColumn = 0; currentColumn < cells.ColumnCount; currentColumn++) 
        {
          if (cells[0, currentColumn].Text.ToUpper().Equals(Constants.FinalColumnName))
          {
            currentColumn++;            
            break;
          }          
        }
        var maximumNumberOfColumns = currentColumn;

        try
        {        
          foreach (SpreadsheetGear.IRange row in worksheet.UsedRange.Rows)
          {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < maximumNumberOfColumns; i++)
            {                
              sb.Append(row[0, i].Value + "\t"); //0.0400 being chopped to 0.04.                                 
            }
            string sAdjustSb = sb.ToString().TrimEnd('\t');

            if ( sAdjustSb.Contains(";") && !DiscardHeaderLine(sAdjustSb))
            {              
              discardedLines.Add("- " + sAdjustSb.Substring(0, sAdjustSb.IndexOf('\t')));              
            }
            else if (!string.IsNullOrEmpty(sAdjustSb) && !DiscardHeaderLine(sAdjustSb))
            {
              ValidateData.CheckForCommasInLine(sAdjustSb);
              StaticVariable.InputXlsxFileDetails.Add(ValidateData.CapitaliseWord(sAdjustSb));              
            }
          }
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessInputXlsxFile::ReadXLSXFileIntoList()");          
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error in reading in XLSX line into list. Is there any data? " );
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        }
      }
      workbook.Close();
      if (workSheetsNotUsed.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessInputXlsxFile::ReadXLSXFileIntoList()");
        foreach (var entry in workSheetsNotUsed)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + entry + " rates are not being used.");
        }
      }
      if (discardedLines.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessInputXlsxFile::ReadXLSXFileIntoList()");        
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Customer destinations discarded.");
        discardedLines.Sort();
        foreach (var entry in discardedLines)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + entry);
        }
      }      
      StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList()-- completed");
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList() -- finished");
    }
    private static void MergeDefaultPricesListWithInputFileList()
    {
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "MergeDefaultPricesListWithInputFileList() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "MergeDefaultPricesListWithInputFileList() -- started");
      foreach (string tok in StaticVariable.DefaultEntries)
      {
        StaticVariable.InputXlsxFileDetails.Add(tok);
      }
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "MergeDefaultPricesListWithInputFileList() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "MergeDefaultPricesListWithInputFileList() -- finished");
    }
    private static void AddToCustomerDetailsDataRecordList(List<string> lst)
    {
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "AddToPreRegExDataRecordList() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "AddToPreRegExDataRecordList() -- started");
      const string undefinedStandardInfo = "undefinedStandardBand\tundefinedStandardName\tundefinedStandardDestination\t";
      foreach (string token in lst)
      {
        try
        {                    
          StaticVariable.CustomerDetailsDataRecord.Add(new DataRecord(undefinedStandardInfo + token));          
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add("ProcessInputXlsxFile::AddToPreRegExDataRecordList()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem adding list to DataRecord");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "AddToPreRegExDataRecordList() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "AddToPreRegExDataRecordList() -- finished");
    }
    public static void MatchInputXlsxFileWithRegEx(List<string> listToUse)
    {
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "MatchInputXlsxFileWithRegEx() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "MatchInputXlsxFileWithRegEx() -- started");            
      string destinationName = string.Empty;
      var tmpList = new List<string>();
      const int destination = 0;
      int uniqueBandCounter = 100;
      Timer newTimer = new System.Timers.Timer(10000); // 2 sec interval
      string regExAlphanumeric = @"[0-9|a-z|A-Z]";
      string regExExtraSpaces = "\x0020{2,}";
      string regExNull = "\x0000";
      string regExNoise = @"\(|\)|( -|- )|\,|'";      
      Regex regExRemoveNull = new Regex(regExNull, RegexOptions.Multiline);
      Regex regExRemoveNoise = new Regex(regExNoise, RegexOptions.Compiled);
      Regex regExRemoveExtraSpaces = new Regex(regExExtraSpaces, RegexOptions.Compiled);
      Regex regExCheckForAlphanumeric = new Regex(regExAlphanumeric, RegexOptions.Compiled);

      foreach (string tok in listToUse)
      {
        var found = false;
        if (!tok.StartsWith(";"))
        {
          string[] aryLine;
          try
          {
            aryLine = tok.Split('\t');
            destinationName = aryLine[destination].Trim();
            destinationName = regExRemoveNoise.Replace(destinationName, " ");
            destinationName = regExRemoveExtraSpaces.Replace(destinationName, " ");
          }
          catch (Exception e)
          {
            StaticVariable.ProgressDetails.Add("ParseInputFile::MatchInputXlsxFileWithRegEx()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There may be a problem with the input destination");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          foreach (string regexExpression in StaticVariable.CombinedRegex)
          {
            aryLine = regexExpression.Split(new char[] {'\t'});
            string regExPattern = aryLine[0];
            string regExBand = aryLine[1].Trim();
            string regexStandardName = aryLine[2].Trim();
            string regExDescription = aryLine[3].Trim();

            try
            {
              var regExCountry = new Regex(regExPattern, RegexOptions.IgnoreCase);
              if (regExCountry.IsMatch(destinationName))
              {
                newTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
                newTimer.AutoReset = true;
                newTimer.Enabled = true;
                found = true;
                
                StaticVariable.CustomerDetailsDataRecord.Add(new DataRecord(regExBand + "\t" + regexStandardName + "\t" + regExDescription + "\t" + tok));
                // for debugging
                //StaticVariable.ProgressDetails.Add("ParseInputFile::MatchInputXlsxFileWithRegEx()  -- Debugging only");
                //StaticVariable.ProgressDetails.Add(regExBand + "\tMatchInputXlsxFileWithRegEx()\t" + regExDescription + "\t" + tok);
              }
            }
            catch (Exception e)
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessInputXlsxFile::MatchInputXlsxFileWithRegEx()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There may be a problem with the regex");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Check if there are more than 1 international or domestic regex files");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
        }
        if (!found)
        {
          tmpList.Add("^" + destinationName.Replace(" ", ".").TrimEnd(' ') + "\tU" + uniqueBandCounter + "\t" + destinationName + "\t" + destinationName.PadRight(20, ' ').Substring(0, 20));
          uniqueBandCounter++;
        }
      }
      if (tmpList.Count > 0)
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessInputXlsxFile::MatchInputXlsxFileWithRegEx()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Destinations without regex");        
        tmpList.Sort();
        foreach (string s in tmpList)
        {
          StaticVariable.ProgressDetails.Add(s);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      newTimer.Enabled = false;
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "MatchInputXlsxFileWithRegEx() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "MatchInputXlsxFileWithRegEx() -- finished");
    }
    private static void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
    {
      Console.Write(".");
    }
    private static bool DiscardHeaderLine(string line)
    {      
      return line.ToUpper().Contains("DESTINATION") || line.ToUpper().Contains("TABLE") || line.ToUpper().Contains("USING") || line.ToUpper().Contains("DEFAULTXX");      
    }
    
  }
}
