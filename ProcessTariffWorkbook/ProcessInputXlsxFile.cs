// <copyright file="  class ProcessInputXlsxFile.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Timers;

namespace ProcessTariffWorkbook
{
  public class ProcessInputXlsxFile
  {
    public static void ParseInputXlsxFile()
    {
      ReadXlsxFileIntoList();
      MergeDefaultPricesListWithInputFileList();
      AddToCustomerDetailsDataRecordList(StaticVariable.InputXlsxFileDetails);
      ValidateData.PreRegExDataRecordValidate();
      StaticVariable.CustomerDetailsDataRecord.Clear();
      MatchInputXlsxFileWithRegExAndAddToDestinationsMatchedByRegExDataRecord();
      ValidateData.PostRegExDataRecordValidate();
    }
    private static void ReadXlsxFileIntoList()
    {      
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList()");
      string[] worksheetsTypes = {Constants.Duration, Constants.Capped, Constants.Pulse};      
      List<string> workSheetNotUsed = new List<string>();
      List<string> workSheetsUsed = new List<string>();
      List<string> discardedLines = new List<string>();
      SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(StaticVariable.InputFile);

      foreach (string wksheet in worksheetsTypes) //get all work sheets present
      {
        try
        {
          SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[wksheet];
          SpreadsheetGear.IRange cells = worksheet.Cells;
          workSheetsUsed.Add(wksheet);
        }
        catch (Exception)
        {
          workSheetNotUsed.Add(wksheet);
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

            if (string.IsNullOrEmpty(sAdjustSb) || sAdjustSb.Contains(";") || DiscardHeaderLine(sAdjustSb))
            {
              if (!string.IsNullOrEmpty(sAdjustSb))
              {
                discardedLines.Add("- " + sAdjustSb.Substring(0, sAdjustSb.IndexOf('\t')));
              }
            }
            else
            {
              ValidateData.CheckForCommasInPrices(sAdjustSb);
              StaticVariable.InputXlsxFileDetails.Add(sAdjustSb);
            }
          }
        }
        catch (Exception e)
        {
          StaticVariable.Errors.Add(Environment.NewLine + "ProcessInputXlsxFile::ReadXLSXFileIntoList()");          
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " : " );
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        }
      }
      workbook.Close();
      if (workSheetNotUsed.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ProcessInputXlsxFile::ReadXLSXFileIntoList()");
        foreach (var entry in workSheetNotUsed)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + entry + " rates are not being used.");
        }
      }
      if (discardedLines.Any())
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ProcessInputXlsxFile::ReadXLSXFileIntoList()");        
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Customer destinations discarded along with header lines.");
        discardedLines.Sort();
        foreach (var entry in discardedLines)
        {
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + entry);
        }
      }      
      StaticVariable.TwbHeader.Add(Environment.NewLine + "ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList()-- completed");
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "ReadXLSXFileIntoList()-- finished");
    }
    private static void MergeDefaultPricesListWithInputFileList()
    {
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "MergeDefaultPricesListWithInputFileList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "MergeDefaultPricesListWithInputFileList()");
      foreach (string tok in StaticVariable.DefaultEntries)
      {
        StaticVariable.InputXlsxFileDetails.Add(tok);
      }
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "MergeDefaultPricesListWithInputFileList()-- finished");
    }
    private static void AddToCustomerDetailsDataRecordList(List<string> lst)
    {
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "AddToPreRegExDataRecordList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "AddToPreRegExDataRecordList()");
      const string undefinedStandardInfo = "undefinedStandardBand\tundefinedStandardName\tundefinedStandardDestination\t";
      foreach (string token in lst)
      {
        try
        {                    
          StaticVariable.CustomerDetailsDataRecord.Add(new DataRecord(undefinedStandardInfo + token));
          
        }
        catch (Exception e)
        {
          StaticVariable.Errors.Add("ProcessInputXlsxFile::AddToPreRegExDataRecordList()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Problem adding list to DataRecord");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "AddToPreRegExDataRecordList()-- finished");
    }
    private static void MatchInputXlsxFileWithRegExAndAddToDestinationsMatchedByRegExDataRecord()
    {
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "RegExMatchInputList()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessInputXlsxFile".PadRight(30, '.') + "RegExMatchInputList()");            
      string destinationName = string.Empty;
      var tmpList = new List<string>();
      int destination = 0;
      int UniqueBandCounter = 100;
      Timer newTimer = new System.Timers.Timer(10000); // 2 sec interval
      string regExAlphanumeric = @"[0-9|a-z|A-Z]"; //@"\w|\s"
      string regExExtraSpaces = "\x0020{2,}";
      string regExNull = "\x0000";
      string regExNoise = @"\(|\)|( -|- )|\,|'";
      //string regEx_Noise = @"\,|'";
      Regex regExRemoveNull = new Regex(regExNull, RegexOptions.Multiline);
      Regex regExRemoveNoise = new Regex(regExNoise, RegexOptions.Compiled);
      Regex regExRemoveExtraSpaces = new Regex(regExExtraSpaces, RegexOptions.Compiled);
      Regex regExCheckForAlphanumeric = new Regex(regExAlphanumeric, RegexOptions.Compiled);

      foreach (string tok in StaticVariable.InputXlsxFileDetails)
      {
        var found = false;
        if (!tok.ToUpper().Contains("NAME") || !tok.StartsWith(";"))
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
            StaticVariable.Errors.Add("ParseInputFile::RegExMatchInputList()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There may be a problem with the input destination");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          foreach (string regexExpression in StaticVariable.CombinedRegex)
          {
            aryLine = regexExpression.Split('\t');
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
                //StaticVariable.Errors.Add("ParseInputFile::RegExMatchInputList()  -- Debugging only");
                //StaticVariable.Errors.Add(regExBand + "\t" + regexStandardName + "\t" + regExDescription + "\t" + tok);
              }
            }
            catch (Exception e)
            {
              StaticVariable.Errors.Add(Environment.NewLine + "ProcessInputXlsxFile::RegExMatchInputList()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There may be a problem with the regex");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Check if there are more than 1 international or domestic regex files");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
        }
        if (!found)
        {
          tmpList.Add("^" + destinationName.Replace(" ", ".").TrimEnd(' ') + "\tU" + UniqueBandCounter + "\t" + destinationName + "\t" + destinationName.PadRight(20, ' ').Substring(0, 20));
          UniqueBandCounter++;
        }
      }
      if (tmpList.Count > 0)
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ProcessInputXlsxFile::RegExMatchInputList()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Destinations without regex");        
        tmpList.Sort();
        foreach (string s in tmpList)
        {
          StaticVariable.Errors.Add(s);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      newTimer.Enabled = false;
      Console.WriteLine("ProcessInputXlsxFile".PadRight(30, '.') + "RegExMatchInputList()-- finished");
    }
    private static void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
    {
      Console.Write(".");
    }
    private static bool DiscardHeaderLine(string line)
    {      
      return line.ToUpper().StartsWith("DESTINATION") || line.ToUpper().Contains("NAME") && line.ToUpper().Contains("TABLE") && line.ToUpper().Contains("USING");      
    }
  }
}
