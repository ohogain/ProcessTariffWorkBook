// <copyright file="RearrangeCompletedFiles.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ProcessTariffWorkbook
{
  class RearrangeCompletedFiles
  {    
    public static void CreateCategoryMatrix()
  {
    Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCategoryMatrix() -- started");    
    StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCategoryMatrix() -- started");
    const string categoryMatrixHeader = "Source\tBand\tBand Type\tBand Description\tEnterprise Destination\tMobile Destination\tLandLine Destination\tEnterprise Call Type Filter\tMobile Top Premium Text Filter\tMobile Top Data Roaming Filter\tMobile Top Directory Enquiry Filter\tTop Directory Enquiry Filter\tLandline Top Directory Enquiry Filter\tEGL Subscription Report\tOperator Revenue\tBill Band Group Filter\tVoice Traffic Filter\tData Traffic Filter\tMessage Traffic Filter\tParentEnterpriseLeagueTables\tMobile Top Roaming Filter\tRoaming Calls Report\tFlat Rate\tRoaming Trips\tRoaming Costs Reports";
    StringBuilder sb = new StringBuilder();
    const string categoryMatrixTab = "CategoryMatrix";    
    int yAxis = 0;
    const string source = "BMS";
    string band = string.Empty;
    string bandType = "1";    
    string enterpriseDestination = string.Empty;
    string mobileDestination = string.Empty;
    string landLineDestination = string.Empty;
    string enterpriseCallTypeFilter = string.Empty;
    string mobileTopPremiumTextFilter = string.Empty;
    string mobileTopDataRoamingFilter = string.Empty;
    string mobileTopDirectoryEnquiryFilter = string.Empty;
    string topDirectoryEnquiryFilter = string.Empty;
    string landlineTopDirectoryEnquiryFilter = string.Empty;
    string eglSubscriptionReport = string.Empty;
    const string operatorRevenue = "usage";
    string billBandGroupFilter = string.Empty;
    string voiceTrafficFilter = string.Empty;
    string dataTrafficFilter = string.Empty;
    string messageTrafficFilter = string.Empty;
    string parentEnterpriseLeagueTables = string.Empty;
    string mobileTopRoamingFilter = string.Empty;
    string roamingCallsReport = string.Empty;
    string flatRate = string.Empty;
    string roamingTrips = string.Empty;
    string roamingCostsReports = string.Empty;
    List<string> categoryMatrix = new List<string>();      

    var categoryMatrixQuery =
      from dr in StaticVariable.CustomerDetailsDataRecord
      select new
      {
        dr.StdBand,
        dr.StdPrefixName,
        dr.StdPrefixDescription,
        dr.CustomerGroupBand,
        dr.CustomerGroupBandDescription,
        dr.CustomerUsingGroupBands,
        dr.CustomerUsingCustomerNames,
        dr.CustomerDestinationType,
        dr.CustomerPrefixName
      };

    categoryMatrix.Add(categoryMatrixHeader);
    foreach (var token in categoryMatrixQuery)
    {
      sb.Append(source + "\t");        
      band = token.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? token.CustomerGroupBand : token.StdBand;        
      sb.Append(band + "\t");
      sb.Append(bandType + "\t");
      string bandDescription = token.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? token.CustomerPrefixName : token.StdPrefixName;        
      sb.Append(bandDescription + "\t");
      enterpriseDestination = token.CustomerDestinationType;
      sb.Append(enterpriseDestination + "\t");
      sb.Append(mobileDestination + "\t");
      landLineDestination = token.CustomerDestinationType;
      sb.Append(landLineDestination + "\t");
      enterpriseCallTypeFilter = token.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? token.CustomerPrefixName : token.StdPrefixName;       
      sb.Append(enterpriseCallTypeFilter + "\t");
      sb.Append(mobileTopPremiumTextFilter + "\t");
      sb.Append(mobileTopDataRoamingFilter + "\t");
      sb.Append(mobileTopDirectoryEnquiryFilter + "\t");
      sb.Append(topDirectoryEnquiryFilter + "\t");
      sb.Append(landlineTopDirectoryEnquiryFilter + "\t");
      sb.Append(eglSubscriptionReport + "\t");
      sb.Append(operatorRevenue + "\t");
      sb.Append(billBandGroupFilter + "\t");
      sb.Append(voiceTrafficFilter + "\t");
      sb.Append(dataTrafficFilter + "\t");
      sb.Append(messageTrafficFilter + "\t");
      sb.Append(parentEnterpriseLeagueTables + "\t");
      sb.Append(mobileTopRoamingFilter + "\t");
      sb.Append(roamingCallsReport + "\t");
      sb.Append(flatRate + "\t");
      sb.Append(roamingTrips + "\t");
      sb.Append(roamingCostsReports);
      categoryMatrix.Add(sb.ToString());
      sb.Length = 0;
    }
    try
    {
      SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(StaticVariable.CategoryMatrixXlsxFile); 
      workbook.Worksheets["Sheet1"].Name = categoryMatrixTab; 
      foreach (string line in categoryMatrix)
      {
        string[] aryLine = line.Split('\t');
        for (int xAxis = 0; xAxis < aryLine.Length; xAxis++)
        {
          workbook.Worksheets[categoryMatrixTab].Cells.NumberFormat = "@";                   
          workbook.Worksheets[categoryMatrixTab].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
          workbook.Worksheets[categoryMatrixTab].Cells[yAxis, xAxis].Value = aryLine[xAxis];
        }
        yAxis++;
      }      
      workbook.Save();
      workbook.Close();
    }
    catch (Exception e)
    {
      StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CreateCategoryMatrix() -- started");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem writing to category matrix xlsx file");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
      ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
    }
    StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CreateCategoryMatrix() -- started");
    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The CategoryMatrix has been written to " + StaticVariable.FinalDirectory);
    Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCategoryMatrix() -- finished");
    StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCategoryMatrix() -- finished");
    }
    public static void WriteToV6TwbXlsxFile()
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV6TwbXlsxFile() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV6TwbXlsxFile() -- started");                        
      try
      { 
        SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(StaticVariable.V6TwbOutputXlsxFile);      
        WriteToBandsWorkSheet(workbook);
        WriteToTariffPlanSheet(workbook); 
        WriteToTableLinksSheet(workbook);
        WriteToPrefixBandsSheet(workbook);        
        WriteToPrefixNumbersSheet(workbook, MatchPrefixesWithDestinations());
        WriteToSourceDestinationBandsSheet(workbook);
        WriteToTimeSchemesSheet(workbook);
        WriteToTimeSchemesExceptionsSheet(workbook);        
        workbook.Save();
        workbook.Close();
      }
      catch (Exception e)
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::WriteToV6TwbXlsxFile() -- started");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem writing to output xlsx file");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV6TwbXlsxFile() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV6TwbXlsxFile() -- finish");
    }
    private static void WriteToBandsWorkSheet(SpreadsheetGear.IWorkbook workbook)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToBandsWorkSheet() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToBandsWorkSheet() -- started");
      const string bandsHeader = "Band Name\tDescription\tRate1\tRate1 Initial\tRate1 Subseq\tRate2\tRate2 Initial\tRate2 Subseq" +
                                 "\tRate3\tRate3 Initial\tRate3 Subseq\tRate4\tRate4 Initial\tRate4 Subseq\tMinimum Call Cost" +
                                 "\tConnection Charge" + "\tWhole Interval Charging\tTime Scheme Name\tInitial Interval Length" +
                                 "\tSubsequent Intervals Length\tMinimum Intervals" + "\tIntervals At Initial Cost\tMinimum Duration" +
                                 "\tIs Multi-Level\tCutOff1 Cost\tCutOff2 Duration\tV5 Destination Type"/*\tCharging Type"*/; // to be added if V6 (TDI) uses capped or pulse rates.
      const string sheetName = "Bands";
      int yAxis = 0;      
      List<string> uniqueEntries = new List<string>();
      workbook.Worksheets["Sheet1"].Name = sheetName;         
      string[] listLine = bandsHeader.Split('\t');
      for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
      {
        workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
        workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
        workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
      }
      yAxis++;
      #region Duration
      var durationQuery =
        from dq in StaticVariable.CustomerDetailsDataRecord
        where dq.ChargingType.ToUpper().Equals("DURATION")
        select dq;
      if (durationQuery.Any())
      {
        foreach (DataRecord d in durationQuery)
        {
          StringBuilder sb = new StringBuilder();         
          if (d.CustomerUsingGroupBands.ToUpper().Equals("TRUE"))
          {
            sb.Append(d.CustomerGroupBand.ToUpper() + "\t" + d.CustomerGroupBandDescription + "\t");
          }
          else
          {
            sb.Append(d.StdBand.ToUpper() + "\t");
            sb.Append(d.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? d.CustomerPrefixName + "\t" : d.StdPrefixName + "\t");
          }          
          sb.Append(StaticVariable.Rate1Name + "\t" + d.CustomerFirstInitialRate + "\t" + d.CustomerFirstSubseqRate + "\t");
          sb.Append(StaticVariable.Rate2Name + "\t" + d.CustomerSecondInitialRate + "\t" + d.CustomerSecondSubseqRate + "\t");
          sb.Append(StaticVariable.Rate3Name + "\t" + d.CustomerThirdInitialRate + "\t" + d.CustomerThirdSubseqRate + "\t");
          sb.Append(StaticVariable.Rate4Name + "\t" + d.CustomerFourthInitialRate + "\t" + d.CustomerFourthSubseqRate + "\t");
          sb.Append(d.CustomerMinCharge + "\t");
          sb.Append(d.CustomerConnectionCost + "\t");
          sb.Append(ValidateData.AdjustRoundingValueForV6Twb(d.CustomerRounding) + "\t");
          sb.Append(ValidateData.CapitaliseWord(d.CustomerTimeScheme) + "\t");
          sb.Append(d.CustomerInitialIntervalLength + "\t");
          sb.Append(d.CustomerSubsequentIntervalLength + "\t");
          sb.Append(d.CustomerMinimumIntervals + "\t");
          sb.Append(d.CustomerIntervalsAtInitialCost + "\t");
          sb.Append(d.CustomerMinimumTime + "\t");
          sb.Append(d.CustomerMultiLevelEnabled + "\t");
          sb.Append(d.CustomerCutOff1Cost + "\t");
          sb.Append(d.CustomerCutOff2Duration + "\t");
          sb.Append(ValidateData.CapitaliseWord(d.CustomerDestinationType)/* + "\t" + d.ChargingType*/); // d.ChargingType to be added if V6 (TDI) uses capped or pulse rates.          

          uniqueEntries.Add(sb.ToString());
        }
        try
        {
          uniqueEntries = uniqueEntries.Distinct().ToList();
          foreach (var lines in uniqueEntries)
          {
            listLine = lines.Split('\t');
            for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
            {
              workbook.Worksheets[sheetName].Cells.NumberFormat = "@";
              workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left;
              workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
            }
            yAxis++;
          }            
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::WriteToBandsWorkSheet() -- started");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem with writing duration Bands into Excel sheet");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }        
      }
      #endregion
      #region Capped 
      var cappedQuery =
        from dq in StaticVariable.CustomerDetailsDataRecord
        where dq.ChargingType.ToUpper().Equals("CAPPED")
        select dq;
      if (cappedQuery.Any())
      {
        foreach (DataRecord d in cappedQuery)
        {
          StringBuilder sb = new StringBuilder();
          sb.Append(d.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? d.CustomerGroupBand.ToUpper() + "\t" : d.StdBand.ToUpper() + "\t");
          sb.Append(d.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? d.CustomerPrefixName + "\t" : d.StdPrefixName + "\t");
          sb.Append(Constants.Rate1 + "\tNULL\tNULL\t");
          sb.Append(Constants.Rate2 + "\t" + d.CustomerFirstInitialRate + "\t" + d.CustomerFirstSubseqRate + "\t");
          sb.Append(Constants.Rate3 + "\t" + d.CustomerSecondInitialRate + "\tNULL\t");
          sb.Append(Constants.Rate4 + "\tNULL\tNULL\t");
          sb.Append(d.CustomerMinCharge + "\t");
          sb.Append(d.CustomerConnectionCost + "\t");
          sb.Append(ValidateData.AdjustRoundingValueForV6Twb(d.CustomerRounding) + "\t");
          sb.Append(ValidateData.CapitaliseWord(d.CustomerTimeScheme) + "\t");
          sb.Append(d.CustomerInitialIntervalLength + "\t");
          sb.Append(d.CustomerSubsequentIntervalLength + "\t");
          sb.Append(d.CustomerMinimumIntervals + "\t");
          sb.Append(d.CustomerIntervalsAtInitialCost + "\t");
          sb.Append(d.CustomerMinimumTime + "\t");
          sb.Append(d.CustomerMultiLevelEnabled + "\t");
          sb.Append(d.CustomerCutOff1Cost + "\t");
          sb.Append(d.CustomerCutOff2Duration + "\t");
          sb.Append(ValidateData.CapitaliseWord(d.CustomerDestinationType) + "\t" + d.ChargingType); // d.ChargingType to be added if V6 (TDI) uses capped or pulse rates.          

          String cappedline = sb.ToString();
          try
          {
            listLine = cappedline.Split('\t');
            for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
            {
              workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
              workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
              workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
            }
            yAxis++;
          }
          catch (Exception e)
          {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::WriteToBandsWorkSheet() -- started");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem with writing capped Bands into Excel sheet");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
      }
      #endregion 
      #region Pulse

      const string wholeIntervalChargingNotUsedInPulse = "NO";
      const string intervalLengthsNotUsedInPulse = "60";
      var pulseQuery =
        from dq in StaticVariable.CustomerDetailsDataRecord
        where dq.ChargingType.ToUpper().Equals("PULSE")
        select dq;
      if (pulseQuery.Any())
      {
        foreach (DataRecord d in pulseQuery)
        {
          StringBuilder sb = new StringBuilder();
          sb.Append(d.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? d.CustomerGroupBand.ToUpper() + "\t" : d.StdBand.ToUpper() + "\t");
          sb.Append(d.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? d.CustomerPrefixName + "\t" : d.StdPrefixName + "\t");
          sb.Append(Constants.Rate1 + "\tNULL\tNULL\t");
          sb.Append(Constants.Rate2 + "\t" + d.CustomerFirstInitialRate + "\t" + d.CustomerFirstSubseqRate + "\t");
          sb.Append(Constants.Rate3 + "\tNULL\tNULL\t");
          sb.Append(Constants.Rate4 + "\tNULL\tNULL\t");
          sb.Append(d.CustomerMinCharge + "\t");
          sb.Append(d.CustomerConnectionCost + "\t");
          sb.Append(wholeIntervalChargingNotUsedInPulse + "\t");
          sb.Append(ValidateData.CapitaliseWord(d.CustomerTimeScheme) + "\t");
          sb.Append(intervalLengthsNotUsedInPulse + "\t");
          sb.Append(intervalLengthsNotUsedInPulse + "\t");
          sb.Append(d.CustomerMinimumIntervals + "\t");
          sb.Append(d.CustomerIntervalsAtInitialCost + "\t");
          sb.Append(d.CustomerMinimumTime + "\t");
          sb.Append(d.CustomerMultiLevelEnabled + "\t");
          sb.Append(d.CustomerCutOff1Cost + "\t");
          sb.Append(d.CustomerCutOff2Duration + "\t");
          sb.Append(ValidateData.CapitaliseWord(d.CustomerDestinationType)/* + "\t" + d.ChargingType*/); // d.ChargingType to be added if V6 (TDI) uses capped or pulse rates.          

          string pulseLine = sb.ToString();

          try
          {
            listLine = pulseLine.Split('\t');
            for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
            {
              workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
              workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
              workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
            }
            yAxis++;
          }
          catch (Exception e)
          {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::WriteToBandsWorkSheet() -- started");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem with writing pulse Bands into Excel sheet");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
      }
      #endregion       // column allocation incorect
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToBandsWorkSheet() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToBandsWorkSheet() -- finish");
    }
    private static void WriteToTariffPlanSheet(SpreadsheetGear.IWorkbook workbook) 
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTariffPlanSheet() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTariffPlanSheet() -- started");
      int yAxis = 0;      
      string line = string.Empty;
      const string sheetName = "TariffPlan";  
          
      workbook.Worksheets.Add();
      workbook.Worksheets["Sheet2"].Name = sheetName;     
      foreach (string token in StaticVariable.TariffPlan)
      {
        string[] tariffPlanLine = token.Split('=');
        if (tariffPlanLine[0].ToUpper().Contains(Constants.TariffPlanName))
        {
          line = tariffPlanLine[0] + "\t" + tariffPlanLine[1] + " " + StaticVariable.VersionValue; 
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::WriteToTariffPlanSheet()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "'Tariff Plan Name' has the version number added to it - '" + tariffPlanLine[1] + " " + StaticVariable.VersionValue + "'");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The version number can be removed when V6 (TDI) uses the tariff plan naming convention");
        }
        else if (tariffPlanLine[0].ToUpper().Contains("RATE"))
        {
          line = "Rate\t" + tariffPlanLine[1];  //Names must be Rate, not rate1, rate2, etc
        }             
        else
        {
          line = tariffPlanLine[0] + "\t" + tariffPlanLine[1];          
        }        
        if (tariffPlanLine[0].ToUpper().Contains(Constants.Holiday))
        {
          List<string> allHolidays = DisplayHolidays(tariffPlanLine);
          foreach (var hol in allHolidays)
          {
            string[] holidayLine = hol.Split('\t');
            for (int xAxis = 0; xAxis < holidayLine.Length; xAxis++)
            {
              workbook.Worksheets[sheetName].Cells.NumberFormat = "@";
              workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left;
              workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = holidayLine[xAxis];
            }
            yAxis++;
          }
        }        
        else
        {
          if (line.ToUpper().Contains("N/A") && ValidateData.CheckForPulseWorksheet()) 
          {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::WriteToTariffPlanSheet");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There is no entry for 'Carrier unit price' in the TariffPlan header even though there is a pulse worksheet.");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If there is no pulse rates, delete the pulse worksheet.");
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
          else if (line.ToUpper().Contains(Constants.CarrierUnitPrice) && ValidateData.CheckForPulseWorksheet())
          {
            string[] listLine = line.Split('\t');
            for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
            {
              workbook.Worksheets[sheetName].Cells.NumberFormat = "@";
              workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left;
              workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
            }
            yAxis++;
          }
          else if (line.ToUpper().Contains(Constants.CarrierUnitPrice)) continue;
          {
            string[] listLine = line.Split('\t');
            for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
            {
              workbook.Worksheets[sheetName].Cells.NumberFormat = "@";
              workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left;
              workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
            }
            yAxis++;
          }
        }                               
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTariffPlanSheet() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTariffPlanSheet() -- finish");
    }
    private static void WriteToTableLinksSheet(SpreadsheetGear.IWorkbook workbook)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTableLinksSheet() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTableLinksSheet() -- started");
      int yAxis = 0;      
      const string tableLinksHeader = "Table Name\tPrefix\tPass Prefix\tDestination";
      const string sheetName = "TableLinks";
      workbook.Worksheets.Add();
      workbook.Worksheets["Sheet3"].Name = sheetName;
      string[] listLine = tableLinksHeader.Split('\t');
      for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
      {
        workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
        workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
        workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
      }
      yAxis++;
      foreach (string token in StaticVariable.TableLinks)
      {
        listLine = token.Split('\t');
        for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
        {
          workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                    
          workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
          workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
        }
        yAxis++;
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTableLinksSheet() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTableLinksSheet() -- finish");
    }
    private static void WriteToPrefixBandsSheet(SpreadsheetGear.IWorkbook workbook)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixBandsSheet() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixBandsSheet() -- started");
      int yAxis = 0;
      string pBDestination = string.Empty;
      string pBBand = string.Empty;
      const string prefixBandsHeader = "Table Name\tDestination\tBand";
      const string sheetName = "PrefixBands";
      workbook.Worksheets.Add();
      workbook.Worksheets["Sheet4"].Name = sheetName;
      string[] listLine = prefixBandsHeader.Split('\t');
      for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
      {
        workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
        workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
        workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
      }
      yAxis++;
      foreach (DataRecord drm in StaticVariable.CustomerDetailsDataRecord)
      {
        pBBand = drm.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? drm.CustomerGroupBand : drm.StdBand;
        pBDestination = drm.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? drm.CustomerPrefixName : drm.StdPrefixName;
        string line = ValidateData.CapitaliseWord(drm.CustomerTableName) + "\t" + pBDestination + "\t" + pBBand;
        listLine = line.Split('\t');
        for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
        {
          workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                    
          workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
          workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
        }
        yAxis++;
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixBandsSheet() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixBandsSheet() -- finish");
    }
    private static void WriteToPrefixNumbersSheet(SpreadsheetGear.IWorkbook workbook, List<string> prefixList)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixNumbersSheet2() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixNumbersSheet2() -- started");
      int yAxis = 0;      
      const string prefixNumbersHeader = "Table Name\tPrefix Number\tPrefix Name";
      const string sheetName = "PrefixNumbers";
      workbook.Worksheets.Add();
      workbook.Worksheets["Sheet5"].Name = sheetName;
      string[] listLine = prefixNumbersHeader.Split('\t');

      for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
      {
        workbook.Worksheets[sheetName].Cells.NumberFormat = "@";
        workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left;
        workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
      }
      yAxis++;

      foreach (var prefix in prefixList)
      {
        listLine = ValidateData.CapitaliseWord(prefix).Split('\t');
        for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
        {
          workbook.Worksheets[sheetName].Cells.NumberFormat = "@";
          workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left;
          workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
        }
        yAxis++;
      }
      List<string> nw = GetNationalDomesticPrefixes();
      foreach (var column in nw)
      {
        listLine = column.Split('\t');// do not capitalise
        for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
        {
          workbook.Worksheets[sheetName].Cells.NumberFormat = "@";
          workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left;
          workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
        }
        yAxis++;
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixNumbersSheet2() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToPrefixNumbersSheet2() -- finished");
    }    
    private static void WriteToSourceDestinationBandsSheet(SpreadsheetGear.IWorkbook workbook)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToSourceDestinationBandsSheet() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToSourceDestinationBandsSheet() -- started");
      int yAxis = 0;
      const string sourceDestinationBandsHeader = "Table Name\tSource\tDestination\tBand";
      const string sheetName = "SourceDestinationBands";
      workbook.Worksheets.Add();
      workbook.Worksheets["Sheet6"].Name = sheetName;
      string[] listLine = sourceDestinationBandsHeader.Split('\t');
      for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
      {
        workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
        workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
        workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
      }
      yAxis++;
      foreach (var source in StaticVariable.SourceDestinationBands)
      {
        listLine = source.Split('\t');
        for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
        {
          workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                    
          workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
          workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
        }
        yAxis++;
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToSourceDestinationBandsSheet() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToSourceDestinationBandsSheet() -- finish");
    }
    private static void WriteToTimeSchemesSheet(SpreadsheetGear.IWorkbook workbook)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesSheet() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesSheet() -- started");
      int yAxis = 0;
      const string timeSchemesHeader = "Time Scheme Name\tHolidays Relevant\tDefault Rate";
      const string sheetName = "TimeSchemes";
      workbook.Worksheets.Add();
      workbook.Worksheets["Sheet7"].Name = sheetName;
      string[] listLine = timeSchemesHeader.Split('\t');
      for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
      {
        workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
        workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
        workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
      }
      yAxis++;
      foreach (var scheme in StaticVariable.TimeSchemes)
      {
        listLine = scheme.Split('\t');
        for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
        {
          workbook.Worksheets[sheetName].Cells.NumberFormat = "@"; 
                                                                   //workbook.Worksheets[sheetName].Cells.VerticalAlignment = SpreadsheetGear.VAlign.Center; // align center of cell
          workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
          workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
        }
        yAxis++;
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesSheet() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesSheet() -- finish");
    }
    private static void WriteToTimeSchemesExceptionsSheet(SpreadsheetGear.IWorkbook workbook)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesExceptionsSheet() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesExceptionsSheet() -- started");
      int yAxis = 0;
      const string timeSchemeExceptionsHeader = "Time Scheme Name\tDay\tStart\tFinish\tRate";
      const string sheetName = "TimeSchemeExceptions";
      workbook.Worksheets.Add();
      workbook.Worksheets["Sheet8"].Name = sheetName;
      string[] listLine = timeSchemeExceptionsHeader.Split('\t');
      for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
      {
        workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                      
        workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
        workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
      }
      yAxis++;
      foreach (var source in StaticVariable.TimeSchemesExceptions)
      {
        listLine = source.Split('\t');
        for (int xAxis = 0; xAxis < listLine.Length; xAxis++)
        {
          workbook.Worksheets[sheetName].Cells.NumberFormat = "@";                    
          workbook.Worksheets[sheetName].Cells.HorizontalAlignment = SpreadsheetGear.HAlign.Left; 
          workbook.Worksheets[sheetName].Cells[yAxis, xAxis].Value = listLine[xAxis];
        }
        yAxis++;
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesExceptionsSheet() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToTimeSchemesExceptionsSheet() -- finish");
    }    
    public static void CopyOutputXlsxFileToV6OpUtilFolder(bool moveXlsxFileToOpUtilFolder)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CopySheetsToDropFolder() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CopySheetsToDropFolder() -- started");                            
      const string v6TwbDropFolder = @"\op-utils\WorkingFolder\TariffTools";
      if (moveXlsxFileToOpUtilFolder)
      {        
        string dropFolderFile = @"\" + v6TwbDropFolder + @"\" + Path.GetFileName(StaticVariable.V6TwbOutputXlsxFile);
        try
        {          
          if (File.Exists(dropFolderFile))
          {
            File.Delete(dropFolderFile);
          }
        }
        catch (Exception de)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CopySheetsToDropFolder() -- started");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem deleting sheet files in TWB drop folder");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TWB Drop Folder is: " + v6TwbDropFolder);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + de.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        try
        {
          File.Copy(StaticVariable.V6TwbOutputXlsxFile, dropFolderFile);
        }
        catch (PathTooLongException ptl)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CopySheetsToDropFolder() -- started");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Path too long. It must not exceed 248 chars");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TWB Drop Folder is: " + v6TwbDropFolder);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + ptl.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        catch (ArgumentException ae)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CopySheetsToDropFolder() -- started");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "sourceFileName or destFileName is a zero-length string, has invalid chars or only white space");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TWB Drop Folder is: " + v6TwbDropFolder);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + ae.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        catch (NotSupportedException nse)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CopySheetsToDropFolder() -- started");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "sourceFileName or destFileName is in an invalid format");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TWB Drop Folder is: " + v6TwbDropFolder);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + nse.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        catch (FileNotFoundException fnf)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CopySheetsToDropFolder() -- started");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Source file not found");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TWB Drop Folder is: " + v6TwbDropFolder);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + fnf.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "RearrangeCompletedFiles::CopySheetsToDropFolder() -- started");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem moving sheet files to TWB drop folder");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "TWB Drop Folder is: " + v6TwbDropFolder);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CopySheetsToDropFolder() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CopySheetsToDropFolder() -- finished");
    }
    public static void WriteOutV5Tc2Files()
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteOutV5Tc2Files() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteOutV5Tc2Files() -- started");
      CreateV5Tc2PricesFile();

      // matrix must come after the duration, pulse & capped rates - change this.

      WritePrefixIniFiles(MatchPrefixesWithDestinations());
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteOutV5Tc2Files() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteOutV5Tc2Files() -- finished");
    }
    private static void CreateV5Tc2PricesFile()
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateV5Tc2PricesFile() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateV5Tc2PricesFile() -- started");
      List<string> v5PricesIni = new List<string>();
      MakeGeneralHeader(v5PricesIni);
      AddDurationPrices(v5PricesIni);
      AddCappedPrices(v5PricesIni);
      AddPulsePrices(v5PricesIni);
      GetDurationMatrix(v5PricesIni);
      GetCappedMatrix(v5PricesIni);
      GetPulseMatrix(v5PricesIni);
      WriteToV5Tc2PricesFile(v5PricesIni);
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateV5Tc2PricesFile() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateV5Tc2PricesFile() -- finished");
    }
    private static void MakeGeneralHeader(List<string> pricesIni )
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "MakeGeneralHeader() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "MakeGeneralHeader() -- started"); 
      pricesIni.Add("[GENERAL]");
      pricesIni.Add("Carrier=" + GetTariffPlanValues("OPERATOR NAME"));
      pricesIni.Add("Release Note=New Prices");
      pricesIni.Add("Release Date=" + GetTariffPlanValues("RELEASE DATE"));
      pricesIni.Add("Customer Reference=0");
      pricesIni.Add("Country=" + GetTariffPlanValues("COUNTRY"));
      pricesIni.Add("Country Code=" + GetTariffPlanValues("COUNTRY CODE"));
      pricesIni.Add("Carrier Unit Price=" + GetTariffPlanValues("CARRIER UNIT PRICE"));
      pricesIni.Add(Environment.NewLine + "[Rates]");
      pricesIni.Add(GetTariffPlanValues("RATE1"));
      pricesIni.Add(GetTariffPlanValues("RATE2"));
      pricesIni.Add(GetTariffPlanValues("RATE3"));
      pricesIni.Add(GetTariffPlanValues("RATE4"));
      pricesIni.Add(Environment.NewLine + "[Destination Types]");
      pricesIni.Add("Local");
      pricesIni.Add("National");
      pricesIni.Add("International");
      pricesIni.Add("International Mobile");
      pricesIni.Add("Mobile");
      pricesIni.Add("Services");
      pricesIni.Add("Other");          
      FillTimeSchemes(pricesIni);
      pricesIni.Add(Environment.NewLine + "[HOLIDAYS]");      
      FillHolidays(pricesIni);
      pricesIni.Add(Environment.NewLine + "[LINK]");
      pricesIni.Add("Start Table=" + GetTariffPlanValues("STARTING POINT TABLE NAME"));
      FillTableLinksValues(StaticVariable.TableLinks, pricesIni);
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "MakeGeneralHeader() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "MakeGeneralHeader() -- finished");
    }    
    private static string GetTariffPlanValues(string word)
    {
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetTariffPlanValues() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetTariffPlanValues() -- started");
      string result = string.Empty;
      const int key = 0;
      const int value = 1;
      foreach (var variable in StaticVariable.TariffPlan)
      {
        string[] values = variable.Split('=');
        if (values[key].ToUpper().Equals(word.ToUpper()))
        {
          result = values[value];
          break;
        }
      }
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetTariffPlanValues() -- finished");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetTariffPlanValues() -- finished");
      return result;
    }
    private static void FillTableLinksValues(List<string> tableLinks, List<string> defaultheader )
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "FillTableLinksValues() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "FillTableLinksValues() -- started");
      const int startTable = 0;
      const int prefix = 1;
      const int passPrefix = 2;
      const int destinationTable = 3;
      const int tableName = 1;
      foreach (var variable in tableLinks)
      {
        string[] line = variable.Split('\t');
        string[] tableMinusCountryCode = line[startTable].Split('_');
        defaultheader.Add(tableMinusCountryCode[tableName] + ", " + line[destinationTable].Substring(line[destinationTable].IndexOf("_", StringComparison.Ordinal) + 1) + ", " + line[prefix] + ", " + line[destinationTable] + ", " + ConvertPassPrefixValue(line[passPrefix]));       
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "FillTableLinksValues() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "FillTableLinksValues() -- finished");
    }
    private static string ConvertPassPrefixValue(string passPrefixValue)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "ConvertPassPrefixValue() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "ConvertPassPrefixValue() -- started");
      string result = string.Empty;
      if (passPrefixValue.ToUpper().Equals("YES")) result = "TRUE";      
      if (passPrefixValue.ToUpper().Equals("NO")) result = "FALSE";      
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "ConvertPassPrefixValue() -- finished");
      return result;
    }
    private static void FillHolidays(List<string> pricesini)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "FillHolidays() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "FillHolidays() -- started");
      string hols = GetTariffPlanValues("HOLIDAY");
      string[] holidays = hols.Split(',');
      foreach (var var in holidays)
      {
        pricesini.Add(var.Trim());
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "FillHolidays() -- finished");
    }  
    private static void FillTimeSchemes(List<string> pricesini )
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "FillTimeSchemes() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "FillTimeSchemes() -- started");
      const int timeSchemeName = 0;      
      const int defaultRate = 2;
      const int day = 1;
      const int startTime = 2;
      const int finishtime = 3;
      const int rate = 4;

      foreach (var var in StaticVariable.TimeSchemes)
      {
        var timeSchemeDetails = var.Split('\t');
        pricesini.Add(Environment.NewLine + "[Time Scheme]");
        pricesini.Add("Scheme Name=" + timeSchemeDetails[timeSchemeName]);
        pricesini.Add("Default Rate=" + timeSchemeDetails[defaultRate]);        
        foreach (var times in StaticVariable.TimeSchemesExceptions)
        {
          var timeSchemeExceptionsDetails = times.Split('\t');          
          if (timeSchemeExceptionsDetails[timeSchemeName].ToUpper().Equals(timeSchemeDetails[timeSchemeName].ToUpper()))
          {
            string holidaySpelling = (timeSchemeExceptionsDetails[day].ToUpper().Equals("HOL")) ? "Holiday" : timeSchemeExceptionsDetails[day];
            pricesini.Add(timeSchemeExceptionsDetails[rate] + "," + holidaySpelling + "," + timeSchemeExceptionsDetails[startTime] + "-" + timeSchemeExceptionsDetails[finishtime]);
          }
        }
      }      
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "FillTimeSchemes() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "FillTimeSchemes() -- finished");
    }    
    private static string AdjustRoundingValueForPricesIni(string value)
    {
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AdjustRoundingValueForPricesIni() -- started");
      string result = string.Empty;

      if (value.ToLower().Equals("yes") || value.ToLower().Equals("y") || value.ToLower().Equals("1") || value.ToLower().Equals("roundup") || value.ToLower().Equals("round up"))
      {
        result = "1";
      }
      else if (value.ToLower().Equals("no") || value.ToLower().Equals("n") || value.ToLower().Equals("3") || value.ToLower().Equals("exact") || value.ToLower().Equals("noround") || value.ToLower().Equals("no round"))
      {
        result = "3";
      }
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AdjustRoundingValueForPricesIni() -- finished");
      return result;
    }
    private static void WriteToV5Tc2PricesFile(List<string> pricesini)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV5Tc2PricesFile() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV5Tc2PricesFile() -- started");
      string pricesFile = StaticVariable.V6TwbOutputXlsxFile.Replace(".xlsx", ".ini");
      string outputDirectory = Path.GetDirectoryName(pricesFile);
      string outputFile = Path.GetFileName(pricesFile);
      string[] nametokens = outputFile.Split('_');
      pricesFile = outputDirectory + @"\" + nametokens[2] + " " + nametokens[3] + " " + nametokens[4] + " " + nametokens[5];

      if (File.Exists(pricesFile))
      {
        File.Delete(pricesFile);
      }
      using (StreamWriter oSw = new StreamWriter(File.OpenWrite(pricesFile), Encoding.Unicode))
      {
        foreach (var variable in pricesini)
        {
          oSw.WriteLine(variable);
        }
        oSw.Close();
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV5Tc2PricesFile() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WriteToV5Tc2PricesFile() -- finished");
    } 
    private static void WritePrefixIniFiles(List<string> prefixList)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WritePrefixIniFiles() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WritePrefixIniFiles() -- started");
      const int tableName = 0;
      const int prefix = 1;
      const int prefixName = 2;

      var queryGetAllPrefixtables =
        (from db in StaticVariable.PrefixNumbersRecord
         select db.TableName).Distinct();      

      foreach (var currentPrefixTable in queryGetAllPrefixtables)
      {        
        string prefixFile = StaticVariable.FinalDirectory + @"\" + currentPrefixTable + ".ini";
        if (File.Exists(prefixFile))
        {
          File.Delete(prefixFile);
        }

        using (StreamWriter oSw = new StreamWriter(File.OpenWrite(prefixFile), Encoding.Unicode))
        {
          List<string> currentPrefixes = new List<string>();
          oSw.WriteLine("[New Prefix]");
          oSw.WriteLine("Table Name=" + currentPrefixTable + Environment.NewLine);          
          foreach (var prefixDetail in prefixList)
          {
            string[] nameAndPrefix = prefixDetail.Split('\t');
            if (nameAndPrefix[tableName].ToUpper().Equals(currentPrefixTable.ToUpper()))
            {
              currentPrefixes.Add(nameAndPrefix[prefixName] + "," + nameAndPrefix[prefix]);
            }            
          }

          if (StaticVariable.NationalTableSpellingValue.ToUpper().Equals(currentPrefixTable.ToUpper()))
          {
            List<string> nationalPrefixes = GetNationalDomesticPrefixes();          
            foreach (var np in nationalPrefixes)
            {
              string[] arr = np.Split('\t');
              currentPrefixes.Add(arr[prefixName] + "," + arr[prefix]);          
            }                        
          }          
          currentPrefixes = currentPrefixes.Distinct().ToList();
          currentPrefixes.Sort();
          foreach (var entry in currentPrefixes)
          {
            oSw.WriteLine(entry);
          }
          oSw.Close();          
        }
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "WritePrefixIniFiles() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "WritePrefixIniFiles() -- finished");
    }
    private static List<string> MatchPrefixesWithDestinations()
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- started");

        var query =
        from drm in StaticVariable.CustomerDetailsDataRecord
        join pn in StaticVariable.PrefixNumbersRecord on drm.StdPrefixName.ToUpper() equals pn.stdPrefixName.ToUpper()        
        select new { pn.TableName, pn.PrefixName, pn.PrefixNumber, drm.CustomerPrefixName, drm.CustomerUsingCustomerNames, pn.stdPrefixName };

        List<string> prefixesMatched = (from entry in query let prefixName = entry.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? entry.CustomerPrefixName : entry.stdPrefixName select entry.TableName + "\t" + entry.PrefixNumber + "\t" + prefixName).ToList();
        prefixesMatched = prefixesMatched.Distinct().ToList();

      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- finished");      
      return prefixesMatched;
    }
    private static List<string> GetNationalDomesticPrefixes()
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetNationalPrefixes() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetNationalPrefixes() -- started");
      var nationalTableResult =
        from pndr in StaticVariable.PrefixNumbersRecord
        where pndr.TableName.ToUpper().Equals(StaticVariable.NationalTableSpellingValue.ToUpper()) && !pndr.PrefixNumber.ToUpper().Equals("?") //exclude national,?     
        select pndr;

      var nationalPrefixes = nationalTableResult.Select(column => column.TableName + "\t" + column.PrefixNumber + "\t" + column.PrefixName).ToList();
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetNationalPrefixes() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetNationalPrefixes() -- finished");
      return nationalPrefixes;
    }
    private static void AddDurationPrices(List<string> pricesini)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "AddDurationPrices() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AddDurationPrices() -- started");
      string previousBand = string.Empty;           

      var queryDuration =
        from db in StaticVariable.CustomerDetailsDataRecord
        where db.ChargingType.ToUpper().Equals("DURATION") 
        orderby db.CustomerUsingGroupBands, db.CustomerGroupBand, db.StdPrefixName
        select db;      
        
      foreach (var details in queryDuration)
      {
        string currentBand = details.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? details.CustomerGroupBand : details.StdBand;        
        string name = details.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? details.CustomerPrefixName : details.StdPrefixName;
        name = name.ToUpper().Equals(currentBand.ToUpper()) ? name.ToUpper() : ValidateData.CapitaliseWord(name);
        
        if (!currentBand.ToUpper().Equals(previousBand.ToUpper()))        
        {
          CreateDurationHeader(pricesini, details, "Duration Rate");                    
        }
        pricesini.Add("(" + currentBand.ToUpper() + ")" + name + "," + GetDurationPrices(details));
        previousBand = currentBand;        
      }      
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "AddDurationPrices() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AddDurationPrices() -- finished");
    }
    private static void AddCappedPrices(List<string> pricesini)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "AddCappedPrices() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AddCappedPrices() -- started");
      string previousBand = string.Empty;            

      var queryCapped =
        from db in StaticVariable.CustomerDetailsDataRecord
        where db.ChargingType.ToUpper().Equals("CAPPED")
        orderby db.CustomerGroupBand
        select db;
      
      foreach (var details in queryCapped)
      {
        string currentBand = details.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? details.CustomerGroupBand : details.StdBand;        
        string name = details.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? details.CustomerPrefixName : details.StdPrefixName;
        name = name.ToUpper().Equals(currentBand.ToUpper()) ? name.ToUpper() : ValidateData.CapitaliseWord(name);

        if (!currentBand.ToUpper().Equals(previousBand.ToUpper()))
        {
          CreateCappedHeader(pricesini, details, "CAPPED");
        }
        pricesini.Add("(" + currentBand.ToUpper() + ")" + name + "," + GetCappedPrices(details));
        previousBand = currentBand;
      }      
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "AddCappedPrices() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AddCappedPrices() -- finished");
    }
    private static void AddPulsePrices(List<string> pricesini)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "AddPulsePrices() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AddPulsePrices() -- started");            
      string previousBand = string.Empty;            

      var queryPulse =
        from db in StaticVariable.CustomerDetailsDataRecord
        where db.ChargingType.ToUpper().Equals("PULSE")
        select db;

      foreach (var details in queryPulse)
      {
        string currentBand = details.CustomerUsingGroupBands.ToUpper().Equals("TRUE") ? details.CustomerGroupBand : details.StdBand;        
        string name = details.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? details.CustomerPrefixName : details.StdPrefixName;
        name = name.ToUpper().Equals(currentBand.ToUpper()) ? name.ToUpper() : ValidateData.CapitaliseWord(name);

        if (!currentBand.ToUpper().Equals(previousBand.ToUpper()))
        {
          CreatePulseHeader(pricesini, details, "PULSE");
        }
        pricesini.Add("(" + currentBand.ToUpper() + ")" + name + "," + GetPulsePrices(details));
        previousBand = currentBand;
      }      
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "AddPulsePrices() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "AddPulsePrices() -- finished");
    }   
    private static void GetDurationMatrix(List<string> priceIni)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationMatrix() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationMatrix() -- started");
      const int sourceColumn = 1;
      const int destinationColumn = 2;
      const int bandColumn = 3;
      const string duration= "DURATION";
      HashSet<string> uniqueBands = new HashSet<string>();      

      foreach (var band in StaticVariable.SourceDestinationBands)
      {
        string[] sourceBands = band.Split('\t');        
        uniqueBands.Add(sourceBands[bandColumn]);
      }

      foreach (var item in uniqueBands)
      {
        var query =
          from db in StaticVariable.CustomerDetailsDataRecord
          where db.CustomerGroupBand.ToUpper().Equals(item.ToUpper()) || db.StdBand.ToUpper().Equals(item.ToUpper()) && db.ChargingType.ToUpper().Equals(duration)
          select db;

        foreach (var result in query)
        {
          CreateDurationHeader(priceIni, result, "MATRIX DURATION");
          foreach (var matrix in StaticVariable.SourceDestinationBands)
          {            
            string[] sourceBands = matrix.Split('\t');
            if (sourceBands[bandColumn].ToUpper().Equals(item.ToUpper()))
            {              
              priceIni.Add("(" + item + ")" + sourceBands[sourceColumn] + "," + sourceBands[destinationColumn] + "," + GetDurationPrices(result));
            }           
          }
        }
      }      
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationMatrix() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationMatrix() -- finished");
    }
    private static void GetCappedMatrix(List<string> priceIni)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedMatrix() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedMatrix() -- started");
      const int sourceColumn = 1;
      const int destinationColumn = 2;
      const int bandColumn = 3;
      const string capped = "CAPPED";
      HashSet<string> uniqueBands = new HashSet<string>();

      foreach (var band in StaticVariable.SourceDestinationBands)
      {
        string[] sourceBands = band.Split('\t');
        uniqueBands.Add(sourceBands[bandColumn]);
      }

      foreach (var item in uniqueBands)
      {
        var query =
          from db in StaticVariable.CustomerDetailsDataRecord
          where db.CustomerGroupBand.ToUpper().Equals(item.ToUpper()) || db.StdBand.ToUpper().Equals(item.ToUpper()) && db.ChargingType.ToUpper().Equals(capped)
          select db;

        foreach (var result in query)
        {
          CreateDurationHeader(priceIni, result, "MATRIX CAPPED");
          foreach (var matrix in StaticVariable.SourceDestinationBands)
          {
            string[] sourceBands = matrix.Split('\t');
            if (sourceBands[bandColumn].ToUpper().Equals(item.ToUpper()))
            {
              priceIni.Add("(" + item + ")" + sourceBands[sourceColumn] + "," + sourceBands[destinationColumn] + "," + GetCappedPrices(result));
            }
          }
        }
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedMatrix() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedMatrix() -- finished");
    }
    private static void GetPulseMatrix(List<string> priceIni)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetPulseMatrix() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetPulseMatrix() -- started");
      const int sourceColumn = 1;
      const int destinationColumn = 2;
      const int bandColumn = 3;
      const string pulse = "PULSE";
      HashSet<string> uniqueBands = new HashSet<string>();

      foreach (var band in StaticVariable.SourceDestinationBands)
      {
        string[] sourceBands = band.Split('\t');
        uniqueBands.Add(sourceBands[bandColumn]);
      }

      foreach (var item in uniqueBands)
      {
        var query =
          from db in StaticVariable.CustomerDetailsDataRecord
          where db.CustomerGroupBand.ToUpper().Equals(item.ToUpper()) || db.StdBand.ToUpper().Equals(item.ToUpper()) && db.ChargingType.ToUpper().Equals(pulse)
          select db;

        foreach (var result in query)
        {
          CreateDurationHeader(priceIni, result, "MATRIX PULSE");
          foreach (var matrix in StaticVariable.SourceDestinationBands)
          {
            string[] sourceBands = matrix.Split('\t');
            if (sourceBands[bandColumn].ToUpper().Equals(item.ToUpper()))
            {
              priceIni.Add("(" + item + ")" + sourceBands[sourceColumn] + "," + sourceBands[destinationColumn] + "," + GetPulsePrices(result));
            }
          }
        }
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetPulseMatrix() -- finished");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetPulseMatrix() -- finished");
    }
    private static void CreateDurationHeader(List<string> pricesini, DataRecord dr, string headerName)
    {
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateDurationHeader() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateDurationHeader() -- started");
      pricesini.Add(Environment.NewLine + "[" + headerName + "]");
      pricesini.Add(Constants.AlwaysAddBandHardCoded);
      pricesini.Add("Time Scheme=" + dr.CustomerTimeScheme);
      pricesini.Add("First Rate=" + GetTariffPlanValues("Rate1"));
      pricesini.Add("Second Rate=" + GetTariffPlanValues("Rate2"));
      pricesini.Add("Third Rate=" + GetTariffPlanValues("Rate3"));
      pricesini.Add("Fourth Rate=" + GetTariffPlanValues("Rate4"));
      pricesini.Add("Minimum Duration=" + dr.CustomerMinimumTime);
      pricesini.Add("Dial Time=" + dr.CustomerDialTime);
      pricesini.Add(Constants.TollFreeHardCoded);
      pricesini.Add("All Schemes=" + dr.CustomerAllSchemes);
      pricesini.Add("Minimum Digits=" + dr.CustomerMinDigits);
      pricesini.Add("Minimum Intervals=" + dr.CustomerMinimumIntervals);
      pricesini.Add("Intervals at Initial Cost=" + dr.CustomerIntervalsAtInitialCost);
      string bandDescription = dr.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? dr.CustomerPrefixName : dr.StdPrefixDescription;
      bandDescription = bandDescription.Length > Constants.V5Tc2BandDescriptionLength ? bandDescription.Substring(0, 20) : bandDescription;
      pricesini.Add("Band Description=" + bandDescription);
      pricesini.Add("Interval Rounding=" + AdjustRoundingValueForPricesIni(dr.CustomerRounding));
      pricesini.Add("Initial Interval Length=" + dr.CustomerInitialIntervalLength);
      pricesini.Add("Subsequent Interval Length=" + dr.CustomerSubsequentIntervalLength);
      pricesini.Add("Destination Type=" + dr.CustomerDestinationType);
      string tableName = headerName.ToUpper().Contains("MATRIX") ? StaticVariable.NationalTableSpellingValue : dr.CustomerTableName;
      pricesini.Add("Prefix Table=" + tableName);
      pricesini.Add("Minimum Cost=" + ValidateData.SetToFourDecimalPlaces(dr.CustomerMinCharge));
      pricesini.Add("Connection Charge=" + ValidateData.SetToFourDecimalPlaces(dr.CustomerConnectionCost + Environment.NewLine));
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateDurationHeader() -- finished");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateDurationHeader() -- finished");
    }
    private static void CreateCappedHeader(List<string> pricesini, DataRecord details, string headerName)
    {
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- started");
      pricesini.Add(Environment.NewLine + "[" + headerName + "]");
      pricesini.Add(Constants.AlwaysAddBandHardCoded);
      pricesini.Add("Time Scheme=" + details.CustomerTimeScheme);
      pricesini.Add("First Rate=" + GetTariffPlanValues("Rate1"));
      pricesini.Add("Cap Time=" + details.CustomerSecondInitialRate);
      pricesini.Add("Minimum Duration=" + details.CustomerMinimumTime);
      pricesini.Add("Dial Time=" + details.CustomerDialTime);
      pricesini.Add(Constants.TollFreeHardCoded);
      pricesini.Add("All Schemes=" + details.CustomerAllSchemes);
      pricesini.Add("Minimum Digits=" + details.CustomerMinDigits);
      pricesini.Add("Minimum Intervals=" + details.CustomerMinimumIntervals);
      pricesini.Add("Intervals at Initial Cost=" + details.CustomerIntervalsAtInitialCost);
      string bandDescription = details.CustomerPrefixName.Length > Constants.V5Tc2BandDescriptionLength ? details.CustomerPrefixName.Substring(0, 20) : details.CustomerPrefixName;
      pricesini.Add("Band Description=" + bandDescription);
      pricesini.Add("Interval Rounding=" + AdjustRoundingValueForPricesIni(details.CustomerRounding));
      pricesini.Add("Initial Interval Length=" + details.CustomerInitialIntervalLength);
      pricesini.Add("Subsequent Interval Length=" + details.CustomerSubsequentIntervalLength);
      pricesini.Add("Destination Type=" + details.CustomerDestinationType);
      string tableName = headerName.ToUpper().Contains("MATRIX") ? StaticVariable.NationalTableSpellingValue : details.CustomerTableName;
      pricesini.Add("Prefix Table=" + tableName);
      pricesini.Add("Minimum Cost=" + ValidateData.SetToFourDecimalPlaces(details.CustomerMinCharge));
      pricesini.Add("Connection Charge=" + ValidateData.SetToFourDecimalPlaces(details.CustomerConnectionCost + Environment.NewLine));
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- finished");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- finished");
    }
    private static void CreatePulseHeader(List<string> pricesini, DataRecord details, string headerName)
    {
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- started");
      pricesini.Add(Environment.NewLine + "[" + headerName + "]");
      pricesini.Add(Constants.AlwaysAddBandHardCoded);
      pricesini.Add("Time Scheme=" + details.CustomerTimeScheme);
      pricesini.Add("First Rate=" + GetTariffPlanValues("Rate1"));
      pricesini.Add("Second Rate=" + GetTariffPlanValues("Rate2"));
      pricesini.Add("Third Rate=" + GetTariffPlanValues("Rate3"));
      pricesini.Add("Fourth Rate=" + GetTariffPlanValues("Rate4"));
      pricesini.Add("Minimum Duration=" + details.CustomerMinimumTime);
      pricesini.Add("Dial Time=" + details.CustomerDialTime);
      pricesini.Add(Constants.TollFreeHardCoded);
      pricesini.Add("All Schemes=" + details.CustomerAllSchemes);
      pricesini.Add("Minimum Digits=" + details.CustomerMinDigits);
      string bandDescription = details.CustomerPrefixName.Length > Constants.V5Tc2BandDescriptionLength ? details.CustomerPrefixName.Substring(0, 20) : details.CustomerPrefixName;
      pricesini.Add("Band Description=" + bandDescription);
      pricesini.Add("Minimum Cost=" + ValidateData.SetToFourDecimalPlaces(details.CustomerMinCharge));
      pricesini.Add("Connection Charge=" + ValidateData.SetToFourDecimalPlaces(details.CustomerConnectionCost));
      pricesini.Add("Destination Type=" + details.CustomerDestinationType);
      string tableName = headerName.ToUpper().Contains("MATRIX") ? StaticVariable.NationalTableSpellingValue : details.CustomerTableName;
      pricesini.Add("Prefix Table=" + tableName + Environment.NewLine);
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- finished");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "CreateCappedHeader() -- finished");
    }
    private static string GetDurationPrices(DataRecord result)
    {
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationPrices() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationPrices() -- started");
      string rate1 = result.CustomerFirstInitialRate.Equals(result.CustomerFirstSubseqRate) ? ValidateData.SetToFourDecimalPlaces(result.CustomerFirstInitialRate) : "(" + ValidateData.SetToFourDecimalPlaces(result.CustomerFirstInitialRate) + "," + ValidateData.SetToFourDecimalPlaces(result.CustomerFirstSubseqRate) + ")";
      string rate2 = result.CustomerSecondInitialRate.Equals(result.CustomerSecondSubseqRate) ? ValidateData.SetToFourDecimalPlaces(result.CustomerSecondInitialRate) : "(" + ValidateData.SetToFourDecimalPlaces(result.CustomerSecondInitialRate) + "," + ValidateData.SetToFourDecimalPlaces(result.CustomerSecondSubseqRate) + ")";
      string rate3 = result.CustomerThirdInitialRate.Equals(result.CustomerThirdSubseqRate) ? ValidateData.SetToFourDecimalPlaces(result.CustomerThirdInitialRate) : "(" + ValidateData.SetToFourDecimalPlaces(result.CustomerThirdInitialRate) + "," + ValidateData.SetToFourDecimalPlaces(result.CustomerThirdSubseqRate) + ")";
      string rate4 = result.CustomerFourthInitialRate.Equals(result.CustomerFourthSubseqRate) ? ValidateData.SetToFourDecimalPlaces(result.CustomerFourthInitialRate) : "(" + ValidateData.SetToFourDecimalPlaces(result.CustomerFourthInitialRate) + "," + ValidateData.SetToFourDecimalPlaces(result.CustomerFourthSubseqRate) + ")";
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationPrices() -- finished");
      return rate1 + "," + rate2 + "," + rate3 + "," + rate4;
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetDurationPrices() -- finished");
      
    }
    private static string GetCappedPrices(DataRecord result)
    {
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedPrices() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedPrices() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedPrices() -- finished");
      return ValidateData.SetToFourDecimalPlaces(result.CustomerFirstInitialRate) + "," + ValidateData.SetToFourDecimalPlaces(result.CustomerFirstSubseqRate);
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetCappedPrices() -- finished");
      
    }
    private static string GetPulsePrices(DataRecord result)
    {
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetPulsePrices() -- started");
      //StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "GetPulsePrices() -- started");
      const string pulseZeroRate = "1,0";
      return result.CustomerFirstInitialRate + "," + result.CustomerFirstSubseqRate + "," + pulseZeroRate + "," + pulseZeroRate + "," + pulseZeroRate;
      //Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "GetPulsePrices() -- finished");
    }
    private static List<string> DisplayHolidays(string[] tariffPlanLine)
    {
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "DisplayHolidays() -- started");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "DisplayHolidays() -- started");
      string firstHoliday = string.Empty;
      List<string> holidaysListed = new List<string>();

      var hols = tariffPlanLine[1].Split(',');
      if (hols.Length > 1)
      {
        for (int i = 0; i < hols.Length; i++)
        {
          firstHoliday = i.Equals(0) ? tariffPlanLine[0] : string.Empty;
          holidaysListed.Add(firstHoliday + "\t" + hols[i].Trim());
        }
      }
      else
      {
        holidaysListed.Add(tariffPlanLine[0] + "\t" + tariffPlanLine[1].Trim());
      }
      Console.WriteLine("RearrangeCompletedFiles".PadRight(30, '.') + "DisplayHolidays() -- finish");
      StaticVariable.ConsoleOutput.Add("RearrangeCompletedFiles".PadRight(30, '.') + "DisplayHolidays() -- finish");
      return holidaysListed;
    }
  }
}
