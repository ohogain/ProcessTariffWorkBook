// <copyright file="StaticVariable.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
using System.Collections.Generic;

namespace ProcessTariffWorkbook
{
  public static class StaticVariable
  {
    public static bool MoveOutputSpreadSheetToV6TwbFolder = false;
    public static bool PulseWorkSheetFound = false;
    public static int NumberOfTimeSchemes = 0;
    public static List<DataRecord> CustomerDetailsDataRecord = new List<DataRecord>();
    public static List<PrefixNumbersDataRecord> PrefixNumbersRecord = new List<PrefixNumbersDataRecord>();
    public static List<StandardInternationalBandsDataRecord> StandardInternationalBands =new List<StandardInternationalBandsDataRecord>();
    public static List<string> CombinedRegex = new List<string>();
    public static List<string> ConsoleOutput = new List<string>();
    public static List<string> DefaultEntries = new List<string>();
    public static List<string> DefaultEntriesPrices = new List<string>();
    public static List<string> HolidayList = new List<string>();
    public static List<string> InputXlsxFileDetails = new List<string>();
    public static List<string> PrefixNumbers = new List<string>();
    public static List<string> ProgressDetails = new List<string>();
    public static List<string> SourceDestinationBands = new List<string>();
    public static List<string> Spelling = new List<string>();
    public static List<string> TableLinks = new List<string>();
    public static List<string> TariffPlan = new List<string>();
    public static List<string> TimeSchemes = new List<string>();
    public static List<string> TimeSchemesExceptions = new List<string>();
    public static List<string> TimeSchemesNames = new List<string>();
    public static List<string> TwbHeader = new List<string>();
    public static string CarrierUnitPriceValue = string.Empty;
    public static string CategoryMatrixXlsx = string.Empty;
    public static string CompletedDirectory = null;
    public static string ConsoleErrorFile = string.Empty;
    public static string CountryCodeValue = string.Empty;
    public static string CountryValue = string.Empty;
    public static string DatasetFolderToUse = string.Empty;
    public static string DatasetsFolder = string.Empty;
    public static string DirectoryName = string.Empty;
    public static string ExportNdsValue = string.Empty;
    public static string FinalDirectory = string.Empty;
    public static string HeaderFile = string.Empty;
    public static string Holidays = string.Empty;
    public static string InputFile = string.Empty;
    public static string IntermediateLog = string.Empty;
    public static string InternationalMobileSpellingValue = string.Empty;
    public static string InternationalTableSpellingValue = string.Empty;
    public static string NationalTableSpellingValue = string.Empty;
    public static string ProgressFile = string.Empty;
    public static string Rate1Name = string.Empty;
    public static string Rate2Name = string.Empty;
    public static string Rate3Name = string.Empty;
    public static string Rate4Name = string.Empty;
    public static string ReleaseDateValue = string.Empty;
    public static string TariffPlanNameValue = string.Empty;
    public static string TariffReferenceNumberValue = string.Empty;
    public static string ToTwbFolder = string.Empty;
    public static string V6TwbOutputXlsx = string.Empty;
    public static string V6TwbOutputXlsxFile = string.Empty;
    public static string VersionValue = string.Empty;
    public static string XlsxFileName = string.Empty;
  }
}
