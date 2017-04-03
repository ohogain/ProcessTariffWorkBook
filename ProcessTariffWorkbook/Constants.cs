//---------
// <copyright file="Constants.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
//---------

namespace ProcessTariffWorkbook
{
  using System.Diagnostics.CodeAnalysis;
  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600:ElementsMustBeDocumented", Justification = "Suppress description for each element")]

  /// <summary>
  /// The class contains all constant variables.
  /// </summary>
 public class Constants
  {    
    public const int KillProgram = 0;
    public const int SplitNameIntoTwoParts = 2;
    public const int NumberOfColumnsXlsxFile = 31;
    public const int NumberOfFieldsExcludingPrefixNameAndRates = 22; // check this figure
    public const int XlsxFileNameSplitIntoTwoParts = 2;
    public const int NumberOfHeadersInHeadersFile = 7;
    public const int V5Tc2BandLengthLimit = 5; // current RMAdmin limit.
    public const int TwbBandLengthLimit = 100; // current TWB (V6) limit 
    public const int V5Tc2BandDescriptionLength = 20;
    public const string alwaysAddBandHardCoded = "Always Add Band=FALSE";
    public const string tollFreeHardCoded = "Toll Free=FALSE";    
    public const string EncodingUnicode = "Unicode";
    public const string XlsxExtension = ".XLSX";
    public const string FiveSpacesPadding = "     ";    
    public const string IniExtensionSearch = @"*ini";
    public const string TxtExtensionSearch = @"*txt";
    public const string TwbWorkSheetName = "TWB";
    public const string Tc2WorkSheetName = "TC2";
    public const string FinalColumnName = "CHARGING TYPE";
    public const string BandsSheet = ".Bands.sheet";
    public const string PrefixBandsSheet = ".PrefixBands.sheet";
    public const string PrefixNumbersSheet = ".PrefixNumbers.sheet";
    public const string SourceDestinationBandsSheet = ".SourceDestinationBands.sheet";
    public const string TableLinksSheet = ".TableLinks.sheet";
    public const string TariffPlanSheet = ".TariffPlan.sheet";
    public const string TimeSchemeExceptionsSheet = ".TimeSchemeExceptions.sheet";
    public const string TimeSchemesSheet = ".TimeSchemes.sheet";
    public const string StdIntAndBands = "Std_Int_Names_Bands.txt";
    public const string IntermediateLog = "IntermediateLog.txt";
    public const string ProgressLog = "ProgressLog.txt";
    public const string ConsoleErrorLog = "ConsoleErrorLog.txt";
    public const string OutputXlsxFile = "OutputXLSX.xlsx";
    public const string TariffPlanName = "TARIFF PLAN NAME";
    public const string OperatorName = "OPERATOR NAME";
    public const string ReleaseDate = "RELEASE DATE";
    public const string EffectiveFrom = "EFFECTIVE FROM";
    public const string Country = "COUNTRY";
    public const string CountryCode = "COUNTRY CODE";
    public const string CurrencyIsoCode = "CURRENCY (ISO CODE)";
    public const string StartingPointTableName = "STARTING POINT TABLE NAME";
    public const string IsPrivate = "IS PRIVATE";
    public const string Rate1 = "RATE1";
    public const string Rate2 = "RATE2";
    public const string Rate3 = "RATE3";
    public const string Rate4 = "RATE4";
    public const string Using = "USING";
    public const string TariffReferenceNumber = "TARIFF REFERENCE NUMBER";
    public const string Version = "VERSION";
    public const string ExportNds = "EXPORTNDS";
    public const string CarrierUnitPrice = "CARRIER UNIT PRICE";
    public const string InternationalMobileSpelling = "INTERNATIONAL MOBILE SPELLING";
    public const string NationalTableSpelling = "NATIONAL TABLE SPELLING";
    public const string InternationalTableSpelling = "INTERNATIONAL TABLE SPELLING";
    public const string Holiday = "HOLIDAY";
    public const string Capped = "CAPPED";
    public const string Pulse = "PULSE";
    public const string Duration = "DURATION";
    public const string Twb = "TWB";
    public const string Tc2 = "TC2";
    public const string CategoryMatrixFile = "CategoryMatrix.xlsx";
    public const string V6TwbDropFolder = @"\op-utils\WorkingFolder\TariffTools";


  }
}
