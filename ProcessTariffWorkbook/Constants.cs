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
        public const int V5Tc2BandDescriptionLength = 20;
        public const int V5Tc2BandLengthLimit = 5; // current RMAdmin limit.
        public const string AlwaysAddBandHardCoded = "Always Add Band=FALSE";
        public const string BandsSheet = ".Bands.sheet";
        public const string Capped = "CAPPED";
        public const string CarrierUnitPrice = "CARRIER UNIT PRICE";
        public const string CategoryMatrixFile = "CategoryMatrix.xlsx";
        public const string ConsoleErrorLog = "ConsoleErrorLog.txt";
        public const string Country = "COUNTRY";
        public const string CountryCode = "COUNTRY CODE";
        public const string CurrencyIsoCode = "CURRENCY (ISO CODE)";
        public const string Duration = "DURATION";
        public const string EffectiveFrom = "EFFECTIVE FROM";
        public const string EncodingUnicode = "Unicode";
        public const string ExportNds = "EXPORTNDS";
        public const string FinalColumnName = "CHARGING TYPE";
        public const string FiveSpacesPadding = " ";
        public const string Holiday = "HOLIDAY";
        public const string InternationalMobileSpelling = "INTERNATIONAL MOBILE SPELLING";
        public const string InternationalTableSpelling = "INTERNATIONAL TABLE SPELLING";
        public const string IsPrivate = "IS PRIVATE";
        public const string NationalTableSpelling = "NATIONAL TABLE SPELLING";
        public const string OperatorName = "OPERATOR NAME";
        public const string ProgressLog = "ProgressLog.txt";
        public const string Pulse = "PULSE";
        public const string Rate1 = "RATE1";
        public const string Rate2 = "RATE2";
        public const string Rate3 = "RATE3";
        public const string Rate4 = "RATE4";        
        public const string ReleaseDate = "RELEASE DATE";
        public const string StartingPointTableName = "STARTING POINT TABLE NAME";
        public const string StdIntAndBands = "Std_Int_Names_Bands.txt";
        public const string TableLinks = "TABLELINKS";
        public const string TariffPlanName = "TARIFF PLAN NAME";
        public const string TariffReferenceNumber = "TARIFF REFERENCE NUMBER";
        public const string TollFreeHardCoded = "Toll Free=FALSE";
        public const string TxtExtensionSearch = @"*txt";
        public const string Using = "USING";
        public const string Version = "VERSION";        
    }
}
