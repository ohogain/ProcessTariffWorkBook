using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ProcessTariffWorkbook
{
  class FindDestinationsMissingFromBaseRates
  {
     # region variables
    private static string inputFile_CurrentIntermediateFile = string.Empty;
    private static string inputFile_Eircom_BT_KPN_ETC_IntermediateFile = string.Empty;
    private static string outputFile = string.Empty;
    private static string errorMessage = string.Empty;
    private static string carrierName = string.Empty;    
    const int numberOfColumns = 33;
    private static List<string> ResultList = new List<string>();
    private static Dictionary <string, string> oDict = new Dictionary<string, string>();
    public static bool processTariffWorkbook = false;
    # endregion

    static void NewMain(string[] args)
    {
      if (!args.Length.Equals(2))
      {
        Console.WriteLine("2 files required: input A, input B "); //output file should not exist
        Console.ReadKey();
      }
      else
      {
        Console.WriteLine("Job running....");
        inputFile_CurrentIntermediateFile = args[0].Trim();
        inputFile_Eircom_BT_KPN_ETC_IntermediateFile = args[1].Trim();

        if (!File.Exists(inputFile_CurrentIntermediateFile))
        {
          Console.WriteLine("arg 1 does not exist");
        }
        if (!File.Exists(inputFile_Eircom_BT_KPN_ETC_IntermediateFile))
        {
          Console.WriteLine("arg 2 does not exist");
        }
        else
        {
          outputFile = Path.GetDirectoryName(inputFile_CurrentIntermediateFile);
          outputFile = outputFile + @"\_Std_Destinations_Missing_From_Original_File.txt";
          if (File.Exists(outputFile))
            File.Delete(outputFile);

          FillDictionaryWithInputFile_0();
          MatchTokens();
          WriteOutResults();          
          Console.WriteLine("Job's OXO" + Environment.NewLine);
          Console.WriteLine("Output File is : " + outputFile);
          Console.ReadKey();

        }
      }
    }
    //==========
    private static void WriteOutResults()
    {
      string newLine = string.Empty;
      string[] ary;
      using (StreamWriter oFileOut = new StreamWriter(File.OpenWrite(outputFile), Encoding.UTF8))
      {
        ResultList = ResultList.Distinct().ToList();
        foreach (string tok in ResultList)
        {
          ary = tok.Split('\t');
          if (!tok.ToUpper().Contains("AZORES") && !tok.ToUpper().Contains("BALEARIC") &&
            !tok.ToUpper().Contains("CANAR") && !tok.ToUpper().Contains("CEUTA") &&
            !tok.ToUpper().Contains("CYPRUS NORTH") && !tok.ToUpper().Contains("MADEIRA") &&
            !tok.ToUpper().Contains("ALASKA") && !tok.ToUpper().Contains("HAWAI") && 
            /*!tok.ToUpper().Contains("NGN") &&*/ !ary[0].TrimEnd('-').ToUpper().Equals("INTERNATIONAL"))
          {
            if (tok.ToUpper().Contains("NGN"))
            {
              newLine = tok.Replace("NGN", "Mobile");
              newLine = ";" + newLine;
              oFileOut.WriteLine(RearrangeResultsListToMatchInputFile(newLine));
            }
            else
            {
              oFileOut.WriteLine(RearrangeResultsListToMatchInputFile(tok));
            }
          }
        }
        oFileOut.Close();
      }
    }
    //==========
    private static void MatchTokens()
    {
      string[] aryLine;
      string prefixName = string.Empty;
      string band = string.Empty;

      using (StreamReader oStreamIn = new StreamReader(File.OpenRead(inputFile_Eircom_BT_KPN_ETC_IntermediateFile)))
      {
        while (!oStreamIn.EndOfStream)
        {
          string sCurrentLine = oStreamIn.ReadLine();
          if (!String.IsNullOrEmpty(sCurrentLine) && !sCurrentLine.StartsWith(";"))
          {
            aryLine = sCurrentLine.Split(new char[] { '\t', ',' });
            prefixName = aryLine[0].Trim();
            band = aryLine[1].Trim();
            try
            {
              string sTemp = oDict[band.ToUpper()];
            }
            catch (KeyNotFoundException)
            {
              //errorMessage = "this key: " + sCurrentLine + " was not found";
              ResultList.Add(sCurrentLine);
            }
          }
        }
        oStreamIn.Close();
      }
    }
    //==========
    private static void FillDictionaryWithInputFile_0()
    {
      string[] aryLine;
      string element0 = string.Empty;
      string element1 = string.Empty;
      //const int processTariffWorkbookTotalColumns = 32;

      using (StreamReader oStreamIn = new StreamReader(File.OpenRead(inputFile_CurrentIntermediateFile)))
      {
        while (!oStreamIn.EndOfStream)
        {
          string sCurrentLine = oStreamIn.ReadLine();          
          if (!String.IsNullOrEmpty(sCurrentLine) && !sCurrentLine.StartsWith(";"))
          {
            aryLine = sCurrentLine.Split(new char[] { '\t', ',' });
            element0 = aryLine[0].Trim();
            element1 = aryLine[1].Trim();           
            //ResultList.Add(sCurrentLine); //use this line if you need to add all of file 1 to output.
            try
            {
              oDict.Add(element1.ToUpper(), sCurrentLine); //element1.ToUpper(), element0
            }
            catch (ArgumentNullException ex)
            {
              errorMessage = "Error 1: Key is null: " + ex.Message;
              ResultList.Add(errorMessage);
              WriteOutResults();
            }
            catch (ArgumentException ex)
            {
              errorMessage = "Error 2: " + ex.Message + " : " + sCurrentLine.ToUpper();
              ResultList.Add(errorMessage);
              WriteOutResults();
            }
          }
        }
        oStreamIn.Close();
      }  
    }
    //==========
    private static void FillDictionaryWithTWBInputFile_0()
    {
      string[] aryLine;
      string element0 = string.Empty;
      string element1 = string.Empty;

      using (StreamReader oStreamIn = new StreamReader(File.OpenRead(inputFile_CurrentIntermediateFile)))
      {
        while (!oStreamIn.EndOfStream)
        {
          string sCurrentLine = oStreamIn.ReadLine();
          if (!String.IsNullOrEmpty(sCurrentLine) && !sCurrentLine.StartsWith(";"))
          {
            aryLine = sCurrentLine.Split(new char[] { '\t', ',' });
            element0 = aryLine[0].Trim();
            element1 = aryLine[1].Trim();            
            try
            {
              oDict.Add(element0.ToUpper(), sCurrentLine); //element1.ToUpper(), element0
            }
            catch (ArgumentNullException ex)
            {
              errorMessage = "Error 1: Key is null: " + ex.Message;
              ResultList.Add(errorMessage);
              WriteOutResults();
            }
            catch (ArgumentException ex)
            {
              errorMessage = "Error 2: " + ex.Message + " : " + sCurrentLine.ToUpper();
              ResultList.Add(errorMessage);
              WriteOutResults();
            }
          }
        }
        oStreamIn.Close();
      }
    }
    private static string RearrangeResultsListToMatchInputFile(string line)
    {
      # region variables
      int prefixNameCol = 2;
      int CheapCol = 28;
      int CheapRateSubseqCol = 29;
      int StandardCol = 31;
      int StdRateSubseqCol = 32;
      int EconomyCol = 34;
      int EconomyRateSubseqCol = 35;
      int PeakCol = 37;
      int PeakRateSubseqCol = 38;
      int MinChargeCol = 10;
      int ConnChargeCol = 11;
      string UsingGroupBandsValue = "false";
      int GroupBandCol = 3;
      int GroupIDDDescriptionCol = 4;
      int TableNameCol = 6;
      int DestinationTypeCol = 7;
      int RoundingCol = 16;
      int TimeSchemeCol = 14;
      string UsingCustomerNamesValue = "false";
      int InitialIntervalLengthCol = 17;
      int SubsequentIntervalLengthCol = 18;
      int MinimumIntervalsCol = 19;
      int IntervalsAtInitialCostCol = 20;
      int MinTimeCol = 8;
      int DialTimeCol = 9;
      int TollFreeCol = 12;
      int AllSchemesCol = 13;
      int MultiLevelEnabledCol = 26;
      int AlwaysAddBandCol = 39;
      int MinDigitsCol = 15;
      int ChargingModeCol = 5;
      int CustomerChargingTablesCol = 21;
      int FlatCostCol = 22;
      int FlatIntervalLengthCol = 23;
      int CutOff1CostCol = 24;
      int CutOff2DurationCol = 25;
      string PricesSuppliedValue = "TRUE";
      int ChargingTypeCol = 40;
      string[] ary;
      string newLine = "Error: in splitting line in RearrangeResultsList";
      # endregion
      
      try
      {
        ary = line.Split(new char[] { '\t', ',' });
        newLine = (ary[prefixNameCol].TrimEnd('-') +"\t"+ ary[CheapCol] +"\t"+ ary[CheapRateSubseqCol] +"\t"+ 
              ary[StandardCol] +"\t"+ ary[StdRateSubseqCol] +"\t"+ ary[EconomyCol] +"\t"+
              ary[EconomyRateSubseqCol] +"\t"+ ary[PeakCol] +"\t"+ ary[PeakRateSubseqCol] +"\t"+
              ary[MinChargeCol] + "\t" + ary[ConnChargeCol] + "\t" + UsingGroupBandsValue +"\t"+
              ary[GroupBandCol] + "\t" + ary[GroupIDDDescriptionCol] + "\t" + carrierName + "\t" + 
              ary[TableNameCol] +"\t"+ ary[DestinationTypeCol] +"\t"+ ary[RoundingCol] +"\t"+ 
              ary[TimeSchemeCol] +"\t"+ UsingCustomerNamesValue +"\t"+ ary[InitialIntervalLengthCol] +"\t"+ 
              ary[SubsequentIntervalLengthCol] +"\t"+ ary[MinimumIntervalsCol] +"\t"+ 
              ary[IntervalsAtInitialCostCol] +"\t"+ ary[MinTimeCol] +"\t"+ ary[DialTimeCol] +"\t"+ 
              ary[TollFreeCol] +"\t"+ ary[AllSchemesCol] +"\t"+ ary[MultiLevelEnabledCol] +"\t"+ 
              ary[AlwaysAddBandCol] +"\t"+ ary[MinDigitsCol] +"\t"+ ary[ChargingModeCol] +"\t"+ 
              ary[CustomerChargingTablesCol] +"\t"+ ary[FlatCostCol] +"\t"+ ary[FlatIntervalLengthCol] +"\t"+
              ary[CutOff1CostCol] +"\t"+ ary[CutOff2DurationCol] +"\t"+ PricesSuppliedValue +"\t"+
              ary[ChargingTypeCol]);
      }
      catch (Exception e)
      {        
        Console.WriteLine(e.Message);
      }

      return newLine;
    }
  }
}
