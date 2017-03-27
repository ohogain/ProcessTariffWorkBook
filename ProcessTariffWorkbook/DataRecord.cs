// <copyright file="DataRecord.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessTariffWorkbook
{
  public class DataRecord
  {
    #region class variables
    public string StdBand { get; set; }
    public string StdPrefixName { get; set; }
    public string StdPrefixDescription { get; set; }
    public string CustomerPrefixName { get; set; }
    public string CustomerFirstInitialRate { get; set; }
    public string CustomerFirstSubseqRate { get; set; }
    public string CustomerSecondInitialRate { get; set; }
    public string CustomerSecondSubseqRate { get; set; }
    public string CustomerThirdInitialRate { get; set; }
    public string CustomerThirdSubseqRate { get; set; }
    public string CustomerFourthInitialRate { get; set; }
    public string CustomerFourthSubseqRate { get; set; }
    public string CustomerMinCharge { get; set; }
    public string CustomerConnectionCost { get; set; }
    public string CustomerUsingGroupBands { get; set; }
    public string CustomerGroupBand { get; set; }
    public string CustomerGroupBandDescription { get; set; }
    public string CustomerTableName { get; set; }
    public string CustomerDestinationType { get; set; }
    public string CustomerRounding { get; set; }
    public string CustomerTimeScheme { get; set; }
    public string CustomerUsingCustomerNames { get; set; }
    public string CustomerInitialIntervalLength { get; set; }
    public string CustomerSubsequentIntervalLength { get; set; }
    public string CustomerMinimumIntervals { get; set; }
    public string CustomerIntervalsAtInitialCost { get; set; }
    public string CustomerMinimumTime { get; set; }
    public string CustomerDialTime { get; set; }
    public string CustomerAllSchemes { get; set; }
    public string CustomerMultiLevelEnabled { get; set; }
    public string CustomerMinDigits { get; set; }
    public string CustomerCutOff1Cost { get; set; }
    public string CustomerCutOff2Duration { get; set; }
    public string ChargingType { get; set; }
    #endregion

    public DataRecord()
    {
    }    
    public DataRecord(string line)
    {      
      try
      {
        string[] ary = line.Split('\t');
        StdBand = ary[0].ToUpper().Trim();
        StdPrefixName = ary[1].Trim();
        StdPrefixDescription = ary[2].Trim();
        CustomerPrefixName = ary[3].Trim();
        CustomerFirstInitialRate = ary[4].Trim();
        CustomerFirstSubseqRate = ary[5].Trim();
        CustomerSecondInitialRate = ary[6].Trim();
        CustomerSecondSubseqRate = ary[7].Trim();
        CustomerThirdInitialRate = ary[8].Trim();
        CustomerThirdSubseqRate = ary[9].Trim();
        CustomerFourthInitialRate = ary[10].Trim();
        CustomerFourthSubseqRate = ary[11].Trim();
        CustomerMinCharge = ary[12].Trim();
        CustomerConnectionCost = ary[13].Trim();
        CustomerUsingGroupBands = ary[14].ToUpper().Trim();
        CustomerGroupBand = ary[15].ToUpper().Trim();
        CustomerGroupBandDescription = ary[16].Trim();
        CustomerTableName = ary[17].Trim();
        CustomerDestinationType = ary[18].Trim();
        CustomerRounding = ValidateData.AdjustRoundingValue(ary[19]).ToUpper().Trim();
        CustomerTimeScheme = ary[20].Trim();
        CustomerUsingCustomerNames = ary[21].ToUpper().Trim();
        CustomerInitialIntervalLength = ary[22].ToUpper().Trim();
        CustomerSubsequentIntervalLength = ary[23].ToUpper().Trim();
        CustomerMinimumIntervals = ary[24].ToUpper().Trim();
        CustomerIntervalsAtInitialCost = ary[25].ToUpper().Trim();
        CustomerMinimumTime = ary[26].ToUpper().Trim();
        CustomerDialTime = ary[27].ToUpper().Trim();
        CustomerAllSchemes = ary[28].ToUpper().Trim();
        CustomerMultiLevelEnabled = ary[29].ToUpper().Trim();
        CustomerMinDigits = ary[30].ToUpper().Trim();
        CustomerCutOff1Cost = ary[31].ToUpper().Trim();
        CustomerCutOff2Duration = ary[32].ToUpper().Trim();
        ChargingType = ary[33].ToUpper().Trim();
      }
      catch (IndexOutOfRangeException i)
      {
        Console.WriteLine("DataRecord: The Index is out of bounds");
        Console.WriteLine(i.Message);
        Console.ReadKey();
      }
      catch (Exception e)
      {
        Console.WriteLine("DataRecord: problem with creating the DataRecord");
        Console.WriteLine(e.Message);
        Console.ReadKey();
      }      
    }
  }
}
