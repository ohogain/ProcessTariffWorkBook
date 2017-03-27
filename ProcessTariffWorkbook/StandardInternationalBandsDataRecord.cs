//---------
// <copyright file="StandardInternationalBandsDataRecord.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
//---------
using System;

namespace ProcessTariffWorkbook
{
  public class StandardInternationalBandsDataRecord
  {
    public string SBand { get; set; }
    public string SPrefixName { get; set; }
    public string SCountryCode { get; set; }

    public StandardInternationalBandsDataRecord()
    {

    }

    public StandardInternationalBandsDataRecord(string line)
    {      
      try
      {
        string[] ary = line.Split('\t');
        SBand = ary[0].ToUpper();
        SPrefixName = ary[1].ToUpper();
        SCountryCode = ary[2].ToUpper();
      }
      catch (Exception e)
      {
        Console.WriteLine("StandardInternationalBandsDataRecord:: problem with creating the Std Int Bands Data Record");
        Console.WriteLine(e.Message);
        Console.ReadKey();
      }
    }
  }
}
