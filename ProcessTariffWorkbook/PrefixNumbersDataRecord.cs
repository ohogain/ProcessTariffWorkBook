using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessTariffWorkbook
{
  public class PrefixNumbersDataRecord
  {
    #region class variables    
    public string TableName { get; set; }
    public string PrefixName { get; set; }
    public string PrefixNumber { get; set; }
    public string stdBand { get; set; }
    public string stdPrefixName { get; set; }
    #endregion

    public PrefixNumbersDataRecord()
    {
    }
    public PrefixNumbersDataRecord(string sLine)
    {
      try
      {
        string[] ary = sLine.Split('\t');
        TableName = ary[0];
        PrefixNumber = ary[1];
        PrefixName = ary[2];
        stdBand = ary[3];
        stdPrefixName = ary[4];
      }
      catch (IndexOutOfRangeException i)
      {
        Console.WriteLine("PrefixNumbersDataRecord: The Index is out of bounds");
        Console.WriteLine(i.Message);
        Console.ReadKey();
      }
      catch (Exception e)
      {
        Console.WriteLine("PrefixNumbersDataRecord: problem with creating the PrefixNumbersDataRecord");
        Console.WriteLine(e.Message);
        Console.ReadKey();
      }
    }
  }
}
