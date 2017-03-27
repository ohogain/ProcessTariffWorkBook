using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace ProcessTariffWorkbook
{
  class Program
  {
    static void Main(string[] args)
    {
      Console.WriteLine("running....");

      ProcessRequiredFiles.GetRequiredData(args);
      ErrorProcessing.CreateIntermediateLog();
      ErrorProcessing.AddRequiredDataDetailsToErrorLog();
      ProcessInputXlsxFile.ParseInputXlsxFile();
      ErrorProcessing.DestinationsAssignedIncorrectTable();
      ErrorProcessing.DestinationsWithoutPrefixes();
      ErrorProcessing.WriteToIntermediateLog();
      RearrangeCompletedFiles.CreateCategoryMatrix();
      RearrangeCompletedFiles.WriteToV6TwbXlsxFile();
      RearrangeCompletedFiles.WriteOutV5Tc2Files();
      RearrangeCompletedFiles.CopyOutputXlsxFileToV6OpUtilFolder(StaticVariable.MoveOutputSpreadSheetToV6TwbFolder);      
      ErrorProcessing.AddMainlandPricesToDependentCountries();
      ErrorProcessing.FindMissingInternationalCountries();


      StaticVariable.Errors.Add(Environment.NewLine + "........finished");
      ErrorProcessing.OutputToErrorLog();
      Console.WriteLine("oxo....");
      ErrorProcessing.OutputConsoleLog();
      MessageBox.Show("oxo");
      Environment.Exit(Constants.KillProgram);
    }
  }
}
