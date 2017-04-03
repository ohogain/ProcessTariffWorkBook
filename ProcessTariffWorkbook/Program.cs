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

      StaticVariable.ProgressDetails.Add(Environment.NewLine + "........finished");
      StaticVariable.ConsoleOutput.Add(Environment.NewLine + "........finished");
      ErrorProcessing.OutputToLogs(StaticVariable.ProgressDetails, StaticVariable.DirectoryName + @"\" + Constants.ProgressLog);
      ErrorProcessing.OutputToLogs(StaticVariable.ConsoleOutput, StaticVariable.DirectoryName + @"\" + Constants.ConsoleErrorLog);
      ErrorProcessing.OutputConsoleLog();
      Console.WriteLine("oxo....");
      MessageBox.Show("oxo");
      Environment.Exit(Constants.KillProgram);
    }
  }
}
