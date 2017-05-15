//---------
// <copyright file="Program.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
//---------

using System;
using System.Windows.Forms;


namespace ProcessTariffWorkbook
{
  class Program
  {
    static void Main(string[] args)
    {
      Console.WriteLine("running....");

      ProcessRequiredFiles.GetDatasetsData(args);          
      ProcessInputXlsxFile.ParseInputXlsxFileIntoCustomerDetailsRecord();    
      ValidateData.PreRegExMatchValidateCustomerDetailsRecord();
      StaticVariable.CustomerDetailsDataRecord.Clear();
      ProcessInputXlsxFile.MatchInputXlsxFileWithRegEx(StaticVariable.InputXlsxFileDetails); 
      Prefixes.ProcessPrefixesData();                         
      ValidateData.PostRegExMatchValidateCustomerDetailsRecord();      
      Prefixes.ValidatePrefixesData();
      ValidateData.DisplayMissingDetails();     
     
      RearrangeCompletedFiles.CreateCategoryMatrix();          
      RearrangeCompletedFiles.WriteToV6TwbXlsxFile();
      RearrangeCompletedFiles.WriteOutV5Tc2Files();                           
      RearrangeCompletedFiles.CopyOutputXlsxFileToV6OpUtilFolder(StaticVariable.MoveOutputSpreadSheetToV6TwbFolder);           
      ErrorProcessing.CreateAndWriteToRegExMatchedLog();           


      //ValidateData.TestMethod();
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
