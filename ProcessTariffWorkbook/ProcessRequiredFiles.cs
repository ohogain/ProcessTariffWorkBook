// <copyright file="ProcessRequiredFiles.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 

namespace ProcessTariffWorkbook
{
  using System;
  using System.Collections.Generic;
  using System.Diagnostics.CodeAnalysis;
  using System.IO;
  using System.Linq;
  using System.Text;  
  using System.Windows.Forms;  
  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600:ElementsMustBeDocumented", Justification = "Suppress description for each element")]
  public class ProcessRequiredFiles
  {
    public static void GetDatasetsData(string[] args)
    {
      GetArguments(args);
      StaticVariable.DirectoryName = GetDirectoryName();
      StaticVariable.XlsxFileName = GetInputXlsxFileName();      
      StaticVariable.CountryCodeValue = GetCountryCode();
      StaticVariable.DatasetFolderToUse = GetDatasetsFolderToUse();
      StaticVariable.HeaderFile = GetHeaderFile();            
      ReadHeaderFileIntoLists();
      ValidateData.CheckTariffPlanList();
      ValidateData.CheckTableLinksList();
      ValidateData.CheckTimeSchemesList();
      ValidateData.CheckTimeSchemeExceptionsList();
      ValidateData.CheckSpellingList();
      ValidateData.CheckSourceDestinationsBandList();
      ValidateData.CheckForStdIntAndBandsFile();   
      RearrangeDefaultEntries();      
      ValidateData.CheckForMoreThanTwoRegExFiles();
      CombineRegExFilesIntoCombinedRegexList(StaticVariable.DatasetFolderToUse);
      CombineRegExFilesIntoCombinedRegexList(StaticVariable.DatasetsFolder); //populates CombinedRegex list. needs more visiblilty    

      CreateFinalFolder();
      StaticVariable.CategoryMatrixXlsxFile = CreateXlsxFileName(Constants.CategoryMatrixFile); 
      CreateOutputXlsxFile(StaticVariable.CategoryMatrixXlsxFile);     
      StaticVariable.V6TwbOutputXlsxFile = CreateXlsxFileName(CreateNewFileName());
      CreateOutputXlsxFile(StaticVariable.V6TwbOutputXlsxFile);                         
    }                       
    public static void GetArguments(string[] args)
      {
        Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetArguments() -- started");
        StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetArguments() -- started");
        StaticVariable.ProgressDetails.Add("Process starting".PadRight(20, '.') + Environment.NewLine);
        string toTwbFolder = string.Empty;
        try
        {
          if (args.Length.Equals(1))
          {
              StaticVariable.InputFile = args[0].Trim();
          }
          else if (args.Length.Equals(2))
          {
            StaticVariable.InputFile = args[0].Trim();
            toTwbFolder = args[1].Trim();
            if (toTwbFolder.ToUpper().Equals("TRUE"))
            {
                StaticVariable.MoveOutputSpreadSheetToV6TwbFolder = true;
            }
          }

          if (!File.Exists(StaticVariable.InputFile))
          {
            Console.WriteLine("The file " + StaticVariable.InputFile + " does not exist. Check the folder." + Environment.NewLine + "The file name must be in this format: countryCode_FileName.txt");
            MessageBox.Show("The file " + StaticVariable.InputFile + " does not exist. Check the folder." + Environment.NewLine + "The file name must be in this format: countryCode_FileName.txt");
            ErrorProcessing.StopProcessDueToFatalError();
          }
        }
        catch (System.Exception e)
        {
          Console.WriteLine("System Error: " + e.Message);
        }
        Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetArguments() -- finished");
        StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetArguments() -- finished");
    }        
    public static string GetDirectoryName()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDirectory() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetDirectory() -- started");
      string directory = string.Empty;
      try
      {
        directory = Path.GetDirectoryName(StaticVariable.InputFile);
      }
      catch (Exception e)
      {
          Console.WriteLine("cannot find directory" + e.Message);
          ErrorProcessing.StopProcessDueToFatalError();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDirectory() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetDirectory() -- finished");
      return directory;
    }        
    public static string GetInputXlsxFileName()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName() -- started");
      string fileName = string.Empty;
      try
      {
        fileName = Path.GetFileName(StaticVariable.InputFile);
      }
      catch (Exception e)
      {
        Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName()");
        Console.WriteLine(e.Message);
        ErrorProcessing.StopProcessDueToFatalError();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName() -- finished");
      return fileName;
    }
    private static void CreateFinalFolder()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder() -- started");
      StaticVariable.FinalDirectory = StaticVariable.DirectoryName + @"\" + StaticVariable.XlsxFileName.Substring(0, StaticVariable.XlsxFileName.IndexOf('.')) + "_Final";
      if (Directory.Exists(StaticVariable.FinalDirectory))
      {
        try
        {          
          Directory.Delete(StaticVariable.FinalDirectory, true);
        }
        catch (IOException io )
        {
          StaticVariable.ProgressDetails.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()");

          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be deleted.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Some of the resons may be:");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "- A file with the same name and location specified by path exists");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "- The directory specified by path is read-only, or recursive is false and path is not an empty directory.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "- The directory is the application's current working directory.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "- The directory contains a read-only file.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "- The directory is being used by another process.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + io.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be deleted");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      try
      {
        Directory.CreateDirectory(StaticVariable.FinalDirectory);
      }
      catch (IOException e)
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::CreateFinalFolder()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be created.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "IO Error: The directory specified by path is a file or The network name is not known.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      catch (Exception e)
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::CreateFinalFolder()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be created.");        
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder() -- finished"); 
    }
    private static string GetCountryCode()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetCountryCode() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetCountryCode() -- started");
      const int xlsxFileNameSplitIntoTwoParts = 2;
      string code = String.Empty;
      try
      {
        string[] xlsxFileNameSplitOnUnderscore = StaticVariable.XlsxFileName.Split('_');        
        if (xlsxFileNameSplitOnUnderscore.Length.Equals(xlsxFileNameSplitIntoTwoParts))
        {
          if (ValidateData.CheckIfInteger(xlsxFileNameSplitOnUnderscore[0]))
          {
            code = xlsxFileNameSplitOnUnderscore[0];
          }
          else
          {
            StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::GetCountryCode(");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The country code is not an integer.");
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        else
        {
          StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::GetCountryCode(");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There is either zero or are more than one '_' underscores in filename." + Environment.NewLine + Constants.FiveSpacesPadding + "There can be only one." + Environment.NewLine + Constants.FiveSpacesPadding + "It must seperate the country code from the name. e.g 353_filename");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      catch (Exception e)
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::GetCountryCode()");
        StaticVariable.ProgressDetails.Add(e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetCountryCode() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetCountryCode() -- finished");
      return code;
    }
    private static string GetDatasetsFolderToUse()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDatasetsFolderToUse() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetDatasetsFolderToUse() -- started");      
      string sPath = string.Empty;
      string datasetFolder = String.Empty;
      int nCount = 0;            
      StaticVariable.DatasetsFolder = StaticVariable.DirectoryName + "\\Datasets";
      if (!Directory.Exists(StaticVariable.DatasetsFolder))
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::GetDatasetsFolderToUse()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Dataset Parent Folder cannot be found");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      else
      {
        string[] datasets = Directory.GetDirectories(StaticVariable.DatasetsFolder);         
        foreach (string tok in datasets)
        {
          string dataset = Path.GetFileName(tok);
          string[] individualDatasets = dataset.Split('_');
          if (individualDatasets[0].Equals(StaticVariable.CountryCodeValue))
          {
            sPath = tok;
            nCount++;
          }
        }
      }
      if (nCount.Equals(1))
      {
        datasetFolder = sPath;
      }
      else
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::GetDatasetsFolderToUse(");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Country Dataset Folder cannot be found or else there is more than one folder with " + StaticVariable.CountryCodeValue + " country code");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDatasetsFolderToUse() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetDatasetsFolderToUse() -- finished");
      return datasetFolder;
    }
    private static string GetHeaderFile()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForHeaderFile() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckForHeaderFile() -- started");
      int nCount = 0;
      string[] aryFiles = Directory.GetFiles(StaticVariable.DatasetFolderToUse, Constants.TxtExtensionSearch);
      string file = String.Empty;      
      foreach (string tok in aryFiles)
      {
        if (tok.ToUpper().Contains("HEADER"))
        {
          nCount++;
          file = tok;
        }
      }
      if (!nCount.Equals(1))
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::CheckForHeaderFile()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + ". Either no header or more than one header file in " + Path.GetFileName(StaticVariable.DatasetFolderToUse));
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForHeaderFile() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckForHeaderFile() -- finished");
      return file;
    }
    private static void ReadHeaderFileIntoLists()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists() -- started");             
      List<string> headerNamesProcessed = new List<string>();
      int tariffPlanCount = 0;
      const int numberOfRequiredEntriesExcludingCarrierUnitPrice = 18;
      const int numberOfHeadersInHeadersFile = 7;
      try
      {
        using (StreamReader oSr = new StreamReader(File.OpenRead(StaticVariable.HeaderFile), Encoding.Unicode))
        {
          while (!oSr.EndOfStream)
          {
            string line = oSr.ReadLine();
            if (!string.IsNullOrEmpty(line) && !line.StartsWith(";"))
            {
              if (line.ToUpper().StartsWith("TITLE="))
              {
                string[] titles = line.Split('=');
                string title = titles[1].ToUpper();

                switch (title)
                {
                  case "TARIFFPLAN":                    
                    while (!line.ToUpper().Equals("ENDTITLE"))
                    {
                      line = oSr.ReadLine().Trim();
                      if (!string.IsNullOrEmpty(line) && !line.StartsWith(";") && !line.ToUpper().Equals("ENDTITLE"))
                      {
                        ValidateData.CheckforAllTariffPlanEntries(line);
                        if (line.ToUpper().Contains(Constants.ReleaseDate))
                        {
                          string[] ary = line.Split('=');
                          StaticVariable.TariffPlan.Add(ary[0] + "=" + ValidateData.CreateDate());
                        }
                        else if (line.ToUpper().Contains(Constants.EffectiveFrom))
                        {
                          string[] ary = line.Split('=');
                          StaticVariable.TariffPlan.Add(ary[0] + "=" + ValidateData.CreateDate());
                        }
                        else
                        {
                          StaticVariable.TariffPlan.Add(line);
                        }
                        if (!line.ToUpper().Contains(Constants.CarrierUnitPrice.ToUpper()))
                        {
                          tariffPlanCount++;
                        }                        
                      }                      
                    }
                    if (!tariffPlanCount.Equals(numberOfRequiredEntriesExcludingCarrierUnitPrice))
                    {
                      StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessRequiredFiles::CheckforAllTariffPlanEntries()");
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There are missing entries in the Tariff Plan header. These are the required entries.");
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.TariffPlanName);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.OperatorName);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.ReleaseDate);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.EffectiveFrom);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Country);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.CountryCode);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.CurrencyIsoCode);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.StartingPointTableName);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.IsPrivate);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Rate1);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Rate2);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Rate3);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Rate4);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Using);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.TariffReferenceNumber);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Version);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.ExportNds);
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.Holiday);
                      StaticVariable.ProgressDetails.Add(Environment.NewLine + "If pulse is not being used, Carrier Unit Price should be deleted / commented out");
                      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + Constants.CarrierUnitPrice);
                      ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                    }
                    headerNamesProcessed.Add("TARIFFPLAN");
                    break;
                  case "TABLELINKS":
                    while (!line.ToUpper().Equals("ENDTITLE"))
                    {
                      try
                      {
                        line = oSr.ReadLine().Trim();
                        if (!string.IsNullOrEmpty(line) && !line.StartsWith(";") && !line.ToUpper().Equals("ENDTITLE"))
                        {
                          StaticVariable.TableLinks.Add(line.Trim());
                        }
                      }
                      catch (NullReferenceException e)
                      {
                        Console.WriteLine(e);
                        throw;
                      }
                    }
                    headerNamesProcessed.Add("TABLELINKS");
                    break;
                  case "TIMESCHEMES":
                    while (!line.ToUpper().Equals("ENDTITLE"))
                    {
                      line = oSr.ReadLine().Trim();
                      if (!string.IsNullOrEmpty(line) && !line.StartsWith(";") && !line.ToUpper().Equals("ENDTITLE"))
                      {
                        StaticVariable.TimeSchemes.Add(line);
                      }
                    }
                    headerNamesProcessed.Add("TIMESCHEMES");
                    break;
                  case "TIMESCHEMEEXCEPTIONS":
                    while (!line.ToUpper().Equals("ENDTITLE"))
                    {
                      line = oSr.ReadLine().Trim();
                      if (!string.IsNullOrEmpty(line) && !line.StartsWith(";") && !line.ToUpper().Equals("ENDTITLE"))
                      {
                        StaticVariable.TimeSchemesExceptions.Add(line);
                      }
                    }
                    headerNamesProcessed.Add("TIMESCHEMEEXCEPTIONS");
                    break;
                  case "DEFAULTENTRIES":
                    while (!line.ToUpper().Equals("ENDTITLE"))
                    {
                      line = oSr.ReadLine().Trim();
                      if (!string.IsNullOrEmpty(line) && !line.StartsWith(";") && !line.ToUpper().Equals("ENDTITLE"))
                      {
                        StaticVariable.DefaultEntriesPrices.Add(line);
                      }
                    }
                    headerNamesProcessed.Add("DEFAULTENTRIES");
                    break;
                  case "SOURCEDESTINATIONBANDS":
                    while (!line.ToUpper().Equals("ENDTITLE"))
                    {
                      line = oSr.ReadLine().Trim();
                      if (!string.IsNullOrEmpty(line) && !line.StartsWith(";") && !line.ToUpper().Equals("ENDTITLE"))
                      {
                        StaticVariable.SourceDestinationBands.Add(line);
                      }
                    }
                    headerNamesProcessed.Add("SOURCEDESTINATIONBANDS");
                    break;
                  case "SPELLING":
                    while (!line.ToUpper().Equals("ENDTITLE"))
                    {
                      line = oSr.ReadLine().Trim();
                      if (!string.IsNullOrEmpty(line) && !line.StartsWith(";") && !line.ToUpper().Equals("ENDTITLE"))
                      {
                        StaticVariable.Spelling.Add(line);
                      }
                    }
                    headerNamesProcessed.Add("SPELLING");
                    break;
                  default:
                    StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::ReadHeaderFileIntoLists()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There may be a Header undefined in the Header file.");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + title);
                    ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                    break;
                }
              }
            }
          }
          oSr.Close();
        }
      }
      catch (Exception e)
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::ReadHeaderFileIntoLists()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "Problem with reading the Header file");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      if (!headerNamesProcessed.Count.Equals(numberOfHeadersInHeadersFile))
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::ReadHeaderFileIntoLists()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + ". The missing Header is.");
        string[] headerTitles = { "TARIFFPLAN", "TABLELINKS", "TIMESCHEMES", "TIMESCHEMEEXCEPTIONS", "SPELLING", "DEFAULTENTRIES", "SOURCEDESTINATIONBANDS" };
        foreach (string tok in headerTitles)
        {
          bool bFound = false;
          foreach (string name in headerNamesProcessed)
          {
            if (name.ToUpper().Equals(tok.ToUpper()))
            {
              bFound = true;
              break;
            }
          }
          if (!bFound && !headerNamesProcessed.Count.Equals(numberOfHeadersInHeadersFile))
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + tok);
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists() -- finished");
    }            
    private static string CreateNewFileName()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateTWBFileName() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateTWBFileName() -- started");      
      string monthNumber = string.Empty;
      string fileName = string.Empty;
      const int year = 2;
      const int day = 0;
      const int month = 1;
      const int dateLength = 3;
      string[] dateTokens = StaticVariable.ReleaseDateValue.Split('-');

      if (dateTokens[month].Length.Equals(dateLength))
      {
        switch (dateTokens[month].ToUpper())
        {
          case "JAN":
            monthNumber = "01";
            break;
          case "FEB":
            monthNumber = "02";
            break;
          case "MAR":
            monthNumber = "03";
            break;
          case "APR":
            monthNumber = "04";
            break;
          case "MAY":
            monthNumber = "05";
            break;
          case "JUN":
            monthNumber = "06";
            break;
          case "JUL":
            monthNumber = "07";
            break;
          case "AUG":
            monthNumber = "08";
            break;
          case "SEP":
            monthNumber = "09";
            break;
          case "OCT":
            monthNumber = "10";
            break;
          case "NOV":
            monthNumber = "11";
            break;
          case "DEC":
            monthNumber = "12";
            break;
        }        
        fileName = StaticVariable.CountryCodeValue + "_" + StaticVariable.CountryValue + "_" + StaticVariable.TariffReferenceNumberValue + "_" + StaticVariable.TariffPlanNameValue.Replace(" ", "") + "_" + StaticVariable.VersionValue + "_" + (dateTokens[year] + monthNumber + dateTokens[day]) + ".xlsx";
      }
      else
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::CreateTWBFileName()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Release Date is not in the correct format. It must be like so: DD-Mmm-YYYY");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateTWBFileName() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateTWBFileName() -- finished");
      return fileName;
    }
    private static string CreateXlsxFileName(string fileName)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateXlsxFileName( " + Path.GetFileName(fileName) + " ) -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateXlsxFileName( " + Path.GetFileName(fileName) + " ) -- started");      
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateXlsxFileName( " + Path.GetFileName(fileName) + " ) -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateXlsxFileName( " + Path.GetFileName(fileName) + " ) -- finished");
      return StaticVariable.FinalDirectory + @"\" + fileName; ;
    }
    private static void CreateOutputXlsxFile(string file)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateOutputXlsxFile( " + Path.GetFileName(file) + " ) -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateOutputXlsxFile( " + Path.GetFileName(file) + " ) -- started");            
      try
      {
        File.Create(file);
      }
      catch (IOException io)
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessRequiredFiles::CreateOutputXlsxFile( " + Path.GetFileName(file) + " )");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "An I/O error occurred while creating the file " + Path.GetFileName(file) + " file");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + io.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      catch (Exception e)
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "ProcessRequiredFiles::CreateOutputXlsxFile( " + Path.GetFileName(file) + " )");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Problem creating output " + Path.GetFileName(file) + " file");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateOutputXlsxFile( " + Path.GetFileName(file) + " ) -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateOutputXlsxFile( " + Path.GetFileName(file) + " ) -- finished");
    }
    private static void RearrangeDefaultEntries()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries()-- starting");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries() -- started");
      const int numberOfFieldsExcludingPrefixNameAndRates = 22;
      string prices = string.Empty;
      string timeScheme = string.Empty;
      string minimumTime = string.Empty;
      string dialTime = string.Empty;
      string allSchemes = string.Empty;
      string minimumDigits = string.Empty;
      string minimumIntervals = string.Empty;
      string intervalsAtInitialCost = string.Empty;
      string wholeIntervalCharging = string.Empty;
      string initialIntervalLength = string.Empty;
      string subsequentIntervalLength = string.Empty;
      string destinationType = string.Empty;
      string prefixTable = string.Empty;
      string minimumCost = string.Empty;
      string connectionCharge = string.Empty;      
      string destination = string.Empty;      
      string usingGroupBands = string.Empty;
      string groupBand = string.Empty;
      string groupDescription = string.Empty;
      string usingCustomerNames = string.Empty;
      string multiLevelEnabled = string.Empty;
      string cutOff1Cost = string.Empty;
      string cutOff2Duration = string.Empty;
      string chargingType = string.Empty;       
      const int entryValue = 1;
      const int oneEntry = 1;
      const int entryName = 0;
      const int eightPricesPlusPrefixName = 9;
      const int splitOnEqualsSign = 2;      
      int priceLineCount = 0;
      int fieldLineCount = 0;      
      int numberOfPricesInFile = 0;
      int numberOfFieldsInFile = 0;


      foreach (string tok in StaticVariable.DefaultEntriesPrices)
      {
        if (tok.Contains('='))
        {
          string[] lines = tok.Split('=');
          if (lines.Length.Equals(splitOnEqualsSign))
          {
            switch (lines[entryName].ToUpper())
            {
              case "TIME SCHEME":
                timeScheme = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "MINIMUM TIME":
                minimumTime = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "DIAL TIME":
                dialTime = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "ALL SCHEMES":
                allSchemes = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "MINIMUM DIGITS":
                minimumDigits = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "MINIMUM INTERVALS":
                minimumIntervals = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "INTERVALS AT INITIAL COST":
                intervalsAtInitialCost = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "WHOLE INTERVAL CHARGING":
                wholeIntervalCharging = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "INITIAL INTERVAL LENGTH":
                initialIntervalLength = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "SUBSEQUENT INTERVAL LENGTH":
                subsequentIntervalLength = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "DESTINATION TYPE":
                destinationType = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "PREFIX TABLE":
                prefixTable = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "MINIMUM COST":
                minimumCost = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "CONNECTION CHARGE":
                connectionCharge = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "USING GROUP BANDS":
                usingGroupBands = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "GROUP BAND":
                groupBand = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "GROUP BAND DESCRIPTION":
                groupDescription = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "USING CUSTOMERS NAMES":
                usingCustomerNames = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "MULTI LEVEL ENABLED":
                multiLevelEnabled = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "CUT OFF1 COST":
                cutOff1Cost = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "CUT OFF2 DURATION":
                cutOff2Duration = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              case "CHARGING TYPE":
                chargingType = lines[entryValue].Trim();
                fieldLineCount++;
                numberOfFieldsInFile++;
                break;
              default:
                StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The column entry is incorrect - " + tok + ". There may be an column not catered for.");
                DisplayAllHeadersFieldsUsed();
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                break;
            }
          }
          else
          {
            StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The column entry is incorrect: " + tok + ". The name must be seperated by an '=' sign.");
            DisplayAllHeadersFieldsUsed();
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        else if (tok.Contains(','))
        {
          try
          {
            string[] lines = tok.Split(',');
            if (lines.Length.Equals(eightPricesPlusPrefixName))
            {
              destination = lines[0];
              StringBuilder sb = new StringBuilder();
              for (int i = 1; i < lines.Length; i++)
              {
                sb.Append(lines[i]);
                sb.Append("\t");
              }
              prices = sb.ToString();
              prices = prices.TrimEnd('\t');
              priceLineCount++;
              numberOfPricesInFile++;
            }
            else
            {
              StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The prices entry is incorrect: " + tok + ". There should be 8 prices and a destination after the (band), comma seperated");              
            }
          }
          catch (Exception e)
          {
            StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The prices entry is incorrect." + tok);
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        if (fieldLineCount.Equals(numberOfFieldsExcludingPrefixNameAndRates) && priceLineCount.Equals(oneEntry))
        {
          string line = ValidateData.CapitaliseWord(destination) + "\t" + prices + "\t" + minimumCost + "\t" + connectionCharge + "\t" + usingGroupBands.ToUpper() + "\t" +
            groupBand.ToUpper() + "\t" + ValidateData.CapitaliseWord(groupDescription) + "\t" + ValidateData.CapitaliseWord(prefixTable) + "\t" + ValidateData.CapitaliseWord(destinationType) + "\t" +
            wholeIntervalCharging.ToUpper() + "\t" + ValidateData.CapitaliseWord(timeScheme) + "\t" + usingCustomerNames.ToUpper() + "\t" + initialIntervalLength + "\t" +
            subsequentIntervalLength + "\t" + minimumIntervals + "\t" + intervalsAtInitialCost + "\t" + minimumTime + "\t" + dialTime + "\t" +
            allSchemes.ToUpper() + "\t" + multiLevelEnabled.ToUpper() + "\t" + minimumDigits + "\t" + cutOff1Cost + "\t" + cutOff2Duration + "\t" + chargingType.ToUpper();

          StaticVariable.DefaultEntries.Add(line);
          priceLineCount = 0;
          fieldLineCount = 0;
        }
      }
      if (!numberOfPricesInFile.Equals(numberOfFieldsInFile / numberOfFieldsExcludingPrefixNameAndRates))
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There may be an extra or missing field / prices entry in Default Entries.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There should be " + numberOfFieldsExcludingPrefixNameAndRates + " fields for every price section found");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Prefix Name and the 8 prices are excluded from the fields. The fields required are:" + Environment.NewLine);
        DisplayAllHeadersFieldsUsed();

        /*StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "ALL SCHEMES" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "CHARGING TYPE" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "CONNECTION CHARGE" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "CUT OFF1 COST" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "CUT OFF2 DURATION" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "DESTINATION TYPE" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "DIAL TIME" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "GROUP BAND DESCRIPTION" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "GROUP BAND" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "INITIAL INTERVAL LENGTH" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "INTERVALS AT INITIAL COST" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "MINIMUM COST" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "MINIMUM DIGITS" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "MINIMUM INTERVALS" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "MINIMUM TIME" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "MULTI LEVEL ENABLED" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "PREFIX TABLE" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "SUBSEQUENT INTERVAL LENGTH" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "TIME SCHEME" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "USING CUSTOMERS NAMES" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "USING GROUP BANDS" + Environment.NewLine +
                          Constants.FiveSpacesPadding + "WHOLE INTERVAL CHARGING" + Environment.NewLine);*/
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Total Number of fields = " + numberOfFieldsInFile + " and Total Number of Prices = " + numberOfPricesInFile);
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The total number of fields (" + numberOfFieldsInFile + ") / number of fields (" + numberOfFieldsExcludingPrefixNameAndRates + ") per price is " + (numberOfFieldsInFile / numberOfFieldsExcludingPrefixNameAndRates) + " for every " + numberOfPricesInFile + " prices found.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "These figures should be equal.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If number is different. There are either too many/few fields or too many/few prices.");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries() -- finished");
    }           
    private static void CombineRegExFilesIntoCombinedRegexList(string sRegExPath)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CombineRegExFiles() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CombineRegExFiles() -- started");      
      string[] aryFiles = Directory.GetFiles(sRegExPath, Constants.TxtExtensionSearch);
      if (aryFiles.Length > 0)
      {
        foreach (string tok in aryFiles)
        {
          if (tok.ToUpper().Contains("REGEX"))
          {
            try
            {
              using (StreamReader oSr = new StreamReader(File.OpenRead(tok), Encoding.Unicode))
              {
                while (!oSr.EndOfStream)
                {
                  string line = oSr.ReadLine().TrimEnd('\t');
                  if (!string.IsNullOrEmpty(line) && !line.Contains(";"))
                  {
                    StaticVariable.CombinedRegex.Add(line);
                  }
                }
                oSr.Close();
              }
            }
            catch (Exception e)
            {
              StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::CombineRegExFiles()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + /*sFileName + */"RegEx files cannot be read");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
        }
      }
      else
      {
        StaticVariable.ProgressDetails.Add("ProcessRequiredFiles::CombineRegExFiles()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + /*sFileName + */"No RegEx files in RegEx Folder");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      StaticVariable.CombinedRegex = StaticVariable.CombinedRegex.Distinct().ToList();
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CombineRegExFiles() -- finished");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CombineRegExFiles() -- finished");
    }
    private static void DisplayAllHeadersFieldsUsed()
    {
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Fields used are listed below");
      StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "All Schemes" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Charging Type" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Connection Charge" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Cut Off1 Cost" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Cut Off2 Duration" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Destination Type" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Dial Time" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Group Band Description" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Group Band" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Initial Interval Length" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Intervals At Initial Cost" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Minimum Cost" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Minimum Digits" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Minimum Intervals" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Minimum Time" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Multi Level Enabled" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Prefix Table" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Subsequent Interval Length" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Time Scheme" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Using Customers Names" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Using Group Bands" + Environment.NewLine +
        Constants.FiveSpacesPadding + "Whole Interval Charging" + Environment.NewLine);
    }   
  }
}
