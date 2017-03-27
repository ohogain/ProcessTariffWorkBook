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
    #region variables
        
    #endregion
    public static void GetRequiredData(string[] args)
    {
      GetArguments(args);
      StaticVariable.DirectoryName = GetDirectoryName();
      StaticVariable.XlsxFileName = GetInputXlsxFileName();
      CreateFinalFolder();
      StaticVariable.CountryCode = GetCountryCode();
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
         
      StaticVariable.CategoryMatrixXlsx = CreateXlsxFileName(Constants.CategoryMatrixFile); 
      CreateOutputXlsxFile(StaticVariable.CategoryMatrixXlsx);

      StaticVariable.V6TwbOutputXlsx = CreateNewFileName();
      StaticVariable.V6TwbOutputXlsxFile = CreateXlsxFileName(StaticVariable.V6TwbOutputXlsx);
      CreateOutputXlsxFile(StaticVariable.V6TwbOutputXlsxFile);

      RearrangeDefaultEntries();
      ReadPrefixesIntoList(StaticVariable.DatasetFolderToUse, "Domestic");
      ReadPrefixesIntoList(StaticVariable.DatasetsFolder, "International");
      CombinePrefixesInDataRecord(StaticVariable.PrefixNumbers);
      ValidateData.CheckForMoreThanTwoRegExFiles();
      CombineRegExFiles(StaticVariable.DatasetFolderToUse);
      CombineRegExFiles(StaticVariable.DatasetsFolder); //populates CombinedRegex list. needs more visiblilty

      Console.WriteLine("End of Dataset data");
    }                       
    public static void GetArguments(string[] args)
      {
        Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetArguments()-- started");
        StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetArguments()");
        StaticVariable.Errors.Add("Process starting".PadRight(20, '.') + Environment.NewLine);
        try
        {
          if (args.Length.Equals(1))
          {
              StaticVariable.InputFile = args[0].Trim();
          }
          else if (args.Length.Equals(2))
          {
            StaticVariable.InputFile = args[0].Trim();
            StaticVariable.ToTwbFolder = args[1].Trim();
            if (StaticVariable.ToTwbFolder.ToUpper().Equals("TRUE"))
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
        Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetArguments()-- finished");        
      }        
    public static string GetDirectoryName()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDirectory()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetDirectory()");
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
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDirectory()-- finished");
      return directory;
    }        
    public static string GetInputXlsxFileName()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName()");
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
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetFileName()-- finished");
      return fileName;
    }
    private static void CreateFinalFolder()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()");
      StaticVariable.FinalDirectory = StaticVariable.DirectoryName + @"\" + StaticVariable.XlsxFileName.Substring(0, StaticVariable.XlsxFileName.IndexOf('.')) + "_Final";
      if (Directory.Exists(StaticVariable.FinalDirectory))
      {
        try
        {          
          Directory.Delete(StaticVariable.FinalDirectory, true);
        }
        catch (IOException io )
        {
          StaticVariable.Errors.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()");

          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be deleted.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Some of the resons may be:");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "- A file with the same name and location specified by path exists");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "- The directory specified by path is read-only, or recursive is false and path is not an empty directory.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "- The directory is the application's current working directory.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "- The directory contains a read-only file.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "- The directory is being used by another process.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + io.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        catch (Exception e)
        {
          StaticVariable.Errors.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be deleted");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      try
      {
        Directory.CreateDirectory(StaticVariable.FinalDirectory);
      }
      catch (IOException e)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CreateFinalFolder()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be created.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "IO Error: The directory specified by path is a file or The network name is not known.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CreateFinalFolder()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.FinalDirectory + " - Final folder cannot be created.");        
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()-- finished");
      //StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateFinalFolder()-- finished"); 
    }
    private static string GetCountryCode()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetCountryCode()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetCountryCode()");
      string code = String.Empty;
      try
      {
        string[] xlsxFileNameSplitOnUnderscore = StaticVariable.XlsxFileName.Split('_');        
        if (xlsxFileNameSplitOnUnderscore.Length.Equals(Constants.XlsxFileNameSplitIntoTwoParts))
        {
          if (ValidateData.CheckIfInteger(xlsxFileNameSplitOnUnderscore[0]))
          {
            code = xlsxFileNameSplitOnUnderscore[0];
          }
          else
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::GetCountryCode(");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The country code is not an integer.");
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        else
        {
          StaticVariable.Errors.Add("ProcessRequiredFiles::GetCountryCode(");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There is either zero or are more than one '_' underscores in filename." + Environment.NewLine + Constants.FiveSpacesPadding + "There can be only one." + Environment.NewLine + Constants.FiveSpacesPadding + "It must seperate the country code from the name. e.g 353_filename");
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::GetCountryCode()");
        StaticVariable.Errors.Add(e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetCountryCode()-- finished");
      return code;
    }
    private static string GetDatasetsFolderToUse()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDatasetsFolderToUse()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "GetDatasetsFolderToUse()");      
      string sPath = string.Empty;
      string datasetFolder = String.Empty;
      int nCount = 0;            
      StaticVariable.DatasetsFolder = StaticVariable.DirectoryName + "\\Datasets";
      if (!Directory.Exists(StaticVariable.DatasetsFolder))
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::GetDatasetsFolderToUse()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Dataset Parent Folder cannot be found");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      else
      {
        string[] datasets = Directory.GetDirectories(StaticVariable.DatasetsFolder);         
        foreach (string tok in datasets)
        {
          string dataset = Path.GetFileName(tok);
          string[] individualDatasets = dataset.Split('_');
          if (individualDatasets[0].Equals(StaticVariable.CountryCode))
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
        StaticVariable.Errors.Add("ProcessRequiredFiles::GetDatasetsFolderToUse(");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Country Dataset Folder cannot be found or else there is more than one folder with " + StaticVariable.CountryCode + " country code");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "GetDatasetsFolderToUse()-- finished");
      return datasetFolder;
    }
    private static string GetHeaderFile()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForHeaderFile()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckForHeaderFile()");
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
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckForHeaderFile()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + ". Either no header or more than one header file in " + Path.GetFileName(StaticVariable.DatasetFolderToUse));
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForHeaderFile()-- finished");
      return file;
    }
    private static void ReadHeaderFileIntoLists()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists()");             
      List<string> headerNamesProcessed = new List<string>();             
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
                      }
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
                    StaticVariable.Errors.Add("ProcessRequiredFiles::ReadHeaderFileIntoLists()");
                    StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There may be a Header undefined in the Header file.");
                    StaticVariable.Errors.Add(Constants.FiveSpacesPadding + title);
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
        StaticVariable.Errors.Add("ProcessRequiredFiles::ReadHeaderFileIntoLists()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + "Problem with reading the Header file");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      if (!headerNamesProcessed.Count.Equals(Constants.NumberOfHeadersInHeadersFile))
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::ReadHeaderFileIntoLists()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + StaticVariable.XlsxFileName + ". The missing Header is.");
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
          if (!bFound && !headerNamesProcessed.Count.Equals(Constants.NumberOfHeadersInHeadersFile))
          {
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + tok);
          }
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists()-- finished");
      //StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "ReadHeaderFileIntoLists()-- finished");
    }            
    private static string CreateNewFileName()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateTWBFileName() -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateTWBFileName()");      
      string monthNumber = string.Empty;
      string fileName = string.Empty;
      const int year = 2;
      const int day = 0;
      const int month = 1;
      const int dateLength = 3;
      string[] dateTokens = StaticVariable.ReleaseDate.Split('-');

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
        fileName = StaticVariable.CountryCode + "_" + StaticVariable.Country + "_" + StaticVariable.TariffReferenceNumber + "_" + StaticVariable.TariffPlanName.Replace(" ", "") + "_" + StaticVariable.Version + "_" + (dateTokens[year] + monthNumber + dateTokens[day]) + ".xlsx";
      }
      else
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CreateTWBFileName()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The Release Date is not in the correct format. It must be like so: DD-Mmm-YYYY");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateTWBFileName() -- Finished");
      return fileName;
    }
    private static string CreateXlsxFileName(string fileName)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateXlsxFileName( " + Path.GetFileName(fileName) + " ) -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateXlsxFileName( " + Path.GetFileName(fileName) + " )");      
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateXlsxFileName( " + Path.GetFileName(fileName) + " )-- Finished");
      return StaticVariable.FinalDirectory + @"\" + fileName; ;
    }
    private static void CreateOutputXlsxFile(string file)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateOutputXlsxFile( " + Path.GetFileName(file) + " ) -- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CreateOutputXlsxFile( " + Path.GetFileName(file) + " )");      
      try
      {
        File.Create(file);
      }
      catch (IOException io)
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ProcessRequiredFiles::CreateOutputXlsxFile( " + Path.GetFileName(file) + " )");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "An I/O error occurred while creating the file " + Path.GetFileName(file) + " file");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + io.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add(Environment.NewLine + "ProcessRequiredFiles::CreateOutputXlsxFile( " + Path.GetFileName(file) + " )");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Problem creating output " + Path.GetFileName(file) + " file");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CreateOutputXlsxFile( " + Path.GetFileName(file) + " )-- Finished");      
    }
    private static void RearrangeDefaultEntries()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries()-- starting");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries()");
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
                StaticVariable.Errors.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
                StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The column entry is incorrect - " + tok + ". There may be an column not catered for.");
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
                break;
            }
          }
          else
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The column entry is incorrect: " + tok + ". The name must be seperated by an '=' sign.");
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
              StaticVariable.Errors.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The prices entry is incorrect: " + tok + ". There should be 8 prices and a destination after the (band), comma seperated");              
            }
          }
          catch (Exception e)
          {
            StaticVariable.Errors.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The prices entry is incorrect." + tok);
            StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        if (fieldLineCount.Equals(Constants.NumberOfFieldsExcludingPrefixNameAndRates) && priceLineCount.Equals(oneEntry))
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
      if (!numberOfPricesInFile.Equals(numberOfFieldsInFile / Constants.NumberOfFieldsExcludingPrefixNameAndRates))
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::RearrangeDefaultEntries()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There may be an extra or missing field / prices entry in Default Entries.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "There should be " + Constants.NumberOfFieldsExcludingPrefixNameAndRates + " fields for every price section found");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Prefix Name and the 8 prices are excluded from the fields. The fields required are:" + Environment.NewLine);
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "ALL SCHEMES" + Environment.NewLine +
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
                          Constants.FiveSpacesPadding + "WHOLE INTERVAL CHARGING" + Environment.NewLine);
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "Total Number of fields = " + numberOfFieldsInFile + " and Total Number of Prices = " + numberOfPricesInFile);
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "The total number of fields (" + numberOfFieldsInFile + ") / number of fields (" + Constants.NumberOfFieldsExcludingPrefixNameAndRates + ") per price is " + (numberOfFieldsInFile / Constants.NumberOfFieldsExcludingPrefixNameAndRates) + " for every " + numberOfPricesInFile + " prices found.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "These figures should be equal.");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + "If number is different. There are either too many/few fields or too many/few prices.");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries()-- finished");
      //ConsoleOutputList.Add("ProcessRequiredFiles".PadRight(30, '.') + "RearrangeDefaultEntries()-- finished");
    }    
    private static void CopyIniFilesToFinalFolder(string[] folder, string message)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CopyIniFiles( " + message + " )-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CopyIniFiles()");      
      try
      {
        foreach (string tok in folder)
        {          
          File.Copy(tok, StaticVariable.FinalDirectory + @"\" + Path.GetFileName(tok));
        }
      }
      catch (Exception e)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::MoveIniFilesToFinalFolder()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + message + " INI files cannot be copied to Final folder");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
        Console.WriteLine(e.Message);
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CopyIniFiles( " + message + " )-- finished");      
    }    
    private static void CombinePrefixesInDataRecord(List<string> lst)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CombinePrefixesInDataRecord()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CombinePrefixesInDataRecord()");
      lst = lst.Distinct().ToList();
      foreach (var tok in lst)
      {
        StaticVariable.PrefixNumbersRecord.Add(new PrefixNumbersDataRecord(tok));
      }
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CombinePrefixesInDataRecord()-- finished");      
    }
    private static void ReadPrefixesIntoList(string folder, string message)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "ReadPrefixesIntoList( " + message + " )-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "ReadPrefixesIntoList( " + message + " )");      
      string[] folders = Directory.GetFiles(folder, Constants.IniExtensionSearch);      
      string tableName = string.Empty;
      string prefixName = string.Empty;
      string prefixNumber = string.Empty;
      foreach (string token in folders)
      {
        try
        {
          using (StreamReader oSr = new StreamReader(File.OpenRead(token), Encoding.Unicode))
          {
            while (!oSr.EndOfStream)
            {
              string line = oSr.ReadLine();
              if (!string.IsNullOrEmpty(line) && !line.StartsWith(";"))
              {
                if (line.ToUpper().Contains("TABLE NAME="))
                {
                  string[] lines = line.Split('=');
                  tableName = StaticVariable.CountryCode + "_" + lines[1];
                }
                if (line.Contains(','))
                {
                  string[] lines = line.Split(',');
                  prefixName = lines[0];
                  prefixNumber = lines[1];
                }
                if (!string.IsNullOrEmpty(tableName) && !string.IsNullOrEmpty(prefixName))
                {
                  StaticVariable.PrefixNumbers.Add(ValidateData.CapitaliseWord(tableName) + "\t" + prefixNumber + "\t" + ValidateData.CapitaliseWord(prefixName));
                }
              }
            }
            oSr.Close();
            tableName = string.Empty;
            prefixName = string.Empty;
            prefixNumber = string.Empty;
          }
        }
        catch (Exception e)
        {
          StaticVariable.Errors.Add("ProcessRequiredFiles::CombinePrefixes()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + message + ": Problem adding prefixes to PrefixNumbersDataRecord.");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      StaticVariable.PrefixNumbers = StaticVariable.PrefixNumbers.Distinct().ToList();
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "ReadPrefixesIntoList( " + message + " )-- finished");
    }
    private static void CheckForDuplicateIniFilesInFinalFolder()
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForDuplicateIniFilesInFinalFolder()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CheckForDuplicateIniFilesInFinalFolder()");
      int intCount = 0;
      int mobileCount = 0;      
      string[] files = Directory.GetFiles(StaticVariable.FinalDirectory, Constants.IniExtensionSearch);
      if (files.Length.Equals(0))
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckForDuplicateIniFilesInFinalFolder()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " There are no INI files in the Final folder");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      foreach (string tok in files)
      {
        string fileName = Path.GetFileName(tok);
        if (fileName.ToUpper().Contains("INT"))
        {
          intCount++;
        }
        if (fileName.ToUpper().Contains("MOB"))
        {
          mobileCount++;
        }
      }
      if (intCount > 1)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckForDuplicateIniFilesInFinalFolder()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " There are " + intCount + " International INI files in the Final folder. Remove one");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      if (mobileCount > 1)
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CheckForDuplicateIniFilesInFinalFolder()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " There are " + mobileCount + " Mobile INI files in the Final folder. Remove one");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      /*foreach (string token in files) //deletes ini files from final folder
      {
        try
        {
          File.Delete(token);
        }
        catch (Exception e)
        {
          StaticVariable.Errors.Add("ProcessRequiredFiles::CheckForDuplicateIniFilesInFinalFolder()");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + " Problems deleting the ini files in the final folder");
          StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }*/
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CheckForDuplicateIniFilesInFinalFolder()-- finished");      
    }
    private static void CombineRegExFiles(string sRegExPath)
    {
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CombineRegExFiles()-- started");
      StaticVariable.ConsoleOutput.Add("ProcessRequiredFiles".PadRight(30, '.') + "CombineRegExFiles()");      
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
              StaticVariable.Errors.Add("ProcessRequiredFiles::CombineRegExFiles()");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + /*sFileName + */"RegEx files cannot be read");
              StaticVariable.Errors.Add(Constants.FiveSpacesPadding + e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
          }
        }
      }
      else
      {
        StaticVariable.Errors.Add("ProcessRequiredFiles::CombineRegExFiles()");
        StaticVariable.Errors.Add(Constants.FiveSpacesPadding + /*sFileName + */"No RegEx files in RegEx Folder");
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      StaticVariable.CombinedRegex = StaticVariable.CombinedRegex.Distinct().ToList();
      Console.WriteLine("ProcessRequiredFiles".PadRight(30, '.') + "CombineRegExFiles()-- finished");      
    }
    
  }
}
