// <copyright file="Prefixes.cs" company="Soft-ex Ltd">
//    Copyright (c) Soft-ex Ltd. All rights reserved.
// </copyright>
// <author>Tomas Ohogain</author> 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Timers;
using System.IO;
  

namespace ProcessTariffWorkbook
{
  public static class Prefixes
  {
    public static void ProcessPrefixesData()
    {
      Console.WriteLine("Prefixes".PadRight(30, '.') + "ProcessPrefixesData() -- started");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "ProcessPrefixesData() -- started");
      StaticVariable.CountryExceptions = CountryExceptions();      
      ReadPrefixesIntoList(StaticVariable.DatasetFolderToUse, "Domestic");
      ReadPrefixesIntoList(StaticVariable.DatasetsFolder, "International"); 
      MatchPrefixNamesWithRegEx(StaticVariable.PrefixNumbersFromIniFiles);
      MatchPrefixNamesAndAddToPrefixesDataRecord(StaticVariable.PrefixesMatchedByRegEx, StaticVariable.PrefixNumbersFromIniFiles);                 
      Console.WriteLine("Prefixes".PadRight(30, '.') + "ProcessPrefixesData() -- finished");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "ProcessPrefixesData() -- finished");
    }
    public static void ValidatePrefixesData()
    {
        Console.WriteLine("Prefixes".PadRight(30, '.') + "ValidatePrefixesData() -- started");
        StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "ValidatePrefixesData() -- started");                      
        CheckForDestinationsWithoutPrefixes();
        CheckForDuplicatePrefixNumbers();
        CheckForNonMatchingPrefixNames();         
        Console.WriteLine("Prefixes".PadRight(30, '.') + "ValidatePrefixesData() -- finished");
        StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "ValidatePrefixesData() -- finished");
    }     
    private static void MatchPrefixNamesAndAddToPrefixesDataRecord(List<string> matchedPrefixNames, List<string> prefixNumbers)
    {
        Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNumbersToRegExStandardNames() -- started");
        StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "MatchPrefixNumbersToRegExStandardNames() -- started");
        const int band = 0;
        const int stdName = 1;
        const int prefixName = 2;         

        prefixNumbers.Sort();
        matchedPrefixNames.Sort();

        foreach (var mpn in matchedPrefixNames)
        {
          string[] matchedName = mpn.Split('\t');
          foreach (var pn in prefixNumbers)
          {
              string[] prefixes = pn.Split('\t');
              if (matchedName[prefixName].ToUpper().Equals(prefixes[prefixName].ToUpper()))
              {
                StaticVariable.PrefixNumbersRecord.Add(new PrefixNumbersDataRecord(pn + "\t" + matchedName[band] + "\t" + matchedName[stdName]));                 
              }
          }
        }
        Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNumbersToRegExStandardNames() -- finished");
        StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "MatchPrefixNumbersToRegExStandardNames() -- finished");
    }      
    private static void CheckForDestinationsWithoutPrefixes()
    {
      Console.WriteLine("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- started");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- started");
      Dictionary<string, string> tableAndNames = new Dictionary<string, string>();
      HashSet<string> sourceDestinations = ValidateData.GetSourceAndDestinationNames();      
                
      var queryMissingPrefixesAndTables =
        (from drm in StaticVariable.CustomerDetailsDataRecord
        orderby drm.StdPrefixName
        select new { drm.StdPrefixName, drm.CustomerTableName}).Distinct();

      foreach(var entry in queryMissingPrefixesAndTables)
      {
        try
        {
          tableAndNames.Add(entry.StdPrefixName.ToUpper(), entry.StdPrefixName.PadRight(40) + " " + entry.CustomerTableName);
        }
        catch(Exception e)
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::DestinationsWithoutPrefixes()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error adding to tableAndNames dictionary");               
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }

      var queryMissingPrefixes =
        from drm in StaticVariable.CustomerDetailsDataRecord
        orderby drm.StdPrefixName
        select drm.StdPrefixName.ToUpper();

      var queryPrefixNames =
        (from pn in StaticVariable.PrefixNumbersRecord           
        orderby pn.stdPrefixName         
        select pn.stdPrefixName.ToUpper()).Distinct();
         
      var missingPrefixes = queryMissingPrefixes.Except(queryPrefixNames).ToList();         
      missingPrefixes.Sort();

      if (missingPrefixes.Any())
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::DestinationsWithoutPrefixes()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "No Prefix Found or the prefix name does not match the Standard or customer name:"); 
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Change the prefix name to match either the Standard or customer name:");                      
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Or else the prefix for that name may not exist in the prefix table." + Environment.NewLine);
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Standard Name".PadRight(41, ' ') + "Table");
        List<string> specificCountries = new List<string>();
        foreach (var entry in missingPrefixes)
        {
          try
          {            
            if(sourceDestinations.Contains(entry.ToUpper())) continue;
            if(StaticVariable.MissingCountryExceptions.Contains(entry.ToUpper()))
            {
              specificCountries.Add(entry.ToUpper());
            }
            else
            {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + ValidateData.CapitaliseWord(tableAndNames[entry.ToUpper()]));
            }                           
          }
          catch(Exception e)
          {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::DestinationsWithoutPrefixes()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error finding table name for prefix name");                     
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }                  
        }
        if(specificCountries.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::DestinationsWithoutPrefixes()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "These prefixes may exist in the prefix file."); 
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If so change the prefix name to match the Xlsx name:");                                         
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Xlsx name:" + Environment.NewLine);

          foreach(var item in specificCountries)
          {
            try
            {
              var query =
                from db in StaticVariable.CustomerDetailsDataRecord
                where item.ToUpper().Equals(db.StdPrefixName.ToUpper())
                select db.CustomerPrefixName;
              foreach(var name in query)
              {
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + ValidateData.CapitaliseWord(name));
              }              
            }
            catch(Exception e)
            {
              StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::DestinationsWithoutPrefixes()");
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Error adding specific countries - " + item);                     
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
              ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            
          } 
        }        
      }
      Console.WriteLine("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- finished");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- finished");
    }
    private static void CheckForNonMatchingPrefixNames()
    {
        Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForNonMatchingPrefixNames() -- started");
        StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "CheckForNonMatchingPrefixNames() -- started");
        List<string> tmpList = new List<string>();
        const int padding50Chars = 50;

        var queryDifferentNames =
          (from pn in StaticVariable.PrefixNumbersRecord
          join cd in StaticVariable.CustomerDetailsDataRecord on pn.stdPrefixName.ToUpper() equals cd.StdPrefixName.ToUpper()
          where !pn.PrefixName.ToUpper().Equals(cd.StdPrefixName.ToUpper()) && !pn.PrefixName.ToUpper().Equals(cd.CustomerPrefixName.ToUpper())
          select new { pn.PrefixName, cd.StdPrefixName, cd.CustomerPrefixName}).Distinct();

        if (queryDifferentNames.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::CheckForNonMatchingPrefixNames()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Prefix names that do not match the standard name nor the name in the input xlsx sheet.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Check if the are matched correctly" + Environment.NewLine);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Prefix Name".PadRight(padding50Chars) + " : " + "Standard Name".PadRight(padding50Chars) + " : Input xlsx Name");
          foreach (var names in queryDifferentNames)
          {
              StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + names.PrefixName.PadRight(padding50Chars) + " : " + names.StdPrefixName.PadRight(padding50Chars) + " : " + names.CustomerPrefixName);
          }
          //ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForNonMatchingPrefixNames() -- finished");
        StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "CheckForNonMatchingPrefixNames() -- finished");
    }
    private static void MatchPrefixNamesWithRegEx(List<string> prefixesFromIniFiles)
    {
      Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNamesWithRegEx() -- started");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "MatchPrefixNamesWithRegEx() -- started");
      const int prefixElement = 1;
      const int prefixNameElement = 2;
      const int tableElement = 0;
      Dictionary<string, string> prefixNamesMatchedMultipleTimes = new Dictionary<string, string>();
      Dictionary<string, string> bandsMatched = new Dictionary<string, string>();
      List<string> prefixNamesOnly = new List<string>();          
      string destinationName = string.Empty;         
      string regExExtraSpaces = "\x0020{2,}";
      string regExNoise = @"\(|\)|( -|- )|\,|'";         
      Timer newTimer = new System.Timers.Timer(10000); // 2 sec interval
      var dupeBandsForSamePrefixName = new List<string>();
      var prefixNamesMatchedMoreThanOnce = new List<string>();
      var prefixNamesNotMatchedByRegex = new List<string>(); 
         
      Regex regExRemoveNoise = new Regex(regExNoise, RegexOptions.Compiled);
      Regex regExRemoveExtraSpaces = new Regex(regExExtraSpaces, RegexOptions.Compiled);                                  
      var sourceDestinations = ValidateData.GetSourceAndDestinationNames();
      foreach (var names in prefixesFromIniFiles)
      {
        string[] name = names.Split('\t');            
        if(sourceDestinations.Contains(name[prefixNameElement].ToUpper())) continue;
        prefixNamesOnly.Add(name[prefixNameElement]);        
      }
      prefixNamesOnly = prefixNamesOnly.Distinct().ToList();
      prefixNamesOnly.Sort();

      foreach (string prefixName in prefixNamesOnly)
      {
        var found = false;
        try
        {
          destinationName = prefixName;
          destinationName = regExRemoveNoise.Replace(destinationName, " ");
          destinationName = regExRemoveExtraSpaces.Replace(destinationName, " ");
        }
        catch (Exception e)
        {
          StaticVariable.ProgressDetails.Add("Prefixes::MatchPrefixNamesWithRegEx()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There may be a problem with the input destination");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
        foreach (string regexExpression in StaticVariable.CombinedRegex)
        {
          string[] aryLine = regexExpression.Split(new char[] {'\t'});
          string regExPattern = aryLine[0];
          string regExBand = aryLine[1].Trim();
          string regexStandardName = aryLine[2].Trim();

          try
          {
            var regExCountry = new Regex(regExPattern, RegexOptions.IgnoreCase);
            
              if(StaticVariable.MissingCountryExceptions.Contains(destinationName.ToUpper()))
              {
                try
                {
                  destinationName = StaticVariable.CountryExceptions[destinationName.ToUpper()];
                }
                catch(Exception e)
                {
                  StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::MatchPrefixNamesWithRegEx()");
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The prefix name was not found in the exceptions list");                  
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
                  ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();                
                }     
              }                              
                    
            if(regExCountry.IsMatch(destinationName))
            {
              newTimer.Elapsed += OnTimedEvent;
              newTimer.AutoReset = true;
              newTimer.Enabled = true;
              found = true;
              try
              {
                bandsMatched.Add(regExBand, prefixName.PadRight(40) + "\t" + regExBand.PadRight(10) + "\t" + regexStandardName.PadRight(40));
              }
              catch(Exception)
              {
                dupeBandsForSamePrefixName.Add(bandsMatched[regExBand] + " : " + prefixName);
              }
              try
              {
                prefixNamesMatchedMultipleTimes.Add(prefixName, prefixName.PadRight(40) + "\t" + regExBand.PadRight(10) + "\t" + regexStandardName.PadRight(40));
              }
              catch(Exception)
              {
                prefixNamesMatchedMoreThanOnce.Add(prefixNamesMatchedMultipleTimes[prefixName] + " : " + regExBand.PadRight(10) + "\t" + regexStandardName);
              }
              StaticVariable.PrefixesMatchedByRegEx.Add(regExBand + "\t" + regexStandardName + "\t" + prefixName);                
            }               
          }
          catch (Exception e)
          {
            StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::MatchPrefixNamesWithRegEx()");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "There may be a problem with the regex");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Check if there are more than 1 international or domestic regex files");
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
            ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
          }
        }
        if (!found)
        {
          prefixNamesNotMatchedByRegex.Add(destinationName);
        }
      }
      if (dupeBandsForSamePrefixName.Any())
      {
        dupeBandsForSamePrefixName.Sort();
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::MatchPrefixNamesWithRegEx()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The Band is assigned to more than one to Prefix Name.");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The same Regex might be matching the 'First Matched Prefix Name' and the 'Subsequent Matched Prefix Name'.");
        StaticVariable.ProgressDetails.Add( Constants.FiveSpacesPadding + "The same Prefix name may be duplicated in other files (eg. satellite & international, if prefixes are supplied");               
        StaticVariable.ProgressDetails.Add( Constants.FiveSpacesPadding + "Are there two International ini or two mobile ini files. Delete the non required files." + Environment.NewLine);               
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "First Matched Prefix Name".PadRight(40) + "\tDuplicate Band".PadRight(10) + "\tStandard Name".PadRight(40) + "  : Subsequent Matched Prefix Name");
        foreach (var item in dupeBandsForSamePrefixName)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + item);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      if (prefixNamesMatchedMoreThanOnce.Any())
      {
        prefixNamesMatchedMoreThanOnce.Sort();
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::MatchPrefixNamesWithRegEx()");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Prefix Name assigned to more than one band");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "E.g. If Inmarsat B. Add new entry to input Xlsx file and change to 'Inmarsat B Land' & 'Inmarsat B Maritime'. ");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If prefixes were supplied you may need to delete these bands in RegEx file and add 'Inmarsat B' regex" + Environment.NewLine);
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "PrefixName".PadRight(40) + "\t" + "Band".PadRight(10) + "\t" + "Standard Name".PadRight(40) + " : Other Band".PadRight(10) + "\tStandard Prefix Name");
        foreach (var item in prefixNamesMatchedMoreThanOnce)
        {
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + item);
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
      }
      ////
      /*if (prefixNamesNotMatchedByRegex.Any()) //Is this needed?
      {
        StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::MatchPrefixNamesWithRegEx(). Is this needed?");
        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Prefixes without regex.");
        prefixNamesNotMatchedByRegex.Sort();
        foreach (string entry in prefixNamesNotMatchedByRegex)
        {
          if(StaticVariable.CountryExceptions.Contains(entry.ToUpper()))
          {
            StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + entry);
          }          
        }
        ErrorProcessing.StopProcessDueToFatalErrorOutputToLog(); 
      }*/
      /////
      newTimer.Enabled = false;
      Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNamesWithRegEx() -- finished");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "MatchPrefixNamesWithRegEx() -- finished");
    }
    public static void CheckForDuplicatePrefixNumbers()
    {
        Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForDuplicatePrefixNumbers() -- started");
        StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "CheckForDuplicatePrefixNumbers() -- started");
        Dictionary<string, string> dict = new Dictionary<string, string>();
        List<string> duplicates = new List<string>();
        string previousNumber = string.Empty;
        string previousName = string.Empty;
        string previousTable = string.Empty;

        var queryPrefixNumbersTotal =
          from db in StaticVariable.PrefixNumbersRecord
          where !db.PrefixNumber.Equals("?")
          orderby db.PrefixNumber
          select db.PrefixNumber;

        foreach (var dupe in queryPrefixNumbersTotal)
        {
          try
          {
              dict.Add(dupe, dupe);
          }
          catch (Exception)
          {
              duplicates.Add(dupe);
          }
        }
        if (duplicates.Any())
        {
          StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::CheckForDuplicatePrefixNumbers()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Some prefixes may be in more than one file. Delete any ini files not required.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "If the tables are different the prefixes may be valid in both files.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Duplicate prefixes in the same table are incorrect." + Environment.NewLine);
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Table 1".PadRight(20) + "Prefix Name 1".PadRight(40) + "Common Prefix".PadRight(15) + "" + "Prefix Name 2".PadRight(40) + "Table Name 2");
          foreach (var dupe in duplicates)
          {
            var result =
              from fnd in StaticVariable.PrefixNumbersRecord
              where fnd.PrefixNumber.Equals(dupe)
              orderby fnd.TableName
              select fnd;                               
            foreach(var num in result)
            {
              if(num.PrefixNumber.Equals(previousNumber))
              {
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + num.TableName.PadRight(20) + num.PrefixName.PadRight(40) + num.PrefixNumber.PadRight(15) + "" + previousName.PadRight(40) + previousTable);
              }
              previousNumber = num.PrefixNumber;
              previousName = num.PrefixName;
              previousTable = num.TableName;
            }               
          }
      }
      Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForDuplicatePrefixNumbers() -- finished");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "CheckForDuplicatePrefixNumbers() -- finished");
    }
    private static void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
    {
      Console.Write("-");
    }
    private static void ReadPrefixesIntoList(string folder, string message)
    {
      Console.WriteLine("Prefixes".PadRight(30, '.') + "ReadPrefixesIntoList( " + message + " ) -- started");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "ReadPrefixesIntoList( " + message + " ) -- started");
      const string iniExtensionSearch = @"*ini";
      string[] folders = Directory.GetFiles(folder, iniExtensionSearch);      
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
                  tableName = StaticVariable.CountryCodeValue + "_" + lines[1];
                }
                if (line.Contains(','))
                {
                  string[] lines = line.Split(',');
                  prefixName = lines[0].Trim();
                  prefixNumber = lines[1].Trim();
                }
                  if (!string.IsNullOrEmpty(tableName) && !string.IsNullOrEmpty(prefixName))
                {
                  StaticVariable.PrefixNumbersFromIniFiles.Add(tableName + "\t" + prefixNumber + "\t" + ValidateData.CapitaliseWord(prefixName));
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
          StaticVariable.ProgressDetails.Add("Prefixes::CombinePrefixes()");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + message + ": Problem adding prefixes to PrefixNumbersDataRecord.");
          StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + e.Message);
          ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
        }
      }
      StaticVariable.PrefixNumbersFromIniFiles.Sort();
      StaticVariable.PrefixNumbersFromIniFiles = StaticVariable.PrefixNumbersFromIniFiles.Distinct().ToList();
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "ReadPrefixesIntoList( " + message + " ) -- finished");
      Console.WriteLine("Prefixes".PadRight(30, '.') + "ReadPrefixesIntoList( " + message + " ) -- finished");
    }       
    public static List<string> MatchPrefixesWithDestinations()
    {
      Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- started");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- started");

      var query =
        from drm in StaticVariable.CustomerDetailsDataRecord
        join pn in StaticVariable.PrefixNumbersRecord on drm.StdPrefixName.ToUpper() equals pn.stdPrefixName.ToUpper()        
        select new { pn.TableName, pn.PrefixName, pn.PrefixNumber, drm.CustomerPrefixName, drm.CustomerUsingCustomerNames, pn.stdPrefixName };

      List<string> prefixesMatched = (from entry in query let prefixName = entry.CustomerUsingCustomerNames.ToUpper().Equals("TRUE") ? entry.CustomerPrefixName : entry.stdPrefixName select entry.TableName + "\t" + entry.PrefixNumber + "\t" + prefixName).ToList();
      prefixesMatched = prefixesMatched.Distinct().ToList();

      Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- finished");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "MatchPrefixesWithDestinations() -- finished");      
      return prefixesMatched;
    }
    public static List<string> GetNationalDomesticPrefixes()
    {
      Console.WriteLine("Prefixes".PadRight(30, '.') + "GetNationalPrefixes() -- started");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "GetNationalPrefixes() -- started"); 
      List<string> nationalPrefixes =  new List<string>();
      const int tableName = 0;
      const int prefix = 1;
      const int prefixName = 2;

      foreach(var name in StaticVariable.PrefixNumbersFromIniFiles)
      {
        string[] nationalOnly = name.Split('\t');
        if(nationalOnly[tableName].ToUpper().Equals(StaticVariable.NationalTableSpellingValue.ToUpper()) && !nationalOnly[prefix].Equals("?"))
        {
            nationalPrefixes.Add(nationalOnly[prefixName] + "," + nationalOnly[prefix]);
        }
      }
      nationalPrefixes.Distinct();
      nationalPrefixes.Sort();
      Console.WriteLine("Prefixes".PadRight(30, '.') + "GetNationalPrefixes() -- finished");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "GetNationalPrefixes() -- finished");
      return nationalPrefixes;
    }
    private static Dictionary<string, string> CountryExceptions()
    {
      Console.WriteLine("Prefixes".PadRight(30, '.') + "CountryExceptions() -- started");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "CountryExceptions() -- started");
      Dictionary<string,string> dict = new Dictionary<string, string>();
      string[] countries = Constants.SpecialCountries.Split(',');
      foreach(var country in countries)
      {
        dict.Add(country.Trim(), country);
      }       
      Console.WriteLine("Prefixes".PadRight(30, '.') + "CountryExceptions() -- finished");
      StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "CountryExceptions() -- finished");
      return dict;
    }
  }
}
