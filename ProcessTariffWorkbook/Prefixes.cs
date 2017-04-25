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
         ReadPrefixesIntoList(StaticVariable.DatasetFolderToUse, "Domestic");
         ReadPrefixesIntoList(StaticVariable.DatasetsFolder, "International"); 
         MatchPrefixNamesWithRegEx(StaticVariable.PrefixNumbersFromIniFiles);
         MatchPrefixNamesToStandardNamesAndAddToPrefixesDataRecord(StaticVariable.PrefixesMatchedByRegEx, StaticVariable.PrefixNumbersFromIniFiles);                 
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
      public static void AddPrefixesToPrefixNumbersRecord(List<string> lst)
        {
            Console.WriteLine("Prefixes".PadRight(30, '.') + "AddPrefixesToPrefixNumbersRecord() -- started");
            StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                             "AddPrefixesToPrefixNumbersRecord() -- started");
            lst = lst.Distinct().ToList();
            foreach (var tok in lst)
            {
                StaticVariable.PrefixNumbersRecord.Add(new PrefixNumbersDataRecord(tok));
            }
            Console.WriteLine("Prefixes".PadRight(30, '.') + "AddPrefixesToPrefixNumbersRecord() -- finished");
            StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                             "AddPrefixesToPrefixNumbersRecord() -- finished");
        }

      private static void MatchPrefixNamesToStandardNamesAndAddToPrefixesDataRecord(List<string> standardPrefixNames, List<string> prefixNumbers)
      {
         Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNumbersToRegExStandardNames() -- started");
         StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                          "MatchPrefixNumbersToRegExStandardNames() -- started");
         const int band = 0;
         const int stdName = 1;
         const int prefixName = 2;

         prefixNumbers.Sort();
         standardPrefixNames.Sort();

         foreach (var spn in standardPrefixNames)
         {
               string[] stdNames = spn.Split('\t');
               foreach (var pn in prefixNumbers)
               {
                  string[] prefixes = pn.Split('\t');
                  if (stdNames[prefixName].ToUpper().Equals(prefixes[prefixName].ToUpper()))
                  {
                     StaticVariable.PrefixNumbersRecord.Add(
                           new PrefixNumbersDataRecord(pn + "\t" + stdNames[band] + "\t" + stdNames[stdName]));
                  }
               }
         }
         Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNumbersToRegExStandardNames() -- finished");
         StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                          "MatchPrefixNumbersToRegExStandardNames() -- finished");
      }      

      private static void CheckForDestinationsWithoutPrefixes()
      {
         Console.WriteLine("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- started");
         StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- started");

         var queryMissingPrefixes =
               from drm in StaticVariable.CustomerDetailsDataRecord
               orderby drm.StdPrefixName
               select drm.StdPrefixName.ToUpper();

         var queryPrefixNames =
         (from pn in StaticVariable.PrefixNumbersRecord
               orderby pn.stdPrefixName
               select pn.stdPrefixName.ToUpper()).Distinct();

         //var missingPrefixes2 = queryPrefixNames.Except(queryMissingPrefixes).ToList();
         var missingPrefixes = queryMissingPrefixes.Except(queryPrefixNames).ToList();

         //missingPrefixes2.Sort();
         missingPrefixes.Sort();

         if (missingPrefixes.Any())
         {
               StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::DestinationsWithoutPrefixes()");
               StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "No Prefix Found:");
               StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "The std  or customer prefix name may not match the name in the appropriate prefix table or else the prefix may not exist in that table.");
               StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Destinations in the National table may be included in the list, even though they exist in the prefix table.");

               foreach (var entry in missingPrefixes)
               {
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + ValidateData.CapitaliseWord(entry));
               }
         }
         Console.WriteLine("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- finished");
         StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "DestinationsWithoutPrefixes() -- finished");
      }

      private static void CheckForNonMatchingPrefixNames()
      {
         Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForNonMatchingPrefixNames() -- started");
         StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                          "CheckForNonMatchingPrefixNames() -- started");
         List<string> tmpList = new List<string>();

         var queryDifferentNames =
         (from pn in StaticVariable.PrefixNumbersRecord
               join cd in StaticVariable.CustomerDetailsDataRecord on pn.stdPrefixName.ToUpper() equals
               cd.StdPrefixName.ToUpper()
               where
               !pn.PrefixName.ToUpper().Equals(cd.StdPrefixName.ToUpper()) &&
               !pn.PrefixName.ToUpper().Equals(cd.CustomerPrefixName.ToUpper())
               select pn.PrefixName).Distinct();

         if (queryDifferentNames.Any())
         {
               StaticVariable.ProgressDetails.Add(Environment.NewLine +
                                                "Prefixes::CheckForNonMatchingPrefixNames()");
               StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                "Prefix names that do not match the standard name nor the name in the input xlsx sheet.");
               foreach (var names in queryDifferentNames)
               {
                  StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + " Prefix file name : " + names);
               }
               ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
         }
         Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForNonMatchingPrefixNames() -- finished");
         StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                          "CheckForNonMatchingPrefixNames() -- finished");
      }
      private static void MatchPrefixNamesWithRegEx(List<string> listToUse)
        {
            Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNamesWithRegEx() -- started");
            StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "MatchPrefixNamesWithRegEx() -- started");
            string destinationName = string.Empty;
            var prefixNamesNotMatchedByRegex = new List<string>();
            var dupeBandsForSamePrefixName = new List<string>();
            var prefixNamesMatchedMoreThanOnce = new List<string>();
            List<string> prefixNamesOnly = new List<string>();
            Dictionary<string, string> preixNamesMatchedWithDuplicateBands = new Dictionary<string, string>();
            Dictionary<string, string> prefixNamesMatchedMultipleTimes = new Dictionary<string, string>();
            const int tableElement = 0;
            const int prefixNameElement = 2;
            Timer newTimer = new System.Timers.Timer(10000); // 2 sec interval
            string regExAlphanumeric = @"[0-9|a-z|A-Z]"; //@"\w|\s"
            string regExExtraSpaces = "\x0020{2,}";
            string regExNull = "\x0000";
            string regExNoise = @"\(|\)|( -|- )|\,|'";
            //string regEx_Noise = @"\,|'";
            Regex regExRemoveNull = new Regex(regExNull, RegexOptions.Multiline);
            Regex regExRemoveNoise = new Regex(regExNoise, RegexOptions.Compiled);
            Regex regExRemoveExtraSpaces = new Regex(regExExtraSpaces, RegexOptions.Compiled);
            Regex regExCheckForAlphanumeric = new Regex(regExAlphanumeric, RegexOptions.Compiled);

            foreach (var names in listToUse)
            {
                string[] name = names.Split('\t');
                if (name[tableElement].ToUpper().Equals(StaticVariable.NationalTableSpellingValue.ToUpper())) continue;
                prefixNamesOnly.Add(name[prefixNameElement]);
            }
            prefixNamesOnly = prefixNamesOnly.Distinct().ToList();
            prefixNamesOnly.Sort();

            foreach (string prefixName in prefixNamesOnly)
            {
                var found = false;
                try
                {
                    destinationName = prefixName.Trim();
                    destinationName = regExRemoveNoise.Replace(destinationName, " ");
                    destinationName = regExRemoveExtraSpaces.Replace(destinationName, " ");
                }
                catch (Exception e)
                {
                    StaticVariable.ProgressDetails.Add("Prefixes::MatchPrefixNamesWithRegEx()");
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                       "There may be a problem with the input destination");
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
                        if (regExCountry.IsMatch(destinationName))
                        {
                            newTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
                            newTimer.AutoReset = true;
                            newTimer.Enabled = true;
                            found = true;
                            try
                            {
                                preixNamesMatchedWithDuplicateBands.Add(regExBand, prefixName + "\t" + regExBand + "\t" + regexStandardName);
                            }
                            catch (Exception )
                            {
                                dupeBandsForSamePrefixName.Add("First Prefix Name : " + preixNamesMatchedWithDuplicateBands[regExBand] + " , Second Prefix Name : " + prefixName);
                            }
                            try
                            {
                                prefixNamesMatchedMultipleTimes.Add(prefixName, prefixName + "\t" + regExBand + "\t" + regexStandardName);
                            }
                            catch (Exception )
                            {
                                prefixNamesMatchedMoreThanOnce.Add(prefixNamesMatchedMultipleTimes[prefixName] + "  Other Band & Standard Prefix Name - " + regExBand + "\t" + regexStandardName);
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
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + "Band assigned to more than one to Prefix Name");
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
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                   "Prefix Name assigned to more than one band");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                   "E.g. If Inmarsat B. Add new entry to input Xlsx file and change to 'Inmarsat B Land' & 'Inmarsat B Maritime'. ");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                   "If prefixes were supplied you may need to delete these bands in RegEx file and add 'Inmarsat B' regex");
                foreach (var item in prefixNamesMatchedMoreThanOnce)
                {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + item);
                }
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            if (prefixNamesNotMatchedByRegex.Any())
            {
                StaticVariable.ProgressDetails.Add(Environment.NewLine + "Prefixes::MatchPrefixNamesWithRegEx()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                   "Prefixes without regex. Change the prefix name to match the Customer name");
                prefixNamesNotMatchedByRegex.Sort();
                foreach (string entry in prefixNamesNotMatchedByRegex)
                {
                    StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding + entry);
                }
                ErrorProcessing.StopProcessDueToFatalErrorOutputToLog();
            }
            newTimer.Enabled = false;
            Console.WriteLine("Prefixes".PadRight(30, '.') + "MatchPrefixNamesWithRegEx() -- finished");
            StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                             "MatchPrefixNamesWithRegEx() -- finished");
        }
      public static void CheckForDuplicatePrefixNumbers()
        {
            Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForDuplicatePrefixNumbers() -- started");
            StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') + "CheckForDuplicatePrefixNumbers() -- started");
            Dictionary<string, string> dict = new Dictionary<string, string>();
            List<string> duplicates = new List<string>();

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
                StaticVariable.ProgressDetails.Add(Environment.NewLine +
                                                   "Prefixes::CheckForDuplicatePrefixNumbers()");
                StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                   "Some prefixes amy be in more than one file. Delete ini files not required.");
                foreach (var dupe in duplicates)
                {
                    var result =
                        from fnd in StaticVariable.PrefixNumbersRecord
                        where fnd.PrefixNumber.Equals(dupe)
                        select fnd;
                    foreach (var num in result)
                    {
                        StaticVariable.ProgressDetails.Add(Constants.FiveSpacesPadding +
                                                           num.PrefixNumber.PadRight(15, ' ') + " : " +
                                                           num.PrefixName.PadRight(40, ' ') + " : " + num.TableName);
                    }
                }
            }
            Console.WriteLine("Prefixes".PadRight(30, '.') + "CheckForDuplicatePrefixNumbers() -- finished");
            StaticVariable.ConsoleOutput.Add("Prefixes".PadRight(30, '.') +
                                             "CheckForDuplicatePrefixNumbers() -- finished");
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
                  prefixName = lines[0];
                  prefixNumber = lines[1];
                }
                if (!string.IsNullOrEmpty(tableName) && !string.IsNullOrEmpty(prefixName))
                {
                  StaticVariable.PrefixNumbersFromIniFiles.Add(tableName + "\t" + prefixNumber + "\t" + prefixName);
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
   }
}
