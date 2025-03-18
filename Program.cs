// See https://aka.ms/new-console-template for more information

using System.Text.Json;
using QuickType;
using System;
using System.IO;
using System.Xml.Serialization;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection.PortableExecutable;



var leadWebsites = new string[] { "linkedin", "facebook", "google" };
var processedFiles = new List<FileInfo>();
var unprocessedFiles = new List<FileInfo>();
var lstLeads = new List<Lead>();

string ConfigFile = "PathConfig.txt";
string userBrowsedFolderPath = "";


DateTime date = DateTime.Now;
var dateString = date.ToString("dd_MM_yyyy_HH_mm_ss");

if (ReadConfigPath())
{
    Console.WriteLine();
    Console.WriteLine("Process Started");

    MainFunction();

    Console.WriteLine();
    Console.WriteLine("Process ends");

    //Console.ReadLine();
}

void MoveJsonFileToCompletedFolder()
{
    string compDirPath = userBrowsedFolderPath + "completed_" + dateString;

    DirectoryInfo compDir = new DirectoryInfo(compDirPath);

    if (!compDir.Exists)
    {
        Directory.CreateDirectory(compDirPath);
    }

    foreach (var processedFile in processedFiles)
    {
        File.Move(processedFile.FullName, compDirPath + "\\" + processedFile.Name);
    }

}

void MoveUnProcessedJsonFile()
{
    string unProcessedDirPath = userBrowsedFolderPath + "unProcessed";

    DirectoryInfo compDir = new DirectoryInfo(unProcessedDirPath);

    if (!compDir.Exists)
    {
        Directory.CreateDirectory(unProcessedDirPath);
    }

    foreach (var unprocessedFile in unprocessedFiles)
    {
        File.Move(unprocessedFile.FullName, unProcessedDirPath + "\\" + unprocessedFile.Name);
    }

}

bool IsLeadWebsite(string filename)
{
    foreach (var leadWebsite in leadWebsites)
    {
        if (filename.ToLower().Trim().Contains(leadWebsite.ToLower().Trim()))
        {
            return true;
        }
    }
    return false;
}

bool isNOTNullorEmpty(string str)
{
    return !string.IsNullOrEmpty(str);
}

bool ReadConfigPath()
{
    try
    {

        string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string strDir = System.IO.Path.GetDirectoryName(strExeFilePath);

        string FilePath = System.IO.Path.Combine(strDir, ConfigFile);

        Console.WriteLine(FilePath);
        
        if (File.Exists(FilePath))
        {
            using (StreamReader strTxtRead = new StreamReader(FilePath))
            {
                var firstLine = strTxtRead.ReadLine();
                if (!string.IsNullOrEmpty(firstLine))
                {
                    userBrowsedFolderPath = firstLine.Trim();
                    if (!userBrowsedFolderPath.EndsWith("\\"))
                    {
                        userBrowsedFolderPath = userBrowsedFolderPath + "\\";
                    }
                }
                else
                {
                    Console.WriteLine("Path not found in " + ConfigFile);
                    Console.ReadLine();
                }
            }
        }
        else
        {
            Console.WriteLine(FilePath +" is not found in the Application directory.");
            Console.ReadLine();
        }

    }
    catch (Exception e)
    {
        Console.WriteLine(e);
        return false;
    }

    return true;

}

void MainFunction()
{
    DirectoryInfo dir = new DirectoryInfo(userBrowsedFolderPath);

    if (dir.Exists)
    {
        var jsonFiles = dir.GetFiles().Where(ss => ss.Extension.Equals(".json") && IsLeadWebsite(ss.Name));

        //Console.WriteLine($"Number of JSON Files are {jsonFiles.Count()}");

        if (jsonFiles.Count() < 15)
        {
            return;
        }

        foreach (FileInfo flInfo in jsonFiles)
        {
            try
            {

                ReadJsonFile(flInfo.FullName, flInfo.CreationTime);
                processedFiles.Add(flInfo);
                //Console.WriteLine($"{processedFiles.Count()}");
            }
            catch (Newtonsoft.Json.JsonReaderException jsonEx)
            {
                unprocessedFiles.Add(flInfo);
                Console.WriteLine(jsonEx);
            }
            catch (Exception generalEx)
            {
                unprocessedFiles.Add(flInfo);
                Console.WriteLine(generalEx);
            }

        }

        if (processedFiles.Count > 0)
        {
            WriteToNewCsv();

            WriteToNewExcel();

            MoveJsonFileToCompletedFolder();
        }

        if (unprocessedFiles.Count > 0)
        {

            MoveUnProcessedJsonFile();
        }


    }
}


//Console.ReadLine();


void WriteToNewCsv()
{
    try
    {


        var file = userBrowsedFolderPath + dateString + ".csv";

        using (var stream = File.CreateText(file))
        {

            string csvHeaderRow = string.Format("{0},{1},{2},{3},{4},{5},{6}", "Name", "Profile", "Website", "Phone", "Email", "Location", "Source");
            stream.WriteLine(csvHeaderRow);

            // Loop through your variables and write them to CSV file
            foreach (var leadddd in lstLeads)
            {
                string csvRow = string.Format("{0},{1},{2},{3},{4},{5},{6}", leadddd.Name, leadddd.Profile, leadddd.Website, leadddd.Phone, leadddd.Email, leadddd.Location, leadddd.Source);

                stream.WriteLine(csvRow);
            }
        }

    }
    catch (Exception generalEx)
    {
        Console.WriteLine(generalEx);

    }

}

void WriteToNewExcel()
{
    try
    {

        // Create a new instance of Excel application
        var excelApp = new Microsoft.Office.Interop.Excel.Application();

        // Open an existing workbook or create a new one
        var workbook = excelApp.Workbooks.Add();

        // Get the active worksheet
        var worksheets = workbook.Worksheets.Add();

        Microsoft.Office.Interop.Excel.Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

        // Initialize row counter
        // Header
        int row = 1;
        worksheet.Cells[row, "A"] = "Name";
        worksheet.Cells[row, "B"] = "Profile";
        worksheet.Cells[row, "C"] = "Website";
        worksheet.Cells[row, "D"] = "Phone";
        worksheet.Cells[row, "E"] = "Email";
        worksheet.Cells[row, "F"] = "Location";
        worksheet.Cells[row, "G"] = "Source";

        // ROW
        row = 2;

        // Loop through your variables and write them to Excel
        foreach (var leadddd in lstLeads)
        {
            if (isNOTNullorEmpty(leadddd.Phone) || isNOTNullorEmpty(leadddd.Email))
            {
                worksheet.Cells[row, "A"] = leadddd.Name;
                worksheet.Cells[row, "B"] = leadddd.Profile;
                worksheet.Cells[row, "C"] = leadddd.Website;
                worksheet.Cells[row, "D"] = "'" + leadddd.Phone;
                worksheet.Cells[row, "E"] = leadddd.Email;
                worksheet.Cells[row, "F"] = leadddd.Location;
                worksheet.Cells[row, "G"] = leadddd.Source;
                row++;
            }

        }

        var newFilename = userBrowsedFolderPath + dateString + ".xlsx";

        // Save the workbook
        workbook.SaveAs(newFilename);

        // Close Excel and release resources
        workbook.Close();
        excelApp.Quit();

    }
    catch (Exception e)
    {
        Console.WriteLine(e);
        throw;
    }




}

void ReadJsonFile(string filePath, DateTime creationTime)
{
    using (StreamReader strRead = new StreamReader(filePath))
    {

        string jsonString = strRead.ReadToEnd();

        // use below syntax to access JSON file
        var jsonFile = Lead.FromJson(jsonString);

        //var profileee = jsonFile.Profile;
        //var websitee = jsonFile.Website;
        //var namee = jsonFile.Name;
        //var phonee = jsonFile.Phone;
        //var emailee = jsonFile.Email;
        //var locationnn = jsonFile.Location;


        var leadd = new Lead();
        leadd.Name = jsonFile.Name;
        leadd.Profile = jsonFile.Profile;
        leadd.Website = jsonFile.Website;
        leadd.Phone = jsonFile.Phone;
        leadd.Email = jsonFile.Email;
        leadd.Location = jsonFile.Location;
        leadd.Source = jsonFile.Source;

        lstLeads.Add(leadd);

        //Console.WriteLine("=================================================================");

        //Console.WriteLine(namee);//Name
        //Console.WriteLine(profileee);//Profile
        //Console.WriteLine(websitee);//Website
        //Console.WriteLine(phonee);//Phone
        //Console.WriteLine(emailee);//Email
        //Console.WriteLine(locationnn);//Location


    }

}

//reference https://app.quicktype.io/
namespace QuickType
{

    using System;
    using System.Collections.Generic;

    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public partial class Lead
    {
        [JsonProperty("Profile")] public string Profile { get; set; }

        [JsonProperty("Website")] public string Website { get; set; }

        [JsonProperty("Name")] public string Name { get; set; }

        [JsonProperty("Phone")] public string Phone { get; set; }

        [JsonProperty("Email")] public string Email { get; set; }

        [JsonProperty("Location")] public string Location { get; set; }

        [JsonProperty("Source")] public string Source { get; set; }
    }

    public partial class Lead
    {
        public static Lead FromJson(string json) =>
            JsonConvert.DeserializeObject<Lead>(json, QuickType.Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this Lead self) =>
            JsonConvert.SerializeObject(self, QuickType.Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }

}





