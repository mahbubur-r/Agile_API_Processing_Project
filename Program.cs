using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;

class Program
{
    static async Task Main()
    {
        // Define multiple API endpoints
        string apiUrl1 = "http://13.48.42.106:8000/request-details/"; // 3B
        string apiUrl2 = "https://dg4gi3uw0m2xs.cloudfront.net/agreement"; // 2B
        string apiUrlProviderA = "http://ec2-52-90-1-48.compute-1.amazonaws.com:4000/users/offers?provider=A"; // 4B
        string apiUrlProviderB = "http://ec2-52-90-1-48.compute-1.amazonaws.com:4000/users/offers?provider=B"; // 4B
        string apiUrlProviderC = "http://ec2-52-90-1-48.compute-1.amazonaws.com:4000/users/offers?provider=C"; // 4B
        string apiUrlProviderD = "http://ec2-52-90-1-48.compute-1.amazonaws.com:4000/users/offers?provider=D"; // 4B

        // Define corresponding sheet names
        string sheetName1 = "OpenServices";
        string sheetName2 = "AgreementDetails";
        string sheetNameAgreementBids = "AgreementBids";

        // Call APIs and get data
        var apiData1 = await CallApi(apiUrl1);
        var apiData2 = await CallApi(apiUrl2);
        var apiDataProviderA = await CallApi(apiUrlProviderA);
        var apiDataProviderB = await CallApi(apiUrlProviderB);
        var apiDataProviderC = await CallApi(apiUrlProviderC);
        var apiDataProviderD = await CallApi(apiUrlProviderD);

        // Write data to Excel
        //WriteToExcel(apiData1, sheetName1, "output.xlsx");
        //WriteToExcel(apiData2, sheetName2, "output.xlsx");

        // Combine data from the AgreementBids APIs
        var agreementBidsData = CombineAgreementBids(apiDataProviderA, apiDataProviderB, apiDataProviderC, apiDataProviderD);

        // Write combined AgreementBids data to Excel
        //WriteToExcel(agreementBidsData, sheetNameAgreementBids, "output.xlsx");

        var sheetData = new Dictionary<string, string>
        {
            { sheetName1, apiData1 },
            { sheetName2, apiData2 },
            { sheetNameAgreementBids, agreementBidsData },
        };

        // Write data to Excel
        WriteToExcel(sheetData, "output.xlsx");

        Console.WriteLine("Data written to Excel successfully!");
    }

    static async Task<string> CallApi(string apiUrl)
    {
        using (var httpClient = new HttpClient())
        {
            // Make the API request
            var response = await httpClient.GetStringAsync(apiUrl);
            return response;
        }
    }

    static void WriteToExcel(Dictionary<string, string> sheetData, string filePath)
    {
        // Set the license context
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        // Create a new Excel package
        using (var excelPackage = new ExcelPackage())
        {
            foreach (var kvp in sheetData)
            {
                var sheetName = kvp.Key;
                var jsonData = kvp.Value;

                // Add a worksheet
                var worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);

                // Deserialize JSON data based on the sheet name
                if (sheetName == "OpenServices")
                {
                    var openServices = JsonConvert.DeserializeObject<List<OpenServices>>(jsonData);
                    PopulateExcelSheet(worksheet, openServices);
                }
                else if (sheetName == "AgreementDetails")
                {
                    var agreementDetails = JsonConvert.DeserializeObject<List<AgreementDetails>>(jsonData);
                    PopulateExcelSheet(worksheet, agreementDetails);
                }

                else if (sheetName == "AgreementBids")
                {
                    var agreementBids = JsonConvert.DeserializeObject<List<AgreementBids>>(jsonData);
                    PopulateExcelSheet(worksheet, agreementBids);
                }
            }

            // Save the Excel file
            excelPackage.SaveAs(new FileInfo(filePath));
        }
    }

    static void PopulateExcelSheet<T>(ExcelWorksheet worksheet, List<T> dataList)
    {
        var properties = typeof(T).GetProperties();

        // Start from the first row
        int row = 1;

        // Write headers in the first row
        for (int i = 0; i < properties.Length; i++)
        {
            worksheet.Cells[row, i + 1].Value = properties[i].Name;
        }

        // Increment the row index for data
        row++;

        // Write data in subsequent rows
        foreach (var data in dataList)
        {
            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cells[row, i + 1].Value = properties[i].GetValue(data);
            }

            // Increment the row index for the next set of data
            row++;
        }
    }

    static string CombineAgreementBids(string apiDataA, string apiDataB, string apiDataC, string apiDataD)
    {
        // Deserialize JSON data
        var agreementBidsA = JsonConvert.DeserializeObject<List<AgreementBids>>(apiDataA);
        var agreementBidsB = JsonConvert.DeserializeObject<List<AgreementBids>>(apiDataB);
        var agreementBidsC = JsonConvert.DeserializeObject<List<AgreementBids>>(apiDataC);
        var agreementBidsD = JsonConvert.DeserializeObject<List<AgreementBids>>(apiDataD);

        // Add provider information to each item
        foreach (var item in agreementBidsA)
        {
            item.Provider = "A";
        }

        foreach (var item in agreementBidsB)
        {
            item.Provider = "B";
        }

        foreach (var item in agreementBidsC)
        {
            item.Provider = "C";
        }

        foreach (var item in agreementBidsD)
        {
            item.Provider = "D";
        }

        // Combine all data into a single list
        var combinedList = new List<AgreementBids>();
        combinedList.AddRange(agreementBidsA);
        combinedList.AddRange(agreementBidsB);
        combinedList.AddRange(agreementBidsC);
        combinedList.AddRange(agreementBidsD);

        // Serialize the combined list to JSON
        var combinedJson = JsonConvert.SerializeObject(combinedList);

        return combinedJson;
    }
}
