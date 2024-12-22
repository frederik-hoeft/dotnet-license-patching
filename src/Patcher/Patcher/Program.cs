using GemBox.Spreadsheet;
using Patcher;

Console.WriteLine("Attempting to load data.csv with unpatched and unlicensed GemBox.Spreadsheet...");
try
{
    _ = ExcelFile.Load("data.csv");
}
catch (Exception e)
{
    Console.WriteLine("Exception caught:");
    Console.WriteLine($"Unable to load data.csv: {e.Message}");
    Console.WriteLine("This is expected, since we don't have a license for GemBox.Spreadsheet.");
}
Console.WriteLine("Attempting to load data.csv with patched GemBox.Spreadsheet...");
GemboxLicensePatches.ApplyAll();
ExcelFile file = ExcelFile.Load("data.csv");
ExcelWorksheet worksheet = file.Worksheets[0];
Console.WriteLine("Here's some data from the first worksheet:");
Console.Write(worksheet.Rows[^1].Cells[1].Value);
Console.Write(": ");
Console.WriteLine(worksheet.Rows[^1].Cells[2].Value);
Console.WriteLine("Even though the assembly was obfuscated, we were able to find the patch target using reflection paths via known signatures and entry points!");