using ZoningAccelerator.Helpers;
using ZoningAccelerator.Models;
using ZoningAccelerator.Services;

namespace ZoningAccelerator.Commands
{
    public class DwellingTypeComparison
    {
        private readonly ExcelService _excelService;

        public DwellingTypeComparison()
        {
            _excelService = new ExcelService();
        }

        public void Execute(string masterFile, string cityFile, string citySheet)
        {

            var masterCodes = _excelService.GetDwellingCodesFromSheet(masterFile, "MD - Dwelling Types", "Code")
                                                         .Select(dt => dt.DwellingType)
                                                         .ToHashSet();

            var cityCodes = _excelService.GetDwellingCodesFromSheet(cityFile, citySheet, "DwellingType")
                                                            .Select(dt => dt.DwellingType)
                                                            .ToList();

            var newDwellings = cityCodes.Where(DwellingType => !masterCodes.Contains(DwellingType)).Distinct().ToList();

            if (newDwellings.Any())
            {
                var outputPath = PathHelper.GetUniqueFilePath("UniqueDwellingTypes", cityFile, ".txt");
                FileHelper.WriteToFile(newDwellings, outputPath);
                Console.WriteLine($"Found {newDwellings.Count} new dwelling types. Details written to {outputPath}.");
            }
            else
            {
                Console.WriteLine("No new dwelling types found.");
            }
        }
    }
}
