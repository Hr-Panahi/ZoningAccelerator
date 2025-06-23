using ZoningAccelerator.Helpers;
using ZoningAccelerator.Models;
using ZoningAccelerator.Services;

namespace ZoningAccelerator.Commands
{
    public class AncillaryTypeComparison
    {
        private readonly ExcelService _excelService;

        public AncillaryTypeComparison()
        {
            _excelService = new ExcelService();
        }

        public void Execute(string masterFile, string cityFile, string citySheet)
        {

            var masterCodes = _excelService.GetAncillaryCodesFromSheet(masterFile, "MD - Ancillary Types", "Code")
                                                         .Select(dt => dt.AncillaryType)
                                                         .ToHashSet();

            var cityCodes = _excelService.GetAncillaryCodesFromSheet(cityFile, citySheet, "AncillaryType")
                                                            .Select(dt => dt.AncillaryType)
                                                            .ToList();

            var newAncillaries = cityCodes.Where(AncillaryType => !masterCodes.Contains(AncillaryType)).Distinct().ToList();

            if(newAncillaries.Any())
            {
                var outputPath = Path.Combine(Environment.CurrentDirectory, "UniqueAncillaryTypes.txt");
                FileHelper.WriteToFile(newAncillaries, outputPath);
                Console.WriteLine($"Found {newAncillaries.Count} new ancillary types. Details written to {outputPath}.");
            }
            else
            {
                Console.WriteLine("No new ancillary types found.");
            }
        }
    }
}
