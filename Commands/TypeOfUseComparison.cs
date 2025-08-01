using ZoningAccelerator.Helpers;
using ZoningAccelerator.Models;
using ZoningAccelerator.Services;

namespace ZoningAccelerator.Commands
{
    public class TypeOfUseComparison
    {
        private readonly ExcelService _excelService;

        public TypeOfUseComparison()
        {
            _excelService = new ExcelService();
        }

        public void Execute(string masterFile, string cityFile, string citySheet)
        {

            var masterCodes = _excelService.GetTypeOfUseCodesFromSheet(masterFile, "MD - Type of Use", "Code")
                                                         .Select(dt => dt.TypeOfUse)
                                                         .ToHashSet();

            var cityCodes = _excelService.GetTypeOfUseCodesFromSheet(cityFile, citySheet, "TypeOfUse")
                                                            .Select(dt => dt.TypeOfUse)
                                                            .ToList();

            var newTypeOfUses = cityCodes.Where(TypeOfUse => !masterCodes.Contains(TypeOfUse)).Distinct().ToList();

            if (newTypeOfUses.Any())
            {
                var outputPath = PathHelper.GetUniqueFilePath("UniqueTypeOfUses", cityFile, ".txt");
                FileHelper.WriteToFile(newTypeOfUses, outputPath);
                Console.WriteLine($"Found {newTypeOfUses.Count} new TypeOfUses. Details written to {outputPath}.");
            }
            else
            {
                Console.WriteLine("No new TypeOfUses found.");
            }

        }
    }
}
