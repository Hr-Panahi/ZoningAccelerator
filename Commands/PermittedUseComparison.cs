using System.Runtime.Intrinsics.X86;
using ZoningAccelerator.Helpers;
using ZoningAccelerator.Models;
using ZoningAccelerator.Services;

namespace ZoningAccelerator.Commands
{
    public class PermittedUseComparison
    {
        private readonly ExcelService _excelService;

        public PermittedUseComparison()
        {
            _excelService = new ExcelService();
        }

        public void Execute(string filePath, string sheet1, string column1, string sheet2, string column2)
        {
            var masterCodes = _excelService.GetPermittedUseCodesFromSheet(filePath, sheet2, column2)
                                           .Select(dt => dt.PermittedUse)
                                           .ToHashSet();

            var cityCodes = _excelService.GetPermittedUseCodesFromSheet(filePath, sheet1, column1)
                                         .Select(dt => dt.PermittedUse)
                                         .ToList();

            var newPermittedUses = cityCodes.Where(code => !masterCodes.Contains(code)).Distinct().ToList();

            if (newPermittedUses.Any())
            {
                var outputPath = PathHelper.GetUniqueFilePath( "UniquePermittedUses", filePath, ".txt");
                FileHelper.WriteToFile(newPermittedUses, outputPath);
                Console.WriteLine($"Found {newPermittedUses.Count} new permitted uses. Details written to {outputPath}.");
            }
            else
            {
                Console.WriteLine("No new permitted uses found.");
            }
        }
    }
}
