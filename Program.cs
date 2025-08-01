using ZoningAccelerator.Commands;
using ZoningAccelerator.Helpers;
using ZoningAccelerator.Services;
using System;
using System.IO;
using System.Linq;

namespace ZoningAccelerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            Console.Title = "Zoning Accelerator CLI";

            while (true)
            {
                ShowMainMenu();

                var choice = Console.ReadLine()?.Trim().ToLower();
                Console.WriteLine();

                if (choice == "q" || choice == "exit")
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("👋 Exiting. Goodbye!");
                    Console.ResetColor();
                    break;
                }

                try
                {
                    switch (choice)
                    {
                        case "1":
                            RunDwellingTypeComparison();
                            break;
                        case "2":
                            RunAncillaryTypeComparison();
                            break;
                        case "3":
                            RunPermittedUseComparison();
                            break;
                        case "4":
                            RunTypeOfUseComparison();
                            break;
                        case "5":
                            RunAllComparisons();
                            break;
                        case "6":
                            GetUniqueDwellings();
                            break;
                        case "7":
                            GetUniqueAncillaries();
                            break;
                        case "8":
                            GetUniquePermittedUses();
                            break;
                        case "9":
                            GetUniqueTypeOfUses();
                            break;
                        default:
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Invalid choice. Please try again.");
                            Console.ResetColor();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"\n❌ An unexpected error occurred: {ex.Message}");
                    Console.ResetColor();
                }

                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("Press Enter to return to main menu...");
                Console.ResetColor();
                Console.ReadLine();
                Console.Clear();
            }
        }

        static void ShowMainMenu()
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("========== ZONING ACCELERATOR ==========\n");

            string[] leftOptions = {
        "1. Compare DwellingTypes",
        "2. Compare AncillaryTypes",
        "3. Compare PermittedUses",
        "4. Compare TypeOfUses",
        "5. Run All Comparisons"
    };

            string[] rightOptions = {
        "6. Get Unique Dwellings",
        "7. Get Unique Ancillaries",
        "8. Get Unique Permitted Uses",
        "9. Get Unique Type Of Uses"
    };

            int leftColumnWidth = 25;          // tighter left column width
            int rightStartPos = leftColumnWidth + 5;  // right column 5 spaces after left

            int maxLines = Math.Max(leftOptions.Length, rightOptions.Length);

            for (int i = 0; i < maxLines; i++)
            {
                string leftText = i < leftOptions.Length ? leftOptions[i] : "";
                string rightText = i < rightOptions.Length ? rightOptions[i] : "";

                // Write left option padded right to fixed width
                Console.Write(leftText.PadRight(leftColumnWidth));

                // Write right option starting at rightStartPos
                if (!string.IsNullOrEmpty(rightText))
                {
                    Console.SetCursorPosition(rightStartPos, Console.CursorTop);
                    Console.Write(rightText);
                }

                Console.WriteLine();
            }

            Console.WriteLine("  q. Quit");
            Console.Write("\nChoice: ");
            Console.ResetColor();
        }

        static void ShowSuccess(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\n✅ {message}");
            Console.ResetColor();
        }

        static void ShowError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n❌ {message}");
            Console.ResetColor();
        }

        static string ReadInput(string prompt)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write(prompt);
            Console.ResetColor();
            return Console.ReadLine()?.Trim().Trim('"');
        }

        static void RunDwellingTypeComparison()
        {
            var masterPath = ReadInput("Enter path to Master Excel (Zoning Master Data): ");
            var cityPath = ReadInput("Enter path to City Excel (Zoning Regulation): ");
            const string citySheet = "Zone Allowed DwellingTypes";

            var comparer = new DwellingTypeComparison();
            comparer.Execute(masterPath, cityPath, citySheet);

            ShowSuccess("DwellingTypes comparison complete.");
        }

        static void RunAncillaryTypeComparison()
        {
            var masterPath = ReadInput("Enter path to Master Excel (Zoning Master Data): ");
            var cityPath = ReadInput("Enter path to City Excel: ");
            const string citySheet = "Zone Allowed AncillaryTypes";

            var comparer = new AncillaryTypeComparison();
            comparer.Execute(masterPath, cityPath, citySheet);

            ShowSuccess("AncillaryTypes comparison complete.");
        }

        static void RunTypeOfUseComparison()
        {
            var masterPath = ReadInput("Enter path to 'Zoning Master Data' excel: ");
            var cityPath = ReadInput("Enter path to Regulation excel: ");
            const string citySheet = "Zone Permitted Uses";

            var comparer = new TypeOfUseComparison();
            comparer.Execute(masterPath, cityPath, citySheet);

            ShowSuccess("TypeOfUses comparison complete.");
        }

        static void RunPermittedUseComparison()
        {
            var filePath = ReadInput("Enter path to Excel file containing both sheets: ");

            var comparer = new PermittedUseComparison();
            comparer.Execute(filePath,
                             "Zone Permitted Uses", "Permitted Use",
                             "MD - Permitted Uses", "Code");

            ShowSuccess("PermittedUses comparison complete.");
        }

        static void RunAllComparisons()
        {
            var masterPath = ReadInput("Enter path to Master Excel (Zoning Master Data): ");
            var cityPath = ReadInput("Enter path to City Excel (Zoning Regulation): ");
            var excelService = new ExcelService();

            Console.WriteLine("\nComparing data...");

            // DwellingTypes
            var masterDwelling = excelService.GetDwellingCodesFromSheet(masterPath, "MD - Dwelling Types", "Code")
                                             .Select(c => c.DwellingType).ToHashSet();

            var cityDwelling = excelService.GetDwellingCodesFromSheet(cityPath, "Zone Allowed DwellingTypes", "DwellingType")
                                           .Select(c => c.DwellingType).ToList();
            var newDwelling = cityDwelling.Where(c => !masterDwelling.Contains(c)).Distinct().ToList();

            // AncillaryTypes
            var masterAncillary = excelService.GetAncillaryCodesFromSheet(masterPath, "MD - Ancillary Types", "Code")
                                              .Select(c => c.AncillaryType).ToHashSet();

            var cityAncillary = excelService.GetAncillaryCodesFromSheet(cityPath, "Zone Allowed AncillaryTypes", "AncillaryType")
                                            .Select(c => c.AncillaryType).ToList();
            var newAncillary = cityAncillary.Where(c => !masterAncillary.Contains(c)).Distinct().ToList();
            // TypeOfUses
            var masterTypeOfUse = excelService.GetTypeOfUseCodesFromSheet(masterPath, "MD - Type Of Use", "Code")
                                              .Select(c => c.TypeOfUse).ToHashSet();
            var cityTypeOfUse = excelService.GetTypeOfUseCodesFromSheet(cityPath, "Zone Permitted Uses", "TypeOfUse")
                                            .Select(c => c.TypeOfUse).ToList();
            var newTypeOfUse = cityTypeOfUse.Where(c => !masterTypeOfUse.Contains(c)).Distinct().ToList();

            // PermittedUses
            var masterPermitted = excelService.GetPermittedUseCodesFromSheet(cityPath, "MD - Permitted Uses", "Code")
                                              .Select(p => p.PermittedUse).ToHashSet();

            var cityPermitted = excelService.GetPermittedUseCodesFromSheet(cityPath, "Zone Permitted Uses", "Permitted Use")
                                            .Select(p => p.PermittedUse).ToList();
            var newPermitted = cityPermitted.Where(p => !masterPermitted.Contains(p)).Distinct().ToList();

            // Write to Excel
            var outputPath = PathHelper.GetUniqueFilePath("AllComparisonResults", cityPath, ".xlsx");
            FileHelper.WriteToExcel(outputPath, newDwelling, newAncillary, newPermitted, newTypeOfUse);

            ShowSuccess($"All comparisons complete.\n📁 Results saved to: {outputPath}");
        }

        static void GetUniqueDwellings()
        {
            var cityPath = ReadInput("Enter path to City Excel (Zoning Regulation): ");
            var sheetName = "Zone Allowed DwellingTypes";
            var columnName = "DwellingType";

            var excelService = new ExcelService();

            try
            {
                var dwellings = excelService.GetDwellingCodesFromSheet(cityPath, sheetName, columnName)
                                           .Select(d => d.DwellingType.Trim())
                                           .Distinct()
                                           .OrderBy(d => d)
                                           .ToList();

                var outputFilePath = PathHelper.GetUniqueFilePath("UniqueDwellings", cityPath, ".txt");
                FileHelper.WriteToFile(dwellings,outputFilePath);

                ShowSuccess($"New Unique DwellingTypess saved to:\n{outputFilePath}");
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        static void GetUniqueAncillaries()
        {
            var cityPath = ReadInput("Enter path to City Excel (Zoning Regulation): ");
            var sheetName = "Zone Allowed AncillaryTypes";
            var columnName = "AncillaryType";

            var excelService = new ExcelService();

            try
            {
                var ancillaries = excelService.GetAncillaryCodesFromSheet(cityPath, sheetName, columnName)
                                             .Select(a => a.AncillaryType.Trim())
                                             .Distinct()
                                             .OrderBy(a => a)
                                             .ToList();

                var outputFilePath = PathHelper.GetUniqueFilePath("UniqueAncillaries", cityPath, ".txt");
                FileHelper.WriteToFile(ancillaries, outputFilePath);

                ShowSuccess($"New Unique AncillaryTypes saved to:\n{outputFilePath}");
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        static void GetUniqueTypeOfUses()
        {
            var cityPath = ReadInput("Enter path to City Excel (Zoning Regulation): ");
            var sheetName = "Zone Permitted Uses";
            var columnName = "TypeOfUse";

            var excelService = new ExcelService();

            try
            {
                var typeOfUses = excelService.GetTypeOfUseCodesFromSheet(cityPath, sheetName, columnName)
                                             .Select(a => a.TypeOfUse.Trim())
                                             .Distinct()
                                             .OrderBy(a => a)
                                             .ToList();

                var outputFilePath = PathHelper.GetUniqueFilePath("UniqueTypeOfUses", cityPath, ".txt");
                FileHelper.WriteToFile(typeOfUses, outputFilePath);

                ShowSuccess($"New Unique TypeOfUses saved to:\n{outputFilePath}");
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        static void GetUniquePermittedUses()
        {
            var cityPath = ReadInput("Enter path to City Excel (Zoning Regulation): ");
            var sheetName = "Zone Permitted Uses";
            var columnName = "Permitted Use";

            var excelService = new ExcelService();

            try
            {
                var permittedUses = excelService.GetPermittedUseCodesFromSheet(cityPath, sheetName, columnName)
                                               .Select(p => p.PermittedUse?.Trim())
                                               .Where(p => !string.IsNullOrEmpty(p))
                                               .Distinct()
                                               .OrderBy(p => p)
                                               .ToList();

                var outputFilePath = PathHelper.GetUniqueFilePath("UniquePermittedUses", cityPath, ".txt");
                FileHelper.WriteToFile(permittedUses, outputFilePath);

                ShowSuccess($"New Unique PermittedUses saved to:\n{outputFilePath}");
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }
    }
}
