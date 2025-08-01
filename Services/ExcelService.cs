using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZoningAccelerator.Models;

namespace ZoningAccelerator.Services
{
    public class ExcelService
    {
        private List<T> GetCodesFromSheet<T>(string filePath, string sheetName, string columnName, Func<string, T> mapFunc)
        {
            var result = new List<T>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(sheetName);

            var column = worksheet.Row(1).Cells().FirstOrDefault(c => c.Value.ToString().Trim() == columnName);

            if (column == null)
                throw new Exception($"Column '{columnName}' not found in sheet '{sheetName}'.");

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var cellValue = row.Cell(column.Address.ColumnNumber).GetString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    result.Add(mapFunc(cellValue));
                }
            }

            return result;
        }

        public List<DwellingTypes> GetDwellingCodesFromSheet(string filePath, string sheetName, string columnName)
        {
            return GetCodesFromSheet(filePath, sheetName, columnName, code => new DwellingTypes { DwellingType = code });
        }

        public List<AncillaryTypes> GetAncillaryCodesFromSheet(string filePath, string sheetName, string columnName)
        {
            return GetCodesFromSheet(filePath, sheetName, columnName, code => new AncillaryTypes { AncillaryType = code });
        }

        public List<PermittedUses> GetPermittedUseCodesFromSheet(string filePath, string sheetName, string columnName)
        {
            return GetCodesFromSheet(filePath, sheetName, columnName, code => new PermittedUses { PermittedUse = code });
        }

        public List<TypeOfUses> GetTypeOfUseCodesFromSheet(string filePath, string sheetName, string columnName)
        {
            return GetCodesFromSheet(filePath, sheetName, columnName, code => new TypeOfUses { TypeOfUse = code });
        }
    }
}
