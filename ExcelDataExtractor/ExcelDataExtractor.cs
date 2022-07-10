using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace ExcelDataExtractor
{
    public static class Excel
    {
        public static IEnumerable<T> ReadFromExcel<T>(IFormFile excelFile, string[]? requiredHeaders, string[]? nullableFields, string[]? ignoreFields, string duplicateComparer) where T : new()
        {
            List<Dictionary<string, string>> excelData = ExcelDataExtractorExtension.ExtractFromExcel(excelFile, nullableFields, ignoreFields);

            if (requiredHeaders != null)
            {
                excelData.ValidateFields(requiredHeaders);
            }
            
            excelData.CheckDuplicate(duplicateComparer);

            IEnumerable<T> recordsToUpload = DictionaryToObjectConverterExtension.DictionaryToObjects<T>(excelData);
            return recordsToUpload;
        }

        public static MemoryStream WriteToExcel<T>(this IList<T> collection, string worksheetName, int[] unlockedColumns, int[] hiddenColumns)
        {
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var ep = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells.LoadFromCollection(collection, true);

                worksheet.Protection.IsProtected = true;

                if (unlockedColumns != null && unlockedColumns.Any())
                {
                    foreach (var column in unlockedColumns)
                        worksheet.Column(column).Style.Locked = false;
                }

                if (hiddenColumns != null && hiddenColumns.Any())
                {
                    foreach (var column in hiddenColumns)
                        worksheet.Column(column).Hidden = true;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                ep.Save();
            }
            stream.Position = 0;
            return stream;
        }
    }
}