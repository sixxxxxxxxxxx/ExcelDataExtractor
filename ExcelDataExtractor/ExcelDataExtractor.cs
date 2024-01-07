using ExcelDataExtractor.Dtos.Responses;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace ExcelDataExtractor
{
    public static class Excel
    {
        public static ExcelReadResponse<T> ReadFromExcel<T>(IFormFile excelFile, string[]? requiredHeaders, string[]? nullableFields, string[]? ignoreFields, string duplicateComparer, string[]? uniqueColumns) where T : new()
        {
            List<Dictionary<string, string>> excelData = ExcelDataExtractorExtension.ExtractFromExcel(excelFile, nullableFields, ignoreFields);

            if (requiredHeaders != null)
            {
                excelData.ValidateFields(requiredHeaders);
            }
            (List<Dictionary<string, string>> data, string distictErrMSg) = excelData.Distinct(uniqueColumns);
            
            IEnumerable<T> recordsToUpload = DictionaryToObjectConverterExtension.DictionaryToObjects<T>(excelData);
            return new ExcelReadResponse<T>(recordsToUpload, distictErrMSg);           
        }

        public static ExcelWriteResponse WriteToExcel<T>(this IList<T> collection, string worksheetName, int[] unlockedColumns, int[] hiddenColumns)
        {
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var ep = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells.LoadFromCollection(collection, true);

                worksheet.Protection.IsProtected = true;
                worksheet.Protection.SetPassword("Danfodio2!");
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
            var fileName = $"{worksheetName}-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";

            return new ExcelWriteResponse { ExcelStream = stream, ExcelName = fileName };
        }
    }
}