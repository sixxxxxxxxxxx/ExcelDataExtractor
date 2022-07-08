using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace ExcelDataExtractor
{
    public static class ExcelDataExtractorExtension
    {
        public static List<Dictionary<string, string>> ExtractFromExcel(IFormFile file, string[]? nullableFields, string[]? ignoreFields)
        {
            int headerRow = 1;
            int contentRow = 2;

            if (file == null)
            {
                throw new InvalidOperationException("File is empty");
            }

            if (!Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Incorrect file format");
            }

            List<Dictionary<string, string>> excelData = new();

            List<string> possibleEmptyFields = new();


            if (nullableFields != null)
                possibleEmptyFields.AddRange(nullableFields);

            possibleEmptyFields = possibleEmptyFields.Select(c => c.ToLower()).ToList();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using MemoryStream stream = new();
            file.CopyTo(stream);
            using ExcelPackage ep = new(stream);
            ExcelWorksheet worksheet = ep.Workbook.Worksheets.First();
            int rowCount = worksheet.Dimension.Rows;
            int columnCount = worksheet.Dimension.Columns;
            nullableFields = nullableFields?.Select(c => c.ToLower()).ToArray();
            ignoreFields = ignoreFields?.Select(c => c.ToLower()).ToArray();

            if (nullableFields != null)
                worksheet.Validate(nullableFields, columnCount);

            if (ignoreFields != null)
                worksheet.Validate(ignoreFields, columnCount);


            for (int row = contentRow; row <= rowCount; row++)
            {
                Dictionary<string, string> cell = new();
                for (int column = 1; column <= columnCount; column++)
                {
                    string headerCell = worksheet.Cells[headerRow, column].Value.ToString();

                    if (ignoreFields != null && ignoreFields.Contains(headerCell?.ToLower()))
                        continue;

                    string value;
                    if (possibleEmptyFields.Any(l => !string.IsNullOrWhiteSpace(headerCell) && l == headerCell.ToLower()))
                    {
                        value = worksheet.Cells[row, column].Value?.ToString();
                    }
                    else if (worksheet.Cells[row, column].Value == null)
                    {
                        throw new InvalidDataException("Excel has empty fields. Crosscheck it and submit again");
                    }
                    else
                    {
                        value = worksheet.Cells[row, column].Value.ToString();
                    }

                    if (headerCell != null)
                        cell.Add(headerCell, value);
                }

                excelData.Add(cell);
            }
            ep.Save();
            return excelData;
        }

        public static void Validate(this ExcelWorksheet worksheet, string[] fields, int columnCount)
        {
            if (fields == null) return;
            foreach (var field in fields)
            {
                if (!Headers(worksheet, columnCount).Contains(field))
                    throw new InvalidDataException($"{field} Field Doesn't Exist");
            }
        }
        public static void ValidateFields(this List<Dictionary<string, string>> source, string[] fields)
        {
            var equal = source.FirstOrDefault()?.Select(c => c.Key).OrderBy(c => c, StringComparer.OrdinalIgnoreCase)
                .SequenceEqual(fields.Select(c => c).OrderBy(c => c, StringComparer.OrdinalIgnoreCase));

            if (equal != null && !equal.Value)
                throw new InvalidOperationException("The excel sheet uploaded is not for this purpose, crosscheck your headers or download the sample excel provided");
        }

        public static void CheckDuplicate(this List<Dictionary<string, string>> source, string column)
        {
            List<string> list = new List<string>();

            foreach (var row in source)
            {
                list.Add(row[column]);
            }

            if (list.GroupBy(x => x).Any(g => g.Count() > 1))
            {
                throw new InvalidOperationException("Data contains duplicate records");
            }
        }

        private static List<string> Headers(ExcelWorksheet worksheet, int columnCount)
        {
            List<string> headers = new List<string>(columnCount);
            for (int column = 1; column <= columnCount; column++)
            {
                var header = worksheet.Cells[1, column].Value.ToString()?.ToLower();
                headers.Add(header);
            }
            return headers;
        }
    }
}
