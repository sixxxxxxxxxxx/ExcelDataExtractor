using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.Text;

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

        public static (List<Dictionary<string, string>> data, string errMSg) Distinct(
            this List<Dictionary<string, string>> source, string[]? columns)
        {
            StringBuilder sb = new StringBuilder();
            List<Dictionary<string, string>> listWithoutDuplicates = new List<Dictionary<string, string>>();

            if (columns == null) return (source, string.Empty);

            foreach (string column in columns)
            {
                listWithoutDuplicates = new List<Dictionary<string, string>>();

                StringBuilder msgBuilder = new StringBuilder();


                foreach (Dictionary<string, string> item in source)
                {
                    var columnValue = item.TryGetValue(column, out var value) ? value : null;

                    if (listWithoutDuplicates.Any(i => i.TryGetValue(column, out var existingValue) && existingValue == columnValue))
                    {

                        Dictionary<string, string> itemToRemove = listWithoutDuplicates.SingleOrDefault(l => l.ContainsValue(item[column]));

                        int indexOfItemToRemove = listWithoutDuplicates.IndexOf(itemToRemove);

                        listWithoutDuplicates.RemoveAt(indexOfItemToRemove);
                        msgBuilder.AppendLine(item[column]);
                        continue;
                    }

                    if (!listWithoutDuplicates.Contains(item) && item.ContainsKey(column)) listWithoutDuplicates.Add(item);
                }

                source = listWithoutDuplicates.ToList();

                string duplicateRecordsErrorMsg = string.IsNullOrEmpty(msgBuilder.ToString())
                    ? string.Empty
                    : msgBuilder.Insert(0,
                            $"The following '{column}' values are duplicates and not uploaded !\nReview them and re-upload !\n")
                        .ToString();

                if (!string.IsNullOrEmpty(duplicateRecordsErrorMsg)) sb.AppendLine(duplicateRecordsErrorMsg);
            }

            return (listWithoutDuplicates, sb.ToString());
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
