using System.Collections.Generic;
using System.IO;

namespace ExcelDataExtractor.Dtos.Responses
{
    public class ExcelReadResponse<T>
    {
        public IEnumerable<T> RecordsToUpload { get; }
        public string ErrorMessage { get; }

        public ExcelReadResponse(IEnumerable<T> recordsToUpload, string errorMessage)
        {
            RecordsToUpload = recordsToUpload;
            ErrorMessage = errorMessage;
        }
    }

    public class ExcelWriteResponse
    {
        public string ExcelName { get; set; }
        public MemoryStream ExcelStream { get; set; }
    }
}
