using System;
using System.Collections.Generic;
using System.Text;

namespace OpenOfficeUploader.Services
{
    public class ExcelIterationResult
    {
        public int RowNumber { get; set; }
        public bool Status { get; set; }
        public List<string> Reasons { get; set; }
    }
}
