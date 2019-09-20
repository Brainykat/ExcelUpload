using OpenOfficeUploader.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenOfficeUploader.Services
{

    public class UploadResult
    {
        private UploadResult() { }
        public UploadResult(bool status = true, string message = null)
        {
            Status = status;
            Message = message;
            Results = new List<ExcelIterationResult>();
            People = new List<Person>();
        }

        public bool Status { get; set; }
        public string Message { get; set; }
        public List<ExcelIterationResult> Results { get; set; }
        public List<Person> People { get; set; }
    }

}
