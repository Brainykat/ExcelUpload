using OfficeOpenXml;
using OpenOfficeUploader.Entities;
using OpenOfficeUploader.Validators;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace OpenOfficeUploader.Services
{
    public class AddPeopleFromExcel
    {
        // Await for db call on need
        public async Task<UploadResult> AddPeople(string path, string pwd = default)
        {
            var file = new FileInfo(path);
            var uploadResult = new UploadResult();
            if (file != null && file.Length > 0 && !string.IsNullOrEmpty(file.Name))
            {
                //You can save a copy of the file if u wish
                var fileName = Path.GetFileName(file.Name);
                if (Path.GetExtension(fileName) == ".xls" || Path.GetExtension(fileName) == ".xlsx")
                {
                    //If file has a password supply it as second parameter below
                    using (var package = new ExcelPackage(file))
                    {
                        List<Person> people = new List<Person>();
                        List<ExcelIterationResult> excelIterationResults = new List<ExcelIterationResult>();
                        //var workSheet = currentSheet.First(); ///Use this for only single worksheet
                        foreach (var workSheet in package.Workbook.Worksheets)
                        {
                            //Assumes first row is header row
                            for (int rowIterator = 2; rowIterator <= workSheet.Dimension.End.Row; rowIterator++)
                            {
                                ExcelIterationResult Er = new ExcelIterationResult();
                                //Columns have to be in order
                                var surName = workSheet.Cells[rowIterator, 1].Value.ToString().Trim();
                                var firstName = workSheet.Cells[rowIterator, 2].Value.ToString().Trim();
                                var middleName = workSheet.Cells[rowIterator, 3].Value.ToString().Trim();
                                var phone = workSheet.Cells[rowIterator, 4].Value.ToString().Trim();
                                var email = workSheet.Cells[rowIterator, 5].Value.ToString().Trim();
                                var dob = workSheet.Cells[rowIterator, 6].Value.ToString().Trim();
                                List<string> mess = new List<string>();
                                if (string.IsNullOrWhiteSpace(surName)) mess.Add("Surname is required");
                                if (string.IsNullOrWhiteSpace(firstName)) mess.Add("Surname is required");
                                if (string.IsNullOrWhiteSpace(phone)) mess.Add("Surname is required");
                                if (string.IsNullOrWhiteSpace(email)) mess.Add("Surname is required");
                                //Validate your values
                                if (!EmailValidation.IsValidEmail(email)) mess.Add($"{email} is invalid");
                                if (!PhoneNumberValidation.IsValidPhoneNumber(phone)) mess.Add($"{phone} is invalid");
                                if (mess.Any())
                                {
                                    Er.Status = false;
                                    Er.RowNumber = rowIterator;
                                    Er.Reasons = mess;
                                    excelIterationResults.Add(Er);
                                }
                                else
                                {
                                    people.Add(Person.Create(surName, firstName, phone, email, Convert.ToDateTime(dob), middleName));
                                }
                            }
                        }
                        uploadResult.People.AddRange(people);
                        uploadResult.Results.AddRange(excelIterationResults);
                        if (excelIterationResults.Any()) { uploadResult.Status = false; uploadResult.Message = "Some data was invalid"; }
                        else { uploadResult.Status = true; }
                    }
                }
                else
                {
                    uploadResult.Status = false;
                    uploadResult.Message = "Not a valid excel file";
                }
            }
            else
            {
                uploadResult.Status = false;
                uploadResult.Message = "Empty file";
            }
            return uploadResult;
        }
    }
}
