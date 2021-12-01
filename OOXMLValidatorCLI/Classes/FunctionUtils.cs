using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Newtonsoft.Json;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOXMLValidatorCLI.Classes
{
    public class FunctionUtils : IFunctionUtils
    {
        private readonly IDocument _document;
        private Nullable<FileFormatVersions> _fileFormatVersions;

        public FileFormatVersions OfficeVersion
        {
            get
            {
                return _fileFormatVersions ?? Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Max();
            }
        }
        public FunctionUtils(IDocument document)
        {
            _document = document;
            _fileFormatVersions = null;
        }

        public dynamic GetDocument(string filePath)
        {
            string fileExtension = filePath.Substring(Math.Max(0, filePath.Length - 4)).ToLower();

            if (!new string[] { "docx", "pptx", "xlsx" }.Contains(fileExtension))
            {
                throw new ArgumentException("file must be a .docx, .xlsx, or .pptx");
            }

            dynamic doc = null;

            switch (fileExtension)
            {
                case "docx":
                    doc = _document.OpenWordprocessingDocument(filePath);
                    break;
                case "pptx":
                    doc = _document.OpenPresentationDocument(filePath);
                    break;
                case "xlsx":
                    doc = _document.OpenSpreadsheetDocument(filePath);
                    break;
                default:
                    break;
            }

            return doc;
        }

        public void SetOfficeVersion(int? version)
        {
            if (version != null && Enum.IsDefined(typeof(FileFormatVersions), version))
            {
                _fileFormatVersions = (FileFormatVersions)(version);
            }
            else
            {
                FileFormatVersions currentVersion = Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Last();
                _fileFormatVersions = currentVersion;
            }
        }

        public IEnumerable<ValidationErrorInfo> GetValidationErrors(dynamic doc)
        {
            return _document.Validate(doc, OfficeVersion);
        }

        public string GetValidationErrorsJson(IEnumerable<ValidationErrorInfo> validationErrors)
        {
            List<dynamic> res = new List<dynamic>();

            foreach (ValidationErrorInfo validationErrorInfo in validationErrors)
            {
                dynamic dyno = new ExpandoObject();
                dyno.Description = validationErrorInfo.Description;
                dyno.Path = validationErrorInfo.Path;
                dyno.Id = validationErrorInfo.Id;
                dyno.ErrorType = validationErrorInfo.ErrorType;
                res.Add(dyno);
            }
            string json = JsonConvert.SerializeObject(res, Formatting.None,
                        new JsonSerializerSettings()
                        {
                            ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                        });

            return json;
        }
    }
}
