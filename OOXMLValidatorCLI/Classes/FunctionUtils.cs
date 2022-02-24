using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Newtonsoft.Json;
using OOXMLValidatorCLI.Interfaces;

namespace OOXMLValidatorCLI.Classes
{
    public class FunctionUtils : IFunctionUtils
    {
        private readonly IDocumentUtils _documentUtils;
        private Nullable<FileFormatVersions> _fileFormatVersions;

        public FileFormatVersions OfficeVersion
        {
            get
            {
                return _fileFormatVersions ?? Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Max();
            }
        }
        public FunctionUtils(IDocumentUtils documentUtils)
        {
            _documentUtils = documentUtils;
            _fileFormatVersions = null;
        }

        public OpenXmlPackage GetDocument(string filePath)
        {
            string fileExtension = filePath.Substring(Math.Max(0, filePath.Length - 4)).ToLower();

            if (!new string[] { "docx", "docm", "dotm", "dotx", "pptx", "pptm", "potm", "potx", "ppam", "ppsm", "ppsx", "xlsx", "xlsm", "xltm", "xltx", "xlam" }.Contains(fileExtension))
            {
                throw new ArgumentException("file must be a .docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, or .xlam");
            }

            OpenXmlPackage doc = null;

            switch (fileExtension)
            {
                case "docx":
                case "docm":
                case "dotm":
                case "dotx":
                    doc = _documentUtils.OpenWordprocessingDocument(filePath);
                    break;
                case "pptx":
                case "pptm":
                case "potm":
                case "potx":
                case "ppam":
                case "ppsm":
                case "ppsx":
                    doc = _documentUtils.OpenPresentationDocument(filePath);
                    break;
                case "xlsx":
                case "xlsm":
                case "xltm":
                case "xltx":
                case "xlam":
                    doc = _documentUtils.OpenSpreadsheetDocument(filePath);
                    break;
                default:
                    break;
            }

            return doc;
        }

        public void SetOfficeVersion(string v)
        {
            if (v != null && Enum.TryParse(v, out FileFormatVersions version))
            {
                _fileFormatVersions = version;
            }
            else
            {
                FileFormatVersions currentVersion = Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Last();
                _fileFormatVersions = currentVersion;
            }
        }

        public IEnumerable<ValidationErrorInfo> GetValidationErrors(OpenXmlPackage doc)
        {
            return _documentUtils.Validate(doc, OfficeVersion);
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
