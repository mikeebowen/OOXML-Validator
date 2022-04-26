using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Newtonsoft.Json;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Xml.Linq;

namespace OOXMLValidatorCLI.Classes
{
    public class FunctionUtils : IFunctionUtils
    {
        private readonly IDocumentUtils _documentUtils;
        private FileFormatVersions? _fileFormatVersions;

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

        public OpenXmlPackage GetDocument(string filePath, string fileExtension)
        {
            OpenXmlPackage doc = null;

            switch (fileExtension)
            {
                case ".docx":
                case ".docm":
                case ".dotm":
                case ".dotx":
                    doc = _documentUtils.OpenWordprocessingDocument(filePath);
                    break;
                case ".pptx":
                case ".pptm":
                case ".potm":
                case ".potx":
                case ".ppam":
                case ".ppsm":
                case ".ppsx":
                    doc = _documentUtils.OpenPresentationDocument(filePath);
                    break;
                case ".xlsx":
                case ".xlsm":
                case ".xltm":
                case ".xltx":
                case ".xlam":
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

        public Tuple<bool, IEnumerable<ValidationErrorInfo>> GetValidationErrors(OpenXmlPackage doc)
        {
            return _documentUtils.Validate(doc, OfficeVersion);
        }

        public object GetValidationErrors(Tuple<bool, IEnumerable<ValidationErrorInfo>> validationInfo, string filePath, bool returnXml)
        {
            if (!returnXml)
            {
                List<dynamic> res = new List<dynamic>();

                foreach (ValidationErrorInfo validationErrorInfo in validationInfo.Item2)
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
            else
            {
                XElement xml = new XElement("ValidationErrorInfoList");
                xml.SetAttributeValue("FilePath", filePath);
                xml.SetAttributeValue("IsStrict", validationInfo.Item1);

                foreach (ValidationErrorInfo validationErrorInfo in validationInfo.Item2)
                {
                    xml.Add(
                        new XElement("ValidationErrorInfo",
                            new XElement("Description", validationErrorInfo.Description),
                            new XElement("Path", validationErrorInfo.Path),
                            new XElement("Id", validationErrorInfo.Id),
                            new XElement("ErrorType", validationErrorInfo.ErrorType)
                        )
                    );
                }

                return new XDocument(xml);
            }
        }
    }
}
