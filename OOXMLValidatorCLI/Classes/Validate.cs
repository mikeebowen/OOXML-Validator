using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Newtonsoft.Json;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OOXMLValidatorCLI.Classes
{
    public class Validate : IValidate
    {
        private readonly IFunctionUtils _functionUtils;
        private readonly string[] validFileExtensions = new string[] { ".docx", ".docm", ".dotm", ".dotx", ".pptx", ".pptm", ".potm", ".potx", ".ppam", ".ppsm", ".ppsx", ".xlsx", ".xlsm", ".xltm", ".xltx", ".xlam" };


        public Validate(IFunctionUtils functionUtils)
        {
            _functionUtils = functionUtils;
        }
        public object OOXML(string filePath, string format, bool returnXml = false)
        {

            _functionUtils.SetOfficeVersion(format);

            FileAttributes fileAttributes = File.GetAttributes(filePath);

            if (fileAttributes.HasFlag(FileAttributes.Directory))
            {
                IEnumerable<string> files = Directory.EnumerateFiles(filePath, "*.*", SearchOption.AllDirectories).Where(f =>
                {
                    return validFileExtensions.Contains(Path.GetExtension(f));
                });
                XDocument xDocument = new XDocument(new XElement("Document"));
                List<object> validationErrorList = new List<object>();

                foreach (string file in files)
                {
                    string fileExtension = Path.GetExtension(file);

                    Tuple<bool, IEnumerable<ValidationErrorInfo>> validationErrorInfos = _getValidationErrors(file, fileExtension, true);

                    if (validationErrorInfos != null)
                    {
                        var data = _functionUtils.GetValidationErrors(validationErrorInfos, file, returnXml);

                        if (returnXml)
                        {
                            xDocument.Root.Add((data as XDocument).Root);
                        }
                        else
                        {
                            var errorList = new { FilePath = file, ValidationErrors = data };

                            validationErrorList.Add(errorList);
                        }
                    }
                }

                return returnXml ? xDocument : JsonConvert.SerializeObject(
                                                validationErrorList,
                                                Formatting.None,
                                                new JsonSerializerSettings()
                                                {
                                                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                                                }
                                            ); ;
            }
            else
            {
                string fileExtension = Path.GetExtension(filePath);
                Tuple<bool, IEnumerable<ValidationErrorInfo>> validationErrorInfos = _getValidationErrors(filePath, fileExtension, false);

                return _functionUtils.GetValidationErrors(validationErrorInfos, filePath, returnXml);
            }

        }

        private Tuple<bool, IEnumerable<ValidationErrorInfo>> _getValidationErrors(string filePath, string fileExtension, bool ignoreInvalidFiles)
        {
            FileAttributes fileAttributes = File.GetAttributes(filePath);

            if (
                fileAttributes.HasFlag(FileAttributes.Directory) ||
                !validFileExtensions.Contains(fileExtension)
            )
            {
                if (!ignoreInvalidFiles)
                {
                    throw new ArgumentException(string.Concat("file must be have one of these extensions: ", string.Join(", ", validFileExtensions)));
                }
                else
                {
                    return null;
                }
            }

            OpenXmlPackage doc = _functionUtils.GetDocument(filePath, fileExtension);

            return _functionUtils.GetValidationErrors(doc);
        }
    }
}
