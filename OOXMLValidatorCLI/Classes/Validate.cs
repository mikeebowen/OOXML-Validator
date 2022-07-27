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
        private readonly IFileService _fileService;
        private readonly IDirectoryService _directoryService;


        public Validate(IFunctionUtils functionUtils, IFileService fileService, IDirectoryService directoryService)
        {
            _functionUtils = functionUtils;
            _fileService = fileService;
            _directoryService = directoryService;
        }
        public object OOXML(string filePath, string format, bool returnXml = false, bool recursive = false, bool includeValid = false)
        {

            _functionUtils.SetOfficeVersion(format);

            FileAttributes fileAttributes = _fileService.GetAttributes(filePath);

            if (fileAttributes.HasFlag(FileAttributes.Directory))
            {
                IEnumerable<string> files = recursive ? _directoryService.EnumerateFiles(filePath, "*.*", SearchOption.AllDirectories).Where(f => validFileExtensions.Contains(Path.GetExtension(f)))
                    : _directoryService.GetFiles(filePath).Where(f => validFileExtensions.Contains(Path.GetExtension(f)));

                XDocument xDocument = new XDocument(new XElement("Document"));
                List<object> validationErrorList = new List<object>();

                foreach (string file in files)
                {
                    string fileExtension = Path.GetExtension(file);

                    Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> validationTupple = _getValidationErrors(file, fileExtension);

                    if (validationTupple.Item2.Count() > 0 || includeValid)
                    {
                        var data = _functionUtils.GetValidationErrorsData(validationTupple, file, returnXml);

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

                if (!validFileExtensions.Contains(fileExtension))
                {
                    throw new ArgumentException(string.Concat("file must be have one of these extensions: ", string.Join(", ", validFileExtensions)));
                }

                Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> validationErrorInfos = _getValidationErrors(filePath, fileExtension);

                return _functionUtils.GetValidationErrorsData(validationErrorInfos, filePath, returnXml);
            }

        }

        private Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> _getValidationErrors(string filePath, string fileExtension)
        {
            try
            {
                OpenXmlPackage doc = _functionUtils.GetDocument(filePath, fileExtension);

                return _functionUtils.GetValidationErrors(doc);
            }
            catch (Exception ex)
            {
                List<ValidationErrorInfoInternal> errors = new List<ValidationErrorInfoInternal> { new ValidationErrorInfoInternal() { Description = ex.Message, ErrorType = "OpenXmlPackageException" } };

                return new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(false, errors);
            }
        }
    }
}
