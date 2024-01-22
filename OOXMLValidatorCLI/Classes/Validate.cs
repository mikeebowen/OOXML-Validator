// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Classes
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using Newtonsoft.Json;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Represents a class that provides validation functionality for OOXML files.
    /// </summary>
    public class Validate : IValidate
    {
        private readonly IFunctionUtils functionUtils;
        private readonly string[] validFileExtensions = new string[] { ".docx", ".docm", ".dotm", ".dotx", ".pptx", ".pptm", ".potm", ".potx", ".ppam", ".ppsm", ".ppsx", ".xlsx", ".xlsm", ".xltm", ".xltx", ".xlam" };
        private readonly IFileService fileService;
        private readonly IDirectoryService directoryService;

        /// <summary>
        /// Initializes a new instance of the <see cref="Validate"/> class.
        /// </summary>
        /// <param name="functionUtils">The function utilities.</param>
        /// <param name="fileService">The file service.</param>
        /// <param name="directoryService">The directory service.</param>
        public Validate(IFunctionUtils functionUtils, IFileService fileService, IDirectoryService directoryService)
        {
            this.functionUtils = functionUtils;
            this.fileService = fileService;
            this.directoryService = directoryService;
        }

        /// <summary>
        /// Validates the specified OOXML file.
        /// </summary>
        /// <param name="filePath">The path to the file.</param>
        /// <param name="format">The format of the file.</param>
        /// <param name="returnXml">Indicates whether to return the validation errors as XML.</param>
        /// <param name="recursive">Indicates whether to recursively validate files in subdirectories.</param>
        /// <param name="includeValid">Indicates whether to include valid files in the result.</param>
        /// <returns>The validation result.</returns>
        public object OOXML(string filePath, string format, bool returnXml = false, bool recursive = false, bool includeValid = false)
        {
            this.functionUtils.SetOfficeVersion(format);

            FileAttributes fileAttributes = this.fileService.GetAttributes(filePath);

            if (fileAttributes.HasFlag(FileAttributes.Directory))
            {
                IEnumerable<string> files = recursive ? this.directoryService.EnumerateFiles(filePath, "*.*", SearchOption.AllDirectories).Where(f => this.validFileExtensions.Contains(Path.GetExtension(f)))
                    : this.directoryService.GetFiles(filePath).Where(f => this.validFileExtensions.Contains(Path.GetExtension(f)));

                XDocument xDocument = new XDocument(new XElement("Document"));
                List<object> validationErrorList = new List<object>();

                foreach (string file in files)
                {
                    string fileExtension = Path.GetExtension(file);

                    Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> validationTuple = this.GetValidationErrors(file, fileExtension);

                    if (validationTuple.Item2.Count() > 0 || includeValid)
                    {
                        var data = this.functionUtils.GetValidationErrorsData(validationTuple, file, returnXml);

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
                                                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                                                });
            }
            else
            {
                string fileExtension = Path.GetExtension(filePath);

                if (!this.validFileExtensions.Contains(fileExtension))
                {
                    throw new ArgumentException(string.Concat("file must have one of these extensions: ", string.Join(", ", this.validFileExtensions)));
                }

                Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> validationErrorInfos = this.GetValidationErrors(filePath, fileExtension);

                return this.functionUtils.GetValidationErrorsData(validationErrorInfos, filePath, returnXml);
            }
        }

        /// <summary>
        /// Gets the validation errors for the specified file.
        /// </summary>
        /// <param name="filePath">The path to the file.</param>
        /// <param name="fileExtension">The extension of the file.</param>
        /// <returns>A tuple containing a boolean indicating if the file is valid and a collection of validation error information.</returns>
        private Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> GetValidationErrors(string filePath, string fileExtension)
        {
            try
            {
                OpenXmlPackage doc = this.functionUtils.GetDocument(filePath, fileExtension);

                return this.functionUtils.GetValidationErrors(doc);
            }
            catch (Exception ex)
            {
                List<ValidationErrorInfoInternal> errors = new List<ValidationErrorInfoInternal> { new ValidationErrorInfoInternal() { Description = ex.Message, ErrorType = "OpenXmlPackageException" } };

                return new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(false, errors);
            }
        }
    }
}
