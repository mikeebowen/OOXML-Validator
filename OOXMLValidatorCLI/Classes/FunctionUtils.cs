// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Classes
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Linq;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using Newtonsoft.Json;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Utility class for performing various functions related to Open XML documents.
    /// </summary>
    public class FunctionUtils : IFunctionUtils
    {
        private readonly IDocumentUtils documentUtils;
        private FileFormatVersions? fileFormatVersions;

        /// <summary>
        /// Initializes a new instance of the <see cref="FunctionUtils"/> class.
        /// </summary>
        /// <param name="documentUtils">The document utility object.</param>
        public FunctionUtils(IDocumentUtils documentUtils)
        {
            this.documentUtils = documentUtils;
            this.fileFormatVersions = null;
        }

        /// <summary>
        /// Gets the maximum supported Office version based on the available FileFormatVersions.
        /// </summary>
        public FileFormatVersions OfficeVersion
        {
            get
            {
                return this.fileFormatVersions ?? Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Max();
            }
        }

        /// <summary>
        /// Gets the OpenXmlPackage object for the specified file.
        /// </summary>
        /// <param name="filePath">The path of the file.</param>
        /// <param name="fileExtension">The extension of the file.</param>
        /// <returns>The OpenXmlPackage object.</returns>
        public OpenXmlPackage GetDocument(string filePath, string fileExtension)
        {
            OpenXmlPackage doc = null;

            switch (fileExtension)
            {
                case ".docx":
                case ".docm":
                case ".dotm":
                case ".dotx":
                    doc = this.documentUtils.OpenWordprocessingDocument(filePath);
                    break;
                case ".pptx":
                case ".pptm":
                case ".potm":
                case ".potx":
                case ".ppam":
                case ".ppsm":
                case ".ppsx":
                    doc = this.documentUtils.OpenPresentationDocument(filePath);
                    break;
                case ".xlsx":
                case ".xlsm":
                case ".xltm":
                case ".xltx":
                case ".xlam":
                    doc = this.documentUtils.OpenSpreadsheetDocument(filePath);
                    break;
                default:
                    break;
            }

            return doc;
        }

        /// <summary>
        /// Sets the Office version based on the provided string value.
        /// </summary>
        /// <param name="v">The string representation of the Office version.</param>
        public void SetOfficeVersion(string v)
        {
            if (v is not null && Enum.TryParse(v, out FileFormatVersions version))
            {
                this.fileFormatVersions = version;
            }
            else
            {
                FileFormatVersions currentVersion = Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Last();
                this.fileFormatVersions = currentVersion;
            }
        }

        /// <summary>
        /// Validates the specified OpenXmlPackage object and returns the validation errors.
        /// </summary>
        /// <param name="doc">The OpenXmlPackage object to validate.</param>
        /// <returns>A tuple containing a boolean value indicating if the validation is strict and a collection of validation error information.</returns>
        public Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> GetValidationErrors(OpenXmlPackage doc)
        {
            return this.documentUtils.Validate(doc, this.OfficeVersion);
        }

        /// <summary>
        /// Gets the validation errors data in the specified format.
        /// </summary>
        /// <param name="validationInfo">The validation information.</param>
        /// <param name="filePath">The path of the file.</param>
        /// <param name="returnXml">A boolean value indicating if the data should be returned in XML format.</param>
        /// <returns>The validation errors data.</returns>
        public object GetValidationErrorsData(Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> validationInfo, string filePath, bool returnXml)
        {
            if (!returnXml)
            {
                List<dynamic> res = new List<dynamic>();

                foreach (ValidationErrorInfoInternal validationErrorInfo in validationInfo.Item2)
                {
                    dynamic dyno = new ExpandoObject();
                    dyno.Description = validationErrorInfo.Description;
                    dyno.Path = validationErrorInfo.Path;
                    dyno.Id = validationErrorInfo.Id;
                    dyno.ErrorType = validationErrorInfo.ErrorType;
                    res.Add(dyno);
                }

                string json = JsonConvert.SerializeObject(
                    res,
                    Formatting.None,
                    new JsonSerializerSettings()
                    {
                        ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                    });

                return json;
            }
            else
            {
                XElement element;
                ValidationErrorInfoInternal first = validationInfo.Item2.FirstOrDefault();

                if (first?.ErrorType == "OpenXmlPackageException")
                {
                    element = new XElement(
                        "Exceptions",
                        new XElement(
                            "OpenXmlPackageException",
                            new XElement("Message", first.Description)));
                }
                else
                {
                    element = new XElement("ValidationErrorInfoList");

                    foreach (ValidationErrorInfoInternal validationErrorInfo in validationInfo.Item2)
                    {
                        element.Add(
                            new XElement(
                                "ValidationErrorInfo",
                                new XElement("Description", validationErrorInfo.Description),
                                new XElement("Path", validationErrorInfo.Path),
                                new XElement("Id", validationErrorInfo.Id),
                                new XElement("ErrorType", validationErrorInfo.ErrorType)));
                    }
                }

                XElement xml = new XElement("File", element);
                xml.SetAttributeValue("FilePath", filePath);
                xml.SetAttributeValue("IsStrict", validationInfo.Item1);

                return new XDocument(xml);
            }
        }
    }
}
