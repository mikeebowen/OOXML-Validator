// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Classes
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Utility class for working with Open XML documents.
    /// </summary>
    public class DocumentUtils : IDocumentUtils
    {
        /// <summary>
        /// Opens a WordprocessingDocument from the specified file path.
        /// </summary>
        /// <param name="filePath">The path of the Word document.</param>
        /// <returns>The opened WordprocessingDocument.</returns>
        public WordprocessingDocument OpenWordprocessingDocument(string filePath)
        {
            return WordprocessingDocument.Open(filePath, false);
        }

        /// <summary>
        /// Opens a PresentationDocument from the specified file path.
        /// </summary>
        /// <param name="filePath">The path of the PowerPoint presentation.</param>
        /// <returns>The opened PresentationDocument.</returns>
        public PresentationDocument OpenPresentationDocument(string filePath)
        {
            return PresentationDocument.Open(filePath, false);
        }

        /// <summary>
        /// Opens a SpreadsheetDocument from the specified file path.
        /// </summary>
        /// <param name="filePath">The path of the Excel spreadsheet.</param>
        /// <returns>The opened SpreadsheetDocument.</returns>
        public SpreadsheetDocument OpenSpreadsheetDocument(string filePath)
        {
            return SpreadsheetDocument.Open(filePath, false);
        }

        /// <summary>
        /// Validates the specified OpenXmlPackage against the specified file format version.
        /// </summary>
        /// <param name="doc">The OpenXmlPackage to validate.</param>
        /// <param name="version">The file format version to validate against.</param>
        /// <returns>A tuple containing a boolean indicating if the validation is strict, and a collection of validation error information.</returns>
        public Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> Validate(OpenXmlPackage doc, FileFormatVersions version)
        {
            OpenXmlValidator openXmlValidator = new OpenXmlValidator(version);
            bool isStrict = doc.StrictRelationshipFound;

            IEnumerable<ValidationErrorInfo> validationErrorInfos = openXmlValidator.Validate(doc);
            IEnumerable<ValidationErrorInfoInternal> errors = validationErrorInfos.Select(e => new ValidationErrorInfoInternal()
            {
                ErrorType = Enum.GetName(e.ErrorType),
                Description = e.Description,
                Path = e.Path,
                Id = e.Id,
            });

            return new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(isStrict, errors);
        }
    }
}
