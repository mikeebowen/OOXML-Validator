// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Interfaces
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using OOXMLValidatorCLI.Classes;

    /// <summary>
    /// Represents a utility for working with Open XML documents.
    /// </summary>
    public interface IDocumentUtils
    {
        /// <summary>
        /// Opens a WordprocessingDocument from the specified temporary file path.
        /// </summary>
        /// <param name="tempFilePath">The temporary file path.</param>
        /// <returns>The opened WordprocessingDocument.</returns>
        WordprocessingDocument OpenWordprocessingDocument(string tempFilePath);

        /// <summary>
        /// Opens a SpreadsheetDocument from the specified temporary file path.
        /// </summary>
        /// <param name="tempFilePath">The temporary file path.</param>
        /// <returns>The opened SpreadsheetDocument.</returns>
        SpreadsheetDocument OpenSpreadsheetDocument(string tempFilePath);

        /// <summary>
        /// Opens a PresentationDocument from the specified temporary file path.
        /// </summary>
        /// <param name="tempFilePath">The temporary file path.</param>
        /// <returns>The opened PresentationDocument.</returns>
        PresentationDocument OpenPresentationDocument(string tempFilePath);

        /// <summary>
        /// Validates the specified OpenXmlPackage against the specified file format versions.
        /// </summary>
        /// <param name="doc">The OpenXmlPackage to validate.</param>
        /// <param name="version">The file format versions to validate against.</param>
        /// <returns>A tuple containing a boolean indicating whether the validation succeeded and a collection of validation error information.</returns>
        Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> Validate(OpenXmlPackage doc, FileFormatVersions version);
    }
}
