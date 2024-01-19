// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Interfaces
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using OOXMLValidatorCLI.Classes;

    /// <summary>
    /// Represents a set of utility functions for working with Open XML documents.
    /// </summary>
    public interface IFunctionUtils
    {
        /// <summary>
        /// Gets the Office version for the Open XML document.
        /// </summary>
        FileFormatVersions OfficeVersion { get; }

        /// <summary>
        /// Sets the Office version for the Open XML document.
        /// </summary>
        /// <param name="version">The Office version to set.</param>
        void SetOfficeVersion(string version);

        /// <summary>
        /// Gets the OpenXmlPackage object for the specified file.
        /// </summary>
        /// <param name="filePath">The path of the file.</param>
        /// <param name="fileExtension">The extension of the file.</param>
        /// <returns>The OpenXmlPackage object.</returns>
        OpenXmlPackage GetDocument(string filePath, string fileExtension);

        /// <summary>
        /// Gets the validation errors for the specified OpenXmlPackage object.
        /// </summary>
        /// <param name="doc">The OpenXmlPackage object.</param>
        /// <returns>A tuple containing a boolean value indicating whether there are validation errors and a collection of validation error information.</returns>
        Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> GetValidationErrors(OpenXmlPackage doc);

        /// <summary>
        /// Gets the validation errors data in the specified format.
        /// </summary>
        /// <param name="data">The tuple containing the validation errors information.</param>
        /// <param name="filePath">The path of the file.</param>
        /// <param name="returnXml">A boolean value indicating whether to return the validation errors data as XML.</param>
        /// <returns>The validation errors data.</returns>
        object GetValidationErrorsData(Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> data, string filePath, bool returnXml);
    }
}
