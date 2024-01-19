// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Interfaces
{
    /// <summary>
    /// Represents an interface for validating OOXML files.
    /// </summary>
    public interface IValidate
    {
        /// <summary>
        /// Validates the specified OOXML file.
        /// </summary>
        /// <param name="filePath">The path to the OOXML file.</param>
        /// <param name="format">The format of the validation result.</param>
        /// <param name="returnXml">Specifies whether to return the validation result as XML.</param>
        /// <param name="recursive">Specifies whether to validate files recursively in subdirectories.</param>
        /// <param name="includeValid">Specifies whether to include valid files in the validation result.</param>
        /// <returns>The validation result.</returns>
        object OOXML(string filePath, string format, bool returnXml, bool recursive, bool includeValid);
    }
}
