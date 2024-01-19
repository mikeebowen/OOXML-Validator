// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Classes
{
    using DocumentFormat.OpenXml;

    /// <summary>
    /// Represents information about a validation error.
    /// </summary>
    public class ValidationErrorInfoInternal
    {
        /// <summary>
        /// Gets or sets the type of the error.
        /// </summary>
        public string ErrorType { get; set; }

        /// <summary>
        /// Gets or sets the description of the error.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the XML path of the error.
        /// </summary>
        public XmlPath Path { get; set; }

        /// <summary>
        /// Gets or sets the ID of the error.
        /// </summary>
        public string Id { get; set; }
    }
}
