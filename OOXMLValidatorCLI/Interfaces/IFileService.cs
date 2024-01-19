// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Interfaces
{
    using System.IO;

    /// <summary>
    /// Represents a file service that provides operations related to file attributes.
    /// </summary>
    public interface IFileService
    {
        /// <summary>
        /// Gets the attributes of the specified file.
        /// </summary>
        /// <param name="path">The path to the file.</param>
        /// <returns>The attributes of the file.</returns>
        FileAttributes GetAttributes(string path);
    }
}
