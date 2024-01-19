// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Classes
{
    using System.IO;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Represents a service for working with files.
    /// </summary>
    public class FileService : IFileService
    {
        /// <summary>
        /// Gets the attributes of a file at the specified path.
        /// </summary>
        /// <param name="path">The path of the file.</param>
        /// <returns>The attributes of the file.</returns>
        public FileAttributes GetAttributes(string path)
        {
            return File.GetAttributes(path);
        }
    }
}
