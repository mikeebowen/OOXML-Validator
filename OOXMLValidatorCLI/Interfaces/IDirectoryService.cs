// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Interfaces
{
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Represents a directory service for working with directories and files.
    /// </summary>
    public interface IDirectoryService
    {
        /// <summary>
        /// Enumerates files in a specified directory that match a specified search pattern.
        /// </summary>
        /// <param name="path">The path to the directory to search.</param>
        /// <param name="searchPattern">The search string to match against the names of files in the directory.</param>
        /// <param name="searchOption">Specifies whether to search the current directory, or all subdirectories as well.</param>
        /// <returns>An enumerable collection of the full names (including paths) for the files in the directory that match the specified search pattern.</returns>
        IEnumerable<string> EnumerateFiles(string path, string searchPattern, SearchOption searchOption);

        /// <summary>
        /// Returns the names of files in a specified directory.
        /// </summary>
        /// <param name="path">The path to the directory to search.</param>
        /// <returns>An array of the full names (including paths) for the files in the directory.</returns>
        string[] GetFiles(string path);
    }
}
