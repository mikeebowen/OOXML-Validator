// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLI.Classes
{
    using System.Collections.Generic;
    using System.IO;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Provides methods for working with directories and files.
    /// </summary>
    internal class DirectoryService : IDirectoryService
    {
        /// <summary>
        /// Enumerates files in a directory that match the specified search pattern and search option.
        /// </summary>
        /// <param name="path">The path to the directory.</param>
        /// <param name="searchPattern">The search pattern to match against the file names.</param>
        /// <param name="searchOption">Specifies whether to search the current directory only or all subdirectories as well.</param>
        /// <returns>An enumerable collection of file names that match the search pattern and search option.</returns>
        public IEnumerable<string> EnumerateFiles(string path, string searchPattern, SearchOption searchOption)
        {
            return Directory.EnumerateFiles(path, searchPattern, searchOption);
        }

        /// <summary>
        /// Returns the names of files in the specified directory.
        /// </summary>
        /// <param name="path">The path to the directory.</param>
        /// <returns>An array of file names in the specified directory.</returns>
        public string[] GetFiles(string path)
        {
            return Directory.GetFiles(path);
        }
    }
}
