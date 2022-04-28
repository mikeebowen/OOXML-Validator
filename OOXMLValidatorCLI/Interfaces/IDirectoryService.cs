using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOXMLValidatorCLI.Interfaces
{
    public interface IDirectoryService
    {
        IEnumerable<string> EnumerateFiles(string path, string searchPattern, SearchOption searchOption);
        string[] GetFiles(string path);
    }
}
