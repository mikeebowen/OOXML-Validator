using OOXMLValidatorCLI.Interfaces;
using System.IO;

namespace OOXMLValidatorCLI.Classes
{
    public class DefaultFileService : IFileService
    {
        public FileAttributes GetAttributes(string path)
        {
            return File.GetAttributes(path);
        }
    }
}
