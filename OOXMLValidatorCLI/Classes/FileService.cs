using OOXMLValidatorCLI.Interfaces;
using System.IO;

namespace OOXMLValidatorCLI.Classes
{
    public class FileService : IFileService
    {
        public FileAttributes GetAttributes(string path)
        {
            return File.GetAttributes(path);
        }
    }
}
