using System.IO;

namespace OOXMLValidatorCLI.Interfaces
{
    public interface IFileService
    {
        public FileAttributes GetAttributes(string path);
    }
}
