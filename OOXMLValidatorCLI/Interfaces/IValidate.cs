using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace OOXMLValidatorCLI.Interfaces
{
    public interface IValidate
    {
        string OOXML(string filePath, int? format);
    }
}
