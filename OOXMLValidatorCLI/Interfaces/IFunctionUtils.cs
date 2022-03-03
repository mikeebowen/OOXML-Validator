using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
namespace OOXMLValidatorCLI.Interfaces
{
    public interface IFunctionUtils
    {
        FileFormatVersions OfficeVersion { get; }
        void SetOfficeVersion(string version);
        OpenXmlPackage GetDocument(string filePath);
        Tuple<bool, IEnumerable<ValidationErrorInfo>> GetValidationErrors(OpenXmlPackage doc);
        string GetValidationErrorsJson(Tuple<bool, IEnumerable<ValidationErrorInfo>> data);
    }
}
