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
        IEnumerable<ValidationErrorInfo> GetValidationErrors(OpenXmlPackage doc);
        string GetValidationErrorsJson(IEnumerable<ValidationErrorInfo> validationErrors);
    }
}
