using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace OOXMLValidatorCLI.Interfaces
{
    public interface IFunctionUtils
    {
        FileFormatVersions OfficeVersion { get; }
        void SetOfficeVersion(string version);
        dynamic GetDocument(string filePath);
        IEnumerable<ValidationErrorInfo> GetValidationErrors(dynamic doc);
        string GetValidationErrorsJson(IEnumerable<ValidationErrorInfo> validationErrors);
    }
}
