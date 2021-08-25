using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace OOXMLValidator.Interfaces
{
    public interface IFunctionUtils
    {
        FileFormatVersions OfficeVersion { get; }
        void SetOfficeVersion(int? version);
        dynamic GetDocument(string filePath);
        IEnumerable<ValidationErrorInfo> GetValidationErrors(dynamic doc);
        string GetValidationErrorsJson(IEnumerable<ValidationErrorInfo> validationErrors);
    }
}
