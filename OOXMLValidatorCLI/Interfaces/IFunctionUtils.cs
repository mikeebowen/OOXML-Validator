using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
namespace OOXMLValidatorCLI.Interfaces
{
    public interface IFunctionUtils
    {
        FileFormatVersions OfficeVersion { get; }
        void SetOfficeVersion(string version);
        OpenXmlPackage GetDocument(string filePath);
        Tuple<bool, IEnumerable<ValidationErrorInfo>> GetValidationErrors(OpenXmlPackage doc);
        object GetValidationErrors(Tuple<bool, IEnumerable<ValidationErrorInfo>> data, string filePath, bool returnXml);
    }
}
