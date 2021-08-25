using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Text;

namespace OOXMLValidator.Interfaces
{
    public interface IDocument
    {
        dynamic OpenWordprocessingDocument(string tempFilePath);
        dynamic OpenSpreadsheetDocument(string tempFilePath);
        dynamic OpenPresentationDocument(string tempFilePath);
        IEnumerable<ValidationErrorInfo> Validate(dynamic doc, FileFormatVersions version);
    }
}
