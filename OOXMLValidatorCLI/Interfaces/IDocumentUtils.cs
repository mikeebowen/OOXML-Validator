using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OOXMLValidatorCLI.Interfaces
{
    public interface IDocumentUtils
    {
        WordprocessingDocument OpenWordprocessingDocument(string tempFilePath);
        SpreadsheetDocument OpenSpreadsheetDocument(string tempFilePath);
        PresentationDocument OpenPresentationDocument(string tempFilePath);
        Tuple<bool, IEnumerable<ValidationErrorInfo>> Validate(OpenXmlPackage doc, FileFormatVersions version);
    }
}
