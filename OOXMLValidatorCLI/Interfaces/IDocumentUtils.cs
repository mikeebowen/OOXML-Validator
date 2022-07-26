using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OOXMLValidatorCLI.Classes;
using System;
using System.Collections.Generic;

namespace OOXMLValidatorCLI.Interfaces
{
    public interface IDocumentUtils
    {
        WordprocessingDocument OpenWordprocessingDocument(string tempFilePath);
        SpreadsheetDocument OpenSpreadsheetDocument(string tempFilePath);
        PresentationDocument OpenPresentationDocument(string tempFilePath);
        Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> Validate(OpenXmlPackage doc, FileFormatVersions version);
    }
}
