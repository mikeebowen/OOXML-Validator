using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OOXMLValidator.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace OOXMLValidator.Classes
{
    public class Document : IDocument
    {
        public dynamic OpenWordprocessingDocument(string filePath)
        {
            return WordprocessingDocument.Open(filePath, false);
        }
        public dynamic OpenPresentationDocument(string filePath)
        {
            return PresentationDocument.Open(filePath, false);
        }
        public dynamic OpenSpreadsheetDocument(string filePath)
        {
            return SpreadsheetDocument.Open(filePath, false);
        }
        public IEnumerable<ValidationErrorInfo> Validate(dynamic doc, FileFormatVersions version)
        {
            OpenXmlValidator openXmlValidator = new OpenXmlValidator(version);
            return openXmlValidator.Validate(doc);
        }
    }
}
