using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OOXMLValidatorCLI.Interfaces;

namespace OOXMLValidatorCLI.Classes
{
    public class DocumentUtils : IDocumentUtils
    {
        public WordprocessingDocument OpenWordprocessingDocument(string filePath)
        {
            return WordprocessingDocument.Open(filePath, false);
        }
        public PresentationDocument OpenPresentationDocument(string filePath)
        {
            var foo = PresentationDocument.Open(filePath, false);
            var bar = foo.StrictRelationshipFound;
            return foo;
        }
        public SpreadsheetDocument OpenSpreadsheetDocument(string filePath)
        {
            return SpreadsheetDocument.Open(filePath, false);
        }
        public Tuple<bool, IEnumerable<ValidationErrorInfo>> Validate(OpenXmlPackage doc, FileFormatVersions version)
        {
            OpenXmlValidator openXmlValidator = new OpenXmlValidator(version);
            bool isStrict = doc.StrictRelationshipFound;
            IEnumerable<ValidationErrorInfo> errors = openXmlValidator.Validate(doc);
            return new Tuple<bool, IEnumerable<ValidationErrorInfo>>(isStrict, errors);
        }
    }
}
