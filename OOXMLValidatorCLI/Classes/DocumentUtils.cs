using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

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
            return PresentationDocument.Open(filePath, false);
        }
        public SpreadsheetDocument OpenSpreadsheetDocument(string filePath)
        {
            return SpreadsheetDocument.Open(filePath, false);
        }
        public Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> Validate(OpenXmlPackage doc, FileFormatVersions version)
        {
            OpenXmlValidator openXmlValidator = new OpenXmlValidator(version);
            bool isStrict = doc.StrictRelationshipFound;
            IEnumerable<ValidationErrorInfoInternal> errors = openXmlValidator.Validate(doc).Select(e => new ValidationErrorInfoInternal
            {
                ErrorType = e.ErrorType.ToString(),
                Description = e.Description,
                Path = e.Path.ToString(),
                Id = e.Id.ToString()
            });

            return new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(isStrict, errors);
        }
    }
}
