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

            IEnumerable<ValidationErrorInfo> validationErrorInfos = openXmlValidator.Validate(doc);
            IEnumerable<ValidationErrorInfoInternal> errors = validationErrorInfos.Select(e => new ValidationErrorInfoInternal()
            {
                ErrorType = Enum.GetName(e.ErrorType),
                Description = e.Description,
                Path = e.Path,
                Id = e.Id,
            });

            return new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(isStrict, errors);
        }
    }
}
