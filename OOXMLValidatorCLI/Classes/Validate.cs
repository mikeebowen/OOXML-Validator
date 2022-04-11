using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;

namespace OOXMLValidatorCLI.Classes
{
    public class Validate : IValidate
    {
        private readonly IFunctionUtils _functionUtils;

        public Validate(IFunctionUtils functionUtils)
        {
            _functionUtils = functionUtils;
        }
        public object OOXML(string filePath, string format, bool? returnXml)
        {
            _functionUtils.SetOfficeVersion(format);

            OpenXmlPackage doc = _functionUtils.GetDocument(filePath);

            Tuple<bool, IEnumerable<ValidationErrorInfo>> validationErrorInfos = _functionUtils.GetValidationErrors(doc);

            return _functionUtils.GetValidationErrors(validationErrorInfos, filePath, returnXml ?? false);
        }
    }
}
