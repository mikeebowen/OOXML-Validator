using DocumentFormat.OpenXml.Validation;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace OOXMLValidatorCLI.Classes
{
    public class Validate : IValidate
    {
        private readonly IFunctionUtils _functionUtils;

        public Validate(IFunctionUtils functionUtils)
        {
            _functionUtils = functionUtils;
        }
        public string OOXML(string filePath, string format)
        {
            _functionUtils.SetOfficeVersion(format);

            dynamic doc = _functionUtils.GetDocument(filePath);

            IEnumerable<ValidationErrorInfo> validationErrorInfos = _functionUtils.GetValidationErrors(doc);

            return _functionUtils.GetValidationErrorsJson(validationErrorInfos);
        }
    }
}
