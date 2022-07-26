using DocumentFormat.OpenXml.Office.CoverPageProps;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOXMLValidatorCLI.Classes
{
    public class ValidationErrorInfoInternal
    {
        public string ErrorType { get; set; }
        public string Description { get; set; }
        public string Path { get; set; }
        public string Id { get; set; }
    }
}
