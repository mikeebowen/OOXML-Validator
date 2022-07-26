﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OOXMLValidatorCLI.Classes;
using System;
using System.Collections.Generic;
namespace OOXMLValidatorCLI.Interfaces
{
    public interface IFunctionUtils
    {
        FileFormatVersions OfficeVersion { get; }
        void SetOfficeVersion(string version);
        OpenXmlPackage GetDocument(string filePath, string fileExtension);
        Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> GetValidationErrors(OpenXmlPackage doc);
        object GetValidationErrorsData(Tuple<bool, IEnumerable<ValidationErrorInfoInternal>> data, string filePath, bool returnXml);
    }
}
