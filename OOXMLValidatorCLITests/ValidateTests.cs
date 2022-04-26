﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;

namespace OOXMLValidatorCLITests
{
    [TestClass]
    public class ValidateTests
    {
        [TestMethod]
        public void Validate_ShouldCallTheRightMethods()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock);
            string testPath = "path/to/a/file.docx";
            string testFormat = "Office2016";
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);

            memoryStream.Seek(0, SeekOrigin.Begin);

            IEnumerable<ValidationErrorInfo> validationErrorInfos = new List<ValidationErrorInfo>() { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };

            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), ".docx")).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>()))
                .Callback<dynamic>(o =>
                {
                    Assert.AreEqual(testWordDoc, o);
                })
                .Returns(new Tuple<bool, IEnumerable<ValidationErrorInfo>>(true, validationErrorInfos));

            validate.OOXML(testPath, testFormat, false);

            Mock.Get(functionUtilsMock).Verify(f => f.SetOfficeVersion(testFormat), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetDocument(testPath, ".docx"), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetValidationErrors(new Tuple<bool, IEnumerable<ValidationErrorInfo>>(true, validationErrorInfos), testPath, false), Times.Once());
        }
    }
}
