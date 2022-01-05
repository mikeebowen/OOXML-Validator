using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Text;

namespace OOXMLValidatorCLITests
{
    [TestClass]
    public class ValidateTests
    {
        [TestMethod]
        public void Validate_ShouldCallTheRightMethods()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var validate = new Validate(functionUtilsMock);
            string testPath = "path/to/a/file.docx";
            string testFormat = "Office2016";
            dynamic testDynamic = new ExpandoObject();
            testDynamic.Foo = "Bar";
            var validationErrorInfos = new ValidationErrorInfo[] { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };

            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>())).Returns(testDynamic);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<object>()))
                .Callback<dynamic>(o =>
                {
                    Assert.AreEqual(testDynamic, o);
                })
                .Returns(validationErrorInfos);

            validate.OOXML(testPath, testFormat);

            Mock.Get(functionUtilsMock).Verify(f => f.SetOfficeVersion(testFormat), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetDocument(testPath), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetValidationErrorsJson(validationErrorInfos), Times.Once());
        }
    }
}
