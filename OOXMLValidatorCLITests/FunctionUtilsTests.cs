using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Dynamic;
using System.Linq;

namespace OOXMLValidatorCLITests
{
    [TestClass]
    public class FunctionUtilsTests
    {
        [TestMethod]
        public void ShouldSetOfficeVersion()
        {
            var documentMock = Mock.Of<IDocument>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            Assert.AreEqual(functionUtils.OfficeVersion, Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Max());
        }

        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodWord()
        {
            string testPath = "foo/bar/baz.docx";
            var documentMock = Mock.Of<IDocument>();
            dynamic testDynamic = new ExpandoObject();
            testDynamic.Foo = "Bar";

            Mock.Get(documentMock)
                .Setup(d => d.OpenWordprocessingDocument(It.IsAny<string>()))
                .Callback<string>(s =>
                {
                    Assert.AreEqual(testPath, s);
                })
                .Returns(testDynamic);
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            var res = functionUtils.GetDocument(testPath);

            Assert.AreEqual(res, testDynamic);

            Mock.Get(documentMock).Verify(f => f.OpenWordprocessingDocument(testPath), Times.Once);
        }

        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodPresentation()
        {
            string testPath = "foo/bar/baz.pptx";
            var documentMock = Mock.Of<IDocument>();
            dynamic testDynamic = new ExpandoObject();
            testDynamic.Foo = "Bar";

            Mock.Get(documentMock)
                .Setup(d => d.OpenPresentationDocument(It.IsAny<string>()))
                .Callback<string>(s =>
                {
                    Assert.AreEqual(testPath, s);
                })
                .Returns(testDynamic);
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            var res = functionUtils.GetDocument(testPath);

            Assert.AreEqual(res, testDynamic);

            Mock.Get(documentMock).Verify(f => f.OpenPresentationDocument(testPath), Times.Once);
        }

        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodSpreadsheet()
        {
            string testPath = "foo/bar/baz.xlsx";
            var documentMock = Mock.Of<IDocument>();
            dynamic testDynamic = new ExpandoObject();
            testDynamic.Foo = "Bar";

            Mock.Get(documentMock)
                .Setup(d => d.OpenSpreadsheetDocument(It.IsAny<string>()))
                .Callback<string>(s =>
                {
                    Assert.AreEqual(testPath, s);
                })
                .Returns(testDynamic);
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            var res = functionUtils.GetDocument(testPath);

            Assert.AreEqual(res, testDynamic);

            Mock.Get(documentMock).Verify(f => f.OpenSpreadsheetDocument(testPath), Times.Once);
        }

        [TestMethod]
        public void SetOfficeVersion_ShouldThrowAnErrorWithInvalidVersion()
        {
            var documentMock = Mock.Of<IDocument>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            Assert.ThrowsException<ArgumentOutOfRangeException>(() => functionUtils.SetOfficeVersion(99));
        }

        [TestMethod]
        public void SetOfficeVersion_ShouldSetValidVersion()
        {
            var documentMock = Mock.Of<IDocument>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            functionUtils.SetOfficeVersion(8);

            Assert.AreEqual(functionUtils.OfficeVersion, FileFormatVersions.Office2016);
        }

        [TestMethod]
        public void SetOfficeVersion_ShouldSetDefaultVersion()
        {
            var documentMock = Mock.Of<IDocument>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            functionUtils.SetOfficeVersion(null);

            Assert.AreEqual(functionUtils.OfficeVersion, Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Last());
        }

        [TestMethod]
        public void GetValidationErrorsJson_ShouldReturnValidJson()
        {
            var validationErrorInfos = new ValidationErrorInfo[] { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };
            var documentMock = Mock.Of<IDocument>();
            string testJson = "[{\"Description\":\"\",\"Path\":null,\"Id\":null,\"ErrorType\":0},{\"Description\":\"\",\"Path\":null,\"Id\":null,\"ErrorType\":0},{\"Description\":\"\",\"Path\":null,\"Id\":null,\"ErrorType\":0}]";


            var functionUtils = new FunctionUtils(documentMock);

            string res = functionUtils.GetValidationErrorsJson(validationErrorInfos);
            Assert.AreEqual(res, testJson);
        }
    }
}
