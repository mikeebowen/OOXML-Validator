using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OOXMLValidatorCLITests
{
    [TestClass]
    public class FunctionUtilsTests
    {
        [TestMethod]
        public void ShouldSetOfficeVersion()
        {
            var documentMock = Mock.Of<IDocumentUtils>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            Assert.AreEqual(functionUtils.OfficeVersion, Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Max());
        }

        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodWord()
        {
            string testPath = "foo/bar/baz.docx";
            var documentUtilsMock = Mock.Of<IDocumentUtils>();
            MemoryStream memoryStream = new MemoryStream();

            memoryStream.Seek(0, SeekOrigin.Begin);

            using WordprocessingDocument testDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);

            Mock.Get(documentUtilsMock)
                .Setup(d => d.OpenWordprocessingDocument(It.IsAny<string>()))
                .Callback<string>(s =>
                {
                    Assert.AreEqual(testPath, s);
                })
                .Returns(testDocument);
            FunctionUtils functionUtils = new FunctionUtils(documentUtilsMock);

            var res = functionUtils.GetDocument(testPath);

            Assert.AreEqual(res, testDocument);

            Mock.Get(documentUtilsMock).Verify(f => f.OpenWordprocessingDocument(testPath), Times.Once);
        }

        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodPresentation()
        {
            string testPath = "foo/bar/baz.pptx";
            var documentMock = Mock.Of<IDocumentUtils>();
            MemoryStream memoryStream = new MemoryStream();

            memoryStream.Seek(0, SeekOrigin.Begin);

            using PresentationDocument testDocument = PresentationDocument.Create(memoryStream, PresentationDocumentType.Presentation);

            Mock.Get(documentMock)
                .Setup(d => d.OpenPresentationDocument(It.IsAny<string>()))
                .Callback<string>(s =>
                {
                    Assert.AreEqual(testPath, s);
                })
                .Returns(testDocument);
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            var res = functionUtils.GetDocument(testPath);

            Assert.AreEqual(res, testDocument);

            Mock.Get(documentMock).Verify(f => f.OpenPresentationDocument(testPath), Times.Once);
        }

        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodSpreadsheet()
        {
            string testPath = "foo/bar/baz.xlsx";
            var documentMock = Mock.Of<IDocumentUtils>();
            MemoryStream memoryStream = new MemoryStream();
            using SpreadsheetDocument testDynamic = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);

            memoryStream.Seek(0, SeekOrigin.Begin);

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
        public void SetOfficeVersion_ShouldSetValidVersion()
        {
            var documentMock = Mock.Of<IDocumentUtils>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            functionUtils.SetOfficeVersion("Office2016");

            Assert.AreEqual(functionUtils.OfficeVersion, FileFormatVersions.Office2016);
        }

        [TestMethod]
        public void SetOfficeVersion_ShouldSetDefaultVersion()
        {
            var documentMock = Mock.Of<IDocumentUtils>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            functionUtils.SetOfficeVersion(null);

            Assert.AreEqual(functionUtils.OfficeVersion, Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Last());
        }

        [TestMethod]
        public void GetValidationErrors_ShouldReturnValidJson()
        {
            IEnumerable<ValidationErrorInfo> validationErrorInfos = new List<ValidationErrorInfo>() { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };
            var documentMock = Mock.Of<IDocumentUtils>();
            string testJson = "\"[{\\\"Description\\\":\\\"\\\",\\\"Path\\\":null,\\\"Id\\\":null,\\\"ErrorType\\\":0},{\\\"Description\\\":\\\"\\\",\\\"Path\\\":null,\\\"Id\\\":null,\\\"ErrorType\\\":0},{\\\"Description\\\":\\\"\\\",\\\"Path\\\":null,\\\"Id\\\":null,\\\"ErrorType\\\":0}]\"";


            var functionUtils = new FunctionUtils(documentMock);

            object res = functionUtils.GetValidationErrors(Tuple.Create(true, validationErrorInfos), @"C:\test\file\path.xlsx", false);
            string jsonData = JsonConvert.SerializeObject(res);
            Assert.AreEqual(jsonData, testJson);
        }

        [TestMethod]
        public void GetValidationErrors_ShouldReturnValidXml()
        {
            IEnumerable<ValidationErrorInfo> validationErrorInfos = new List<ValidationErrorInfo>() { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };
            var documentMock = Mock.Of<IDocumentUtils>();
            string xmlString = "<ValidationErrorInfoList FilePath='C:\\test\\file\\path.xlsx' IsStrict='true'><ValidationErrorInfo><Description></Description><Path/><Id/><ErrorType>Schema</ErrorType></ValidationErrorInfo><ValidationErrorInfo><Description></Description><Path/><Id/><ErrorType>Schema</ErrorType></ValidationErrorInfo><ValidationErrorInfo><Description></Description><Path/><Id/><ErrorType>Schema</ErrorType></ValidationErrorInfo></ValidationErrorInfoList>";
            XDocument xDoc = XDocument.Parse(xmlString);

            var functionUtils = new FunctionUtils(documentMock);
            object res = functionUtils.GetValidationErrors(Tuple.Create(true, validationErrorInfos), @"C:\test\file\path.xlsx", true);

            Assert.IsTrue(XNode.DeepEquals((res as XDocument), xDoc));
        }
    }
}
