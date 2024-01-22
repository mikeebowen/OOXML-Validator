// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLITests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using Newtonsoft.Json;
    using OOXMLValidatorCLI.Classes;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Unit tests for the FunctionUtils class.
    /// </summary>
    [TestClass]
    public class FunctionUtilsTests
    {
        /// <summary>
        /// Test case to verify that the OfficeVersion property is set correctly.
        /// </summary>
        [TestMethod]
        public void ShouldSetOfficeVersion()
        {
            // Arrange
            var documentMock = Mock.Of<IDocumentUtils>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            // Act
            var officeVersion = functionUtils.OfficeVersion;

            // Assert
            Assert.AreEqual(officeVersion, Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Max());
        }

        /// <summary>
        /// Test case to verify that the GetDocument method calls the correct Open method for Word documents.
        /// </summary>
        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodWord()
        {
            // Arrange
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

            // Act
            var res = functionUtils.GetDocument(testPath, ".docx");

            // Assert
            Assert.AreEqual(res, testDocument);

            Mock.Get(documentUtilsMock).Verify(f => f.OpenWordprocessingDocument(testPath), Times.Once);
        }

        /// <summary>
        /// Test case to verify that the GetDocument method calls the correct Open method for Presentation documents.
        /// </summary>
        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodPresentation()
        {
            // Arrange
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

            // Act
            var res = functionUtils.GetDocument(testPath, ".pptx");

            // Assert
            Assert.AreEqual(res, testDocument);

            Mock.Get(documentMock).Verify(f => f.OpenPresentationDocument(testPath), Times.Once);
        }

        /// <summary>
        /// Test case to verify that the GetDocument method calls the correct Open method for Spreadsheet documents.
        /// </summary>
        [TestMethod]
        public void GetDocument_ShouldCallCorrectOpenMethodSpreadsheet()
        {
            // Arrange
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

            // Act
            var res = functionUtils.GetDocument(testPath, ".xlsx");

            // Assert
            Assert.AreEqual(res, testDynamic);

            Mock.Get(documentMock).Verify(f => f.OpenSpreadsheetDocument(testPath), Times.Once);
        }

        /// <summary>
        /// Test case to verify that the SetOfficeVersion method sets the valid version.
        /// </summary>
        [TestMethod]
        public void SetOfficeVersion_ShouldSetValidVersion()
        {
            // Arrange
            var documentMock = Mock.Of<IDocumentUtils>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            // Act
            functionUtils.SetOfficeVersion("Office2016");

            // Assert
            Assert.AreEqual(functionUtils.OfficeVersion, FileFormatVersions.Office2016);
        }

        /// <summary>
        /// Test case to verify that the SetOfficeVersion method sets the default version.
        /// </summary>
        [TestMethod]
        public void SetOfficeVersion_ShouldSetDefaultVersion()
        {
            // Arrange
            var documentMock = Mock.Of<IDocumentUtils>();
            FunctionUtils functionUtils = new FunctionUtils(documentMock);

            // Act
            functionUtils.SetOfficeVersion(null);

            // Assert
            Assert.AreEqual(functionUtils.OfficeVersion, Enum.GetValues(typeof(FileFormatVersions)).Cast<FileFormatVersions>().Last());
        }

        /// <summary>
        /// Test case to verify that the GetValidationErrorsData method returns valid JSON.
        /// </summary>
        [TestMethod]
        public void GetValidationErrorsData_ShouldReturnValidJson()
        {
            // Arrange
            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };
            var documentMock = Mock.Of<IDocumentUtils>();
            string testJson = "\"[{\\\"Description\\\":null,\\\"Path\\\":null,\\\"Id\\\":null,\\\"ErrorType\\\":null},{\\\"Description\\\":null,\\\"Path\\\":null,\\\"Id\\\":null,\\\"ErrorType\\\":null},{\\\"Description\\\":null,\\\"Path\\\":null,\\\"Id\\\":null,\\\"ErrorType\\\":null}]\"";

            var functionUtils = new FunctionUtils(documentMock);

            // Act
            object res = functionUtils.GetValidationErrorsData(Tuple.Create(true, validationErrorInfos), @"C:\test\file\path.xlsx", false);
            string jsonData = JsonConvert.SerializeObject(res);

            // Assert
            Assert.AreEqual(jsonData, testJson);
        }

        /// <summary>
        /// Test case to verify that the GetValidationErrorsData method returns valid XML.
        /// </summary>
        [TestMethod]
        public void GetValidationErrorsData_ShouldReturnValidXml()
        {
            // Arrange
            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };
            var documentMock = Mock.Of<IDocumentUtils>();
            string xmlString = "<File FilePath=\"C:\\test\\file\\path.xlsx\" IsStrict=\"true\"><ValidationErrorInfoList><ValidationErrorInfo><Description /><Path /><Id /><ErrorType /></ValidationErrorInfo><ValidationErrorInfo><Description /><Path /><Id /><ErrorType /></ValidationErrorInfo><ValidationErrorInfo><Description /><Path /><Id /><ErrorType /></ValidationErrorInfo></ValidationErrorInfoList></File>";
            XDocument xDoc = XDocument.Parse(xmlString);

            var functionUtils = new FunctionUtils(documentMock);

            // Act
            object res = functionUtils.GetValidationErrorsData(Tuple.Create(true, validationErrorInfos), @"C:\test\file\path.xlsx", true);

            // Assert
            Assert.IsTrue(XNode.DeepEquals(res as XDocument, xDoc));
        }

        /// <summary>
        /// Test case to verify that the GetValidationErrors method calls the Validate method.
        /// </summary>
        [TestMethod]
        public void GetValidationErrors_ShouldCallValidate()
        {
            // Arrange
            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };
            var testTup = Tuple.Create(true, validationErrorInfos);
            var documentMock = Mock.Of<IDocumentUtils>();

            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);

            Mock.Get(documentMock).Setup(d => d.Validate(It.IsAny<OpenXmlPackage>(), It.IsAny<FileFormatVersions>())).Returns(testTup);

            var functionUtils = new FunctionUtils(documentMock);

            // Act
            var resTup = functionUtils.GetValidationErrors(testWordDoc);

            // Assert
            Assert.AreEqual(testTup, resTup);
            Mock.Get(documentMock).Verify(d => d.Validate(It.IsAny<OpenXmlPackage>(), It.IsAny<FileFormatVersions>()), Times.Once());
        }
    }
}
