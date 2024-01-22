// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace OOXMLValidatorCLITests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using OOXMLValidatorCLI.Classes;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Contains unit tests for the Validate class.
    /// </summary>
    [TestClass]
    public class ValidateTests
    {
        /// <summary>
        /// Validates a single file.
        /// </summary>
        [TestMethod]
        public void Validate_ShouldValidateASingleFile()
        {
            // Arrange
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/a/file.docx";
            string testFormat = "Office2016";
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);

            memoryStream.Seek(0, SeekOrigin.Begin);

            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };

            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), ".docx")).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>()))
                .Callback<dynamic>(o =>
                {
                    Assert.AreEqual(testWordDoc, o);
                })
                .Returns(new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(true, validationErrorInfos));

            // Act
            validate.OOXML(testPath, testFormat);

            // Assert
            Mock.Get(functionUtilsMock).Verify(f => f.SetOfficeVersion(testFormat), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetDocument(testPath, ".docx"), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetValidationErrorsData(new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(true, validationErrorInfos), testPath, false), Times.Once());
        }

        /// <summary>
        /// Validates that an exception is thrown with an invalid file type.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException), "file must be have one of these extensions: .docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, .xlam")]
        public void Validate_ShouldThrowAnExceptionWithInvalidFileType()
        {
            // Arrange
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/a/file.foo";
            string testFormat = "Office2016";

            // Act and Assert
            validate.OOXML(testPath, testFormat);
            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Normal);
        }

        /// <summary>
        /// Validates a folder and returns the result in JSON format.
        /// </summary>
        [TestMethod]
        public void Validate_ShouldValidateAFolderAndReturnJson()
        {
            // Arrange
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };

            memoryStream.Seek(0, SeekOrigin.Begin);

            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(false, validationErrorInfos));

            // Act
            object validationErrors = validate.OOXML(testPath, testFormat, false, false);

            // Assert
            Assert.IsNotNull(validationErrors);
            Assert.AreEqual(validationErrors, "[{\"FilePath\":\"taco.docx\",\"ValidationErrors\":null},{\"FilePath\":\"cat.pptx\",\"ValidationErrors\":null},{\"FilePath\":\"foo.xlsx\",\"ValidationErrors\":null},{\"FilePath\":\"bar.docm\",\"ValidationErrors\":null}]");
        }

        /// <summary>
        /// Validates a folder and returns the result in XML format.
        /// </summary>
        [TestMethod]
        public void Validate_ShouldValidateAFolderAndReturnXml()
        {
            // Arrange
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };
            memoryStream.Seek(0, SeekOrigin.Begin);

            XDocument testXDocument = new XDocument(new XElement("Document"));
            XElement testXml = new XElement("KittyErrorInfoList");
            testXml.Add(
                new XElement(
                    "KittyErrorInfo",
                    new XElement("Description", "The cat peed on the carpet"),
                    new XElement("Path", "office/carpet"),
                    new XElement("Id", "305"),
                    new XElement("ErrorType", "KittyPee")));

            for (int i = 0; i < 4; i++)
            {
                testXDocument.Root.Add(XElement.Parse(testXml.ToString()));
            }

            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(false, validationErrorInfos));
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrorsData(It.IsAny<Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(new XDocument(testXml));

            // Act
            object validationErrorXml = validate.OOXML(testPath, testFormat, true, false);

            // Assert
            Assert.IsNotNull(validationErrorXml);
            Assert.AreEqual(validationErrorXml.ToString(), testXDocument.ToString());
        }

        /// <summary>
        /// Validates a folder recursively with the flag and returns the result.
        /// </summary>
        [TestMethod]
        public void Validate_ShouldTestRecursivelyWithFlag()
        {
            // Arrange
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };

            memoryStream.Seek(0, SeekOrigin.Begin);

            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(directoryServiceMock).Setup(d => d.EnumerateFiles(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<SearchOption>())).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(false, validationErrorInfos));

            // Act
            object validationErrors = validate.OOXML(testPath, testFormat, false, true);

            // Assert
            Assert.IsNotNull(validationErrors);
            Assert.AreEqual(validationErrors, "[{\"FilePath\":\"taco.docx\",\"ValidationErrors\":null},{\"FilePath\":\"cat.pptx\",\"ValidationErrors\":null},{\"FilePath\":\"foo.xlsx\",\"ValidationErrors\":null},{\"FilePath\":\"bar.docm\",\"ValidationErrors\":null}]");
            Mock.Get(directoryServiceMock).Verify(d => d.EnumerateFiles(testPath, "*.*", SearchOption.AllDirectories), Times.Once);
            Mock.Get(directoryServiceMock).Verify(d => d.GetFiles(testPath), Times.Never);
        }

        /// <summary>
        /// Validates files without testing them recursively.
        /// </summary>
        [TestMethod]
        public void Validate_ShouldNotTestFilesRecursivelyWithoutFlag()
        {
            // Arrange
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfoInternal> validationErrorInfos = new List<ValidationErrorInfoInternal>() { new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal(), new ValidationErrorInfoInternal() };

            memoryStream.Seek(0, SeekOrigin.Begin);

            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(directoryServiceMock).Setup(d => d.EnumerateFiles(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<SearchOption>())).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfoInternal>>(false, validationErrorInfos));

            // Act
            object validationErrors = validate.OOXML(testPath, testFormat, false, false);

            // Assert
            Assert.IsNotNull(validationErrors);
            Assert.AreEqual(validationErrors, "[{\"FilePath\":\"taco.docx\",\"ValidationErrors\":null},{\"FilePath\":\"cat.pptx\",\"ValidationErrors\":null},{\"FilePath\":\"foo.xlsx\",\"ValidationErrors\":null},{\"FilePath\":\"bar.docm\",\"ValidationErrors\":null}]");
            Mock.Get(directoryServiceMock).Verify(d => d.EnumerateFiles(testPath, "*.*", SearchOption.AllDirectories), Times.Never);
            Mock.Get(directoryServiceMock).Verify(d => d.GetFiles(testPath), Times.Once);
        }
    }
}
