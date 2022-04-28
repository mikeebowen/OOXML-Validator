using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace OOXMLValidatorCLITests
{
    [TestClass]
    public class ValidateTests
    {
        [TestMethod]
        public void Validate_ShouldValidateASingleFile()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
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

            validate.OOXML(testPath, testFormat);

            Mock.Get(functionUtilsMock).Verify(f => f.SetOfficeVersion(testFormat), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetDocument(testPath, ".docx"), Times.Once());
            Mock.Get(functionUtilsMock).Verify(f => f.GetValidationErrorsData(new Tuple<bool, IEnumerable<ValidationErrorInfo>>(true, validationErrorInfos), testPath, false), Times.Once());
        }
        // file must be have one of these extensions: .docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, .xlam
        [TestMethod]
        [ExpectedException(typeof(ArgumentException), "file must be have one of these extensions: .docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, .xlam")]
        public void Validate_ShouldThrowAnExceptionWithInvalidFileType()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/a/file.foo";
            string testFormat = "Office2016";

            validate.OOXML(testPath, testFormat);
            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Normal);
        }

        [TestMethod]
        public void Validate_ShouldValidateAFolderAndReturnJson()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfo> validationErrorInfos = new List<ValidationErrorInfo>() { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };

            memoryStream.Seek(0, SeekOrigin.Begin);

            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfo>>(false, validationErrorInfos));

            object validationErrors = validate.OOXML(testPath, testFormat, false, false);
            Assert.IsNotNull(validationErrors);
            Assert.AreEqual(validationErrors, "[{\"FilePath\":\"taco.docx\",\"ValidationErrors\":null},{\"FilePath\":\"cat.pptx\",\"ValidationErrors\":null},{\"FilePath\":\"foo.xlsx\",\"ValidationErrors\":null},{\"FilePath\":\"bar.docm\",\"ValidationErrors\":null}]");
        }


        [TestMethod]
        public void Validate_ShouldValidateAFolderAndReturnXml()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfo> validationErrorInfos = new List<ValidationErrorInfo>() { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };
            memoryStream.Seek(0, SeekOrigin.Begin);

            XDocument testXDocument = new XDocument(new XElement("Document"));
            XElement testXml = new XElement("KittyErrorInfoList");
            testXml.Add(
                        new XElement("KittyErrorInfo",
                            new XElement("Description", "The cat peed on the carpet"),
                            new XElement("Path", "office/carpet"),
                            new XElement("Id", "305"),
                            new XElement("ErrorType", "KittyPee")
                        )
                    );

            for (int i = 0; i < 4; i++)
            {
                testXDocument.Root.Add(XElement.Parse(testXml.ToString()));
            }


            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfo>>(false, validationErrorInfos));
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrorsData(It.IsAny<Tuple<bool, IEnumerable<ValidationErrorInfo>>>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(new XDocument(testXml));

            object validationErrorXml = validate.OOXML(testPath, testFormat, true, false);

            Assert.IsNotNull(validationErrorXml);
            Assert.AreEqual(validationErrorXml.ToString(), testXDocument.ToString());
        }

        [TestMethod]
        public void Validate_ShouldTestRecursivelyWithFlag()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfo> validationErrorInfos = new List<ValidationErrorInfo>() { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };

            memoryStream.Seek(0, SeekOrigin.Begin);

            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(directoryServiceMock).Setup(d => d.EnumerateFiles(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<SearchOption>())).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfo>>(false, validationErrorInfos));

            object validationErrors = validate.OOXML(testPath, testFormat, false, true);
            Assert.IsNotNull(validationErrors);
            Assert.AreEqual(validationErrors, "[{\"FilePath\":\"taco.docx\",\"ValidationErrors\":null},{\"FilePath\":\"cat.pptx\",\"ValidationErrors\":null},{\"FilePath\":\"foo.xlsx\",\"ValidationErrors\":null},{\"FilePath\":\"bar.docm\",\"ValidationErrors\":null}]");
            Mock.Get(directoryServiceMock).Verify(d => d.EnumerateFiles(testPath, "*.*", SearchOption.AllDirectories), Times.Once);
            Mock.Get(directoryServiceMock).Verify(d => d.GetFiles(testPath), Times.Never);
        }

        [TestMethod]
        public void Validate_ShouldNotTestFilesRecursivelyWithouFlag()
        {
            var functionUtilsMock = Mock.Of<IFunctionUtils>();
            var fileServiceMock = Mock.Of<IFileService>();
            var directoryServiceMock = Mock.Of<IDirectoryService>();
            var validate = new Validate(functionUtilsMock, fileServiceMock, directoryServiceMock);
            string testPath = "path/to/files/";
            string testFormat = null;
            MemoryStream memoryStream = new MemoryStream();
            using WordprocessingDocument testWordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document);
            IEnumerable<ValidationErrorInfo> validationErrorInfos = new List<ValidationErrorInfo>() { new ValidationErrorInfo(), new ValidationErrorInfo(), new ValidationErrorInfo() };

            memoryStream.Seek(0, SeekOrigin.Begin);

            Mock.Get(fileServiceMock).Setup(f => f.GetAttributes(testPath)).Returns(FileAttributes.Directory);
            Mock.Get(directoryServiceMock).Setup(d => d.GetFiles(testPath)).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(directoryServiceMock).Setup(d => d.EnumerateFiles(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<SearchOption>())).Returns(new string[] { "taco.docx", "cat.pptx", "foo.xlsx", "bar.docm" });
            Mock.Get(functionUtilsMock).Setup(f => f.GetDocument(It.IsAny<string>(), It.IsAny<string>())).Returns(testWordDoc);
            Mock.Get(functionUtilsMock).Setup(f => f.GetValidationErrors(It.IsAny<OpenXmlPackage>())).Returns(new Tuple<bool, IEnumerable<ValidationErrorInfo>>(false, validationErrorInfos));

            object validationErrors = validate.OOXML(testPath, testFormat, false, false);
            Assert.IsNotNull(validationErrors);
            Assert.AreEqual(validationErrors, "[{\"FilePath\":\"taco.docx\",\"ValidationErrors\":null},{\"FilePath\":\"cat.pptx\",\"ValidationErrors\":null},{\"FilePath\":\"foo.xlsx\",\"ValidationErrors\":null},{\"FilePath\":\"bar.docm\",\"ValidationErrors\":null}]");
            Mock.Get(directoryServiceMock).Verify(d => d.EnumerateFiles(testPath, "*.*", SearchOption.AllDirectories), Times.Never);
            Mock.Get(directoryServiceMock).Verify(d => d.GetFiles(testPath), Times.Once);
        }
    }
}
