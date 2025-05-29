using System;
using System.IO;
using System.Linq;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxMerger.Tests
{
    public class TestDocumentCreatorTests : IDisposable
    {
        private readonly string _testDirectory;

        public TestDocumentCreatorTests()
        {
            _testDirectory = Path.Combine(Path.GetTempPath(), "TestDocumentCreatorTests_" + Guid.NewGuid().ToString("N")[..8]);
            Directory.CreateDirectory(_testDirectory);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDirectory))
            {
                try
                {
                    Directory.Delete(_testDirectory, true);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        [Fact]
        public void CreateTestDocument_WithValidParameters_ShouldCreateDocxFile()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "test.docx");
            var title = "Test Document";
            var content = "This is test content.";

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            Assert.True(File.Exists(filePath), "DOCX file should be created");
        }

        [Fact]
        public void CreateTestDocument_ShouldCreateValidDocxStructure()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "valid_structure.docx");
            var title = "Valid Structure Test";
            var content = "Content for structure validation.";

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            using var doc = WordprocessingDocument.Open(filePath, false);

            Assert.NotNull(doc.MainDocumentPart);
            Assert.NotNull(doc.MainDocumentPart.Document);
            Assert.NotNull(doc.MainDocumentPart.Document.Body);

            var body = doc.MainDocumentPart.Document.Body;
            var paragraphs = body.Elements<Paragraph>();

            Assert.True(paragraphs.Any(), "Document should contain paragraphs");
        }

        [Fact]
        public void CreateTestDocument_ShouldContainSpecifiedTitle()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "title_test.docx");
            var title = "My Custom Title";
            var content = "Some content here.";

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains(title, text);
        }

        [Fact]
        public void CreateTestDocument_ShouldContainSpecifiedContent()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "content_test.docx");
            var title = "Title Here";
            var content = "This is my specific content that should appear in the document.";

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains(content, text);
        }

        [Fact]
        public void CreateTestDocument_WithEmptyTitle_ShouldStillCreateValidDocument()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "empty_title.docx");
            var title = "";
            var content = "Content without title.";

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            Assert.True(File.Exists(filePath));

            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains(content, text);
        }

        [Fact]
        public void CreateTestDocument_WithEmptyContent_ShouldStillCreateValidDocument()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "empty_content.docx");
            var title = "Title Only";
            var content = "";

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            Assert.True(File.Exists(filePath));

            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains(title, text);
        }

        [Fact]
        public void CreateTestDocument_WithSpecialCharacters_ShouldHandleCorrectly()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "special_chars.docx");
            var title = "Special Characters: áéíóú ñ @#$%^&*()";
            var content = "Content with symbols: ©®™ symbols and text";

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            Assert.True(File.Exists(filePath));

            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains("Special Characters", text);
            Assert.Contains("Content with symbols", text);
        }

        [Fact]
        public void CreateTestDocument_WithLongContent_ShouldCreateSuccessfully()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "long_content.docx");
            var title = "Long Content Test";
            var content = string.Join(" ", Enumerable.Repeat("This is a long paragraph with lots of text.", 100));

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            Assert.True(File.Exists(filePath));

            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains(title, text);
            Assert.Contains("long paragraph", text);
            Assert.True(text.Length > 1000, "Document should contain long content");
        }

        [Fact]
        public void CreateTestDocument_WithNullFilePath_ShouldThrowException()
        {
            // Arrange
            var title = "Test";
            var content = "Content";

            // Act & Assert
            Assert.ThrowsAny<Exception>(() => TestDocumentCreator.CreateTestDocument(null, title, content));
        }

        [Fact]
        public void CreateTestDocument_WithInvalidPath_ShouldThrowException()
        {
            // Arrange
            var filePath = "/nonexistent/directory/test.docx";
            var title = "Test";
            var content = "Content";

            // Act & Assert
            Assert.ThrowsAny<Exception>(() => TestDocumentCreator.CreateTestDocument(filePath, title, content));
        }

        [Fact]
        public void CreateTestDocument_OverwriteExistingFile_ShouldSucceed()
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, "overwrite.docx");
            var title1 = "First Title";
            var content1 = "First content";
            var title2 = "Second Title";
            var content2 = "Second content";

            // Act - Create first document
            TestDocumentCreator.CreateTestDocument(filePath, title1, content1);
            Assert.True(File.Exists(filePath));

            // Act - Overwrite with second document
            TestDocumentCreator.CreateTestDocument(filePath, title2, content2);

            // Assert
            Assert.True(File.Exists(filePath));

            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains(title2, text);
            Assert.Contains(content2, text);
            Assert.DoesNotContain(title1, text);
            Assert.DoesNotContain(content1, text);
        }

        [Theory]
        [InlineData("test1.docx", "Document 1", "Content 1")]
        [InlineData("test2.docx", "Document 2", "Content 2")]
        [InlineData("test3.docx", "Document 3", "Content 3")]
        public void CreateTestDocument_WithVariousInputs_ShouldCreateCorrectDocuments(string fileName, string title, string content)
        {
            // Arrange
            var filePath = Path.Combine(_testDirectory, fileName);

            // Act
            TestDocumentCreator.CreateTestDocument(filePath, title, content);

            // Assert
            Assert.True(File.Exists(filePath));

            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;

            Assert.Contains(title, text);
            Assert.Contains(content, text);
        }
    }
}
