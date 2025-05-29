using System;
using System.IO;
using System.Collections.Generic;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxMerger.Tests
{
    public class DocumentMergerTests : IDisposable
    {
        private readonly string _testDirectory;
        private readonly List<string> _testFiles;

        public DocumentMergerTests()
        {
            _testDirectory = Path.Combine(Path.GetTempPath(), "DocxMergerTests_" + Guid.NewGuid().ToString("N")[..8]);
            Directory.CreateDirectory(_testDirectory);
            _testFiles = new List<string>();
        }

        public void Dispose()
        {
            // Clean up test files
            foreach (var file in _testFiles)
            {
                if (File.Exists(file))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }

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

        private string CreateTestFile(string fileName, string title, string content)
        {
            var filePath = Path.Combine(_testDirectory, fileName);
            TestDocumentCreator.CreateTestDocument(filePath, title, content);
            _testFiles.Add(filePath);
            return filePath;
        }

        [Fact]
        public void MergeDocuments_WithValidFiles_ShouldCreateMergedDocument()
        {
            // Arrange
            var file1 = CreateTestFile("test1.docx", "Document 1", "This is the content of document 1.");
            var file2 = CreateTestFile("test2.docx", "Document 2", "This is the content of document 2.");
            var outputFile = Path.Combine(_testDirectory, "merged.docx");
            _testFiles.Add(outputFile);

            var inputFiles = new List<string> { file1, file2 };

            // Act
            DocumentMerger.MergeDocuments(inputFiles.ToArray(), outputFile);

            // Assert
            Assert.True(File.Exists(outputFile), "Merged document should be created");

            // Verify the merged document contains content from both files
            using var doc = WordprocessingDocument.Open(outputFile, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;
            Assert.Contains("Document 1", text);
            Assert.Contains("Document 2", text);
            Assert.Contains("content of document 1", text);
            Assert.Contains("content of document 2", text);
        }

        [Fact]
        public void MergeDocuments_WithSingleFile_ShouldCreateCopyOfDocument()
        {
            // Arrange
            var file1 = CreateTestFile("single.docx", "Single Document", "This is a single document.");
            var outputFile = Path.Combine(_testDirectory, "single_merged.docx");
            _testFiles.Add(outputFile);

            var inputFiles = new List<string> { file1 };

            // Act
            DocumentMerger.MergeDocuments(inputFiles.ToArray(), outputFile);

            // Assert
            Assert.True(File.Exists(outputFile), "Merged document should be created");

            using var doc = WordprocessingDocument.Open(outputFile, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;
            Assert.Contains("Single Document", text);
            Assert.Contains("single document", text);
        }

        [Fact]
        public void MergeDocuments_WithMultipleFiles_ShouldMergeAllDocuments()
        {
            // Arrange
            var file1 = CreateTestFile("multi1.docx", "Doc 1", "First document content.");
            var file2 = CreateTestFile("multi2.docx", "Doc 2", "Second document content.");
            var file3 = CreateTestFile("multi3.docx", "Doc 3", "Third document content.");
            var outputFile = Path.Combine(_testDirectory, "multi_merged.docx");
            _testFiles.Add(outputFile);

            var inputFiles = new List<string> { file1, file2, file3 };

            // Act
            DocumentMerger.MergeDocuments(inputFiles.ToArray(), outputFile);

            // Assert
            Assert.True(File.Exists(outputFile), "Merged document should be created");

            using var doc = WordprocessingDocument.Open(outputFile, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;
            Assert.Contains("Doc 1", text);
            Assert.Contains("Doc 2", text);
            Assert.Contains("Doc 3", text);
            Assert.Contains("First document", text);
            Assert.Contains("Second document", text);
            Assert.Contains("Third document", text);
        }

        [Fact]
        public void MergeDocuments_WithNullInputFiles_ShouldThrowArgumentException()
        {
            // Arrange
            var outputFile = Path.Combine(_testDirectory, "output.docx");

            // Act & Assert
            Assert.Throws<ArgumentException>(() => DocumentMerger.MergeDocuments(null!, outputFile));
        }

        [Fact]
        public void MergeDocuments_WithEmptyInputFiles_ShouldThrowArgumentException()
        {
            // Arrange
            var inputFiles = new List<string>();
            var outputFile = Path.Combine(_testDirectory, "output.docx");

            // Act & Assert
            Assert.Throws<ArgumentException>(() => DocumentMerger.MergeDocuments(inputFiles.ToArray(), outputFile));
        }

        [Fact]
        public void MergeDocuments_WithNullOutputPath_ShouldThrowArgumentException()
        {
            // Arrange
            var file1 = CreateTestFile("test.docx", "Test", "Test content");
            var inputFiles = new List<string> { file1 };

            // Act & Assert
            Assert.Throws<ArgumentException>(() => DocumentMerger.MergeDocuments(inputFiles.ToArray(), null!));
        }

        [Fact]
        public void MergeDocuments_WithEmptyOutputPath_ShouldThrowArgumentException()
        {
            // Arrange
            var file1 = CreateTestFile("test.docx", "Test", "Test content");
            var inputFiles = new List<string> { file1 };

            // Act & Assert
            Assert.Throws<ArgumentException>(() => DocumentMerger.MergeDocuments(inputFiles.ToArray(), ""));
        }

        [Fact]
        public void MergeDocuments_WithNonExistentFile_ShouldThrowFileNotFoundException()
        {
            // Arrange
            var existingFile = CreateTestFile("existing.docx", "Existing", "Content");
            var nonExistentFile = Path.Combine(_testDirectory, "nonexistent.docx");
            var outputFile = Path.Combine(_testDirectory, "output.docx");

            var inputFiles = new List<string> { existingFile, nonExistentFile };

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => DocumentMerger.MergeDocuments(inputFiles.ToArray(), outputFile));
        }

        [Fact]
        public void MergeDocuments_WithInvalidDocxFile_ShouldThrowException()
        {
            // Arrange
            var invalidFile = Path.Combine(_testDirectory, "invalid.docx");
            File.WriteAllText(invalidFile, "This is not a valid DOCX file");
            _testFiles.Add(invalidFile);

            var inputFiles = new List<string> { invalidFile };
            var outputFile = Path.Combine(_testDirectory, "output.docx");

            // Act & Assert
            Assert.ThrowsAny<Exception>(() => DocumentMerger.MergeDocuments(inputFiles.ToArray(), outputFile));
        }

        [Fact]
        public void MergeDocuments_WithReadOnlyOutputDirectory_ShouldThrowException()
        {
            // Arrange
            var file1 = CreateTestFile("test.docx", "Test", "Content");
            var inputFiles = new List<string> { file1 };

            // Try to write to a non-existent directory
            var invalidOutputFile = Path.Combine("/nonexistent/directory", "output.docx");

            // Act & Assert
            Assert.ThrowsAny<Exception>(() => DocumentMerger.MergeDocuments(inputFiles.ToArray(), invalidOutputFile));
        }

        [Fact]
        public void MergeDocuments_OverwriteExistingOutput_ShouldSucceed()
        {
            // Arrange
            var file1 = CreateTestFile("test1.docx", "Test 1", "Content 1");
            var file2 = CreateTestFile("test2.docx", "Test 2", "Content 2");
            var outputFile = Path.Combine(_testDirectory, "overwrite.docx");
            _testFiles.Add(outputFile);

            // Create an existing output file
            File.WriteAllText(outputFile, "Existing content");

            var inputFiles = new List<string> { file1, file2 };

            // Act
            DocumentMerger.MergeDocuments(inputFiles.ToArray(), outputFile);

            // Assert
            Assert.True(File.Exists(outputFile));

            using var doc = WordprocessingDocument.Open(outputFile, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;
            Assert.Contains("Test 1", text);
            Assert.Contains("Test 2", text);
            Assert.DoesNotContain("Existing content", text);
        }
    }
}
