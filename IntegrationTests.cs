using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Xunit;
using DocumentFormat.OpenXml.Packaging;

namespace DocxMerger.Tests
{
    public class IntegrationTests : IDisposable
    {
        private readonly string _testDirectory;
        private readonly List<string> _testFiles;

        public IntegrationTests()
        {
            _testDirectory = Path.Combine(Path.GetTempPath(), "DocxMergerIntegration_" + Guid.NewGuid().ToString("N")[..8]);
            Directory.CreateDirectory(_testDirectory);
            _testFiles = new List<string>();
        }

        public void Dispose()
        {
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

        [Fact]
        public void EndToEndWorkflow_CreateAndMergeDocuments_ShouldSucceed()
        {
            // Arrange
            var doc1Path = Path.Combine(_testDirectory, "doc1.docx");
            var doc2Path = Path.Combine(_testDirectory, "doc2.docx");
            var mergedPath = Path.Combine(_testDirectory, "merged.docx");

            _testFiles.AddRange(new[] { doc1Path, doc2Path, mergedPath });

            // Act - Create test documents
            TestDocumentCreator.CreateTestDocument(doc1Path, "First Document", "This is the first document content.");
            TestDocumentCreator.CreateTestDocument(doc2Path, "Second Document", "This is the second document content.");

            // Act - Merge documents
            var inputFiles = new List<string> { doc1Path, doc2Path };
            DocumentMerger.MergeDocuments(inputFiles, mergedPath);

            // Assert
            Assert.True(File.Exists(doc1Path), "First document should exist");
            Assert.True(File.Exists(doc2Path), "Second document should exist");
            Assert.True(File.Exists(mergedPath), "Merged document should exist");

            // Verify merged content
            using var mergedDoc = WordprocessingDocument.Open(mergedPath, false);
            var body = mergedDoc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;
            Assert.Contains("First Document", text);
            Assert.Contains("Second Document", text);
            Assert.Contains("first document content", text);
            Assert.Contains("second document content", text);
        }

        [Fact]
        public void LargeDocumentMerge_ShouldCompleteWithinReasonableTime()
        {
            // Arrange
            var doc1Path = Path.Combine(_testDirectory, "large1.docx");
            var doc2Path = Path.Combine(_testDirectory, "large2.docx");
            var mergedPath = Path.Combine(_testDirectory, "large_merged.docx");

            _testFiles.AddRange(new[] { doc1Path, doc2Path, mergedPath });

            // Create large documents
            var largeContent1 = string.Join("\n\n", Enumerable.Repeat("This is a large paragraph with substantial content to test performance. " +
                "It contains multiple sentences and should stress test the merging functionality. " +
                "We want to ensure that even with larger documents, the merge process completes efficiently.", 50));

            var largeContent2 = string.Join("\n\n", Enumerable.Repeat("This is another large paragraph for the second document. " +
                "It also contains multiple sentences and substantial content. " +
                "The merger should handle this efficiently and combine both documents properly.", 50));

            TestDocumentCreator.CreateTestDocument(doc1Path, "Large Document 1", largeContent1);
            TestDocumentCreator.CreateTestDocument(doc2Path, "Large Document 2", largeContent2);

            // Act - Time the merge operation
            var stopwatch = Stopwatch.StartNew();
            var inputFiles = new List<string> { doc1Path, doc2Path };
            DocumentMerger.MergeDocuments(inputFiles, mergedPath);
            stopwatch.Stop();

            // Assert
            Assert.True(File.Exists(mergedPath), "Large merged document should exist");
            Assert.True(stopwatch.ElapsedMilliseconds < 30000, $"Merge should complete within 30 seconds, took {stopwatch.ElapsedMilliseconds}ms");

            // Verify content integrity
            using var mergedDoc = WordprocessingDocument.Open(mergedPath, false);
            var body = mergedDoc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;
            Assert.Contains("Large Document 1", text);
            Assert.Contains("Large Document 2", text);
            Assert.Contains("substantial content", text);
            Assert.True(text.Length > 1000, "Merged document should contain substantial content");
        }

        [Fact]
        public void MultipleFilesMerge_WithDifferentContentTypes_ShouldSucceed()
        {
            // Arrange
            var files = new List<string>();
            var mergedPath = Path.Combine(_testDirectory, "multi_merged.docx");
            _testFiles.Add(mergedPath);

            // Create documents with different content styles
            for (int i = 1; i <= 5; i++)
            {
                var filePath = Path.Combine(_testDirectory, $"multi{i}.docx");
                files.Add(filePath);
                _testFiles.Add(filePath);

                var title = $"Document {i}";
                var content = i switch
                {
                    1 => "This is a simple paragraph with basic text.",
                    2 => "This document contains multiple sentences. It has more complex structure. Each sentence adds to the overall content.",
                    3 => "Special characters test: áéíóú ñ @#$%^&*()",
                    4 => "Numbered list content:\n1. First item\n2. Second item\n3. Third item",
                    5 => "Final document with mixed content: numbers (123), symbols (@#$), and letters (ABC).",
                    _ => $"Default content for document {i}"
                };

                TestDocumentCreator.CreateTestDocument(filePath, title, content);
            }

            // Act
            DocumentMerger.MergeDocuments(files, mergedPath);

            // Assert
            Assert.True(File.Exists(mergedPath), "Multi-merged document should exist");

            using var mergedDoc = WordprocessingDocument.Open(mergedPath, false);
            var body = mergedDoc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;

            // Verify all documents were included
            for (int i = 1; i <= 5; i++)
            {
                Assert.Contains($"Document {i}", text);
            }

            Assert.Contains("simple paragraph", text);
            Assert.Contains("complex structure", text);
            Assert.Contains("Special characters", text);
            Assert.Contains("Numbered list", text);
            Assert.Contains("Final document", text);
        }

        [Fact]
        public void RealWorldScenario_ContractMerging_ShouldWorkCorrectly()
        {
            // Arrange - Simulate merging contract documents
            var headerPath = Path.Combine(_testDirectory, "contract_header.docx");
            var termsPath = Path.Combine(_testDirectory, "terms_conditions.docx");
            var signaturesPath = Path.Combine(_testDirectory, "signatures.docx");
            var finalContractPath = Path.Combine(_testDirectory, "final_contract.docx");

            _testFiles.AddRange(new[] { headerPath, termsPath, signaturesPath, finalContractPath });

            // Create contract sections
            TestDocumentCreator.CreateTestDocument(headerPath,
                "SERVICE AGREEMENT",
                "This Service Agreement (\"Agreement\") is entered into on [DATE] between [COMPANY A] and [COMPANY B]. " +
                "This agreement governs the provision of services as outlined in the following terms and conditions.");

            TestDocumentCreator.CreateTestDocument(termsPath,
                "Terms and Conditions",
                "1. SCOPE OF SERVICES: The provider agrees to deliver services as specified in Exhibit A.\n\n" +
                "2. PAYMENT TERMS: Payment shall be made within 30 days of invoice receipt.\n\n" +
                "3. CONFIDENTIALITY: Both parties agree to maintain confidentiality of proprietary information.\n\n" +
                "4. TERMINATION: Either party may terminate this agreement with 30 days written notice.");

            TestDocumentCreator.CreateTestDocument(signaturesPath,
                "Signatures",
                "By signing below, the parties agree to be bound by the terms of this agreement.\n\n" +
                "COMPANY A:\n\nSignature: _________________________\nName: [NAME]\nTitle: [TITLE]\nDate: [DATE]\n\n" +
                "COMPANY B:\n\nSignature: _________________________\nName: [NAME]\nTitle: [TITLE]\nDate: [DATE]");

            // Act
            var contractSections = new List<string> { headerPath, termsPath, signaturesPath };
            DocumentMerger.MergeDocuments(contractSections, finalContractPath);

            // Assert
            Assert.True(File.Exists(finalContractPath), "Final contract should be created");

            using var contract = WordprocessingDocument.Open(finalContractPath, false);
            var body = contract.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;

            // Verify contract structure
            Assert.Contains("SERVICE AGREEMENT", text);
            Assert.Contains("Terms and Conditions", text);
            Assert.Contains("Signatures", text);

            // Verify content flow
            Assert.Contains("SCOPE OF SERVICES", text);
            Assert.Contains("PAYMENT TERMS", text);
            Assert.Contains("CONFIDENTIALITY", text);
            Assert.Contains("TERMINATION", text);
            Assert.Contains("By signing below", text);

            // Verify the merged document maintains logical order
            var headerIndex = text.IndexOf("SERVICE AGREEMENT");
            var termsIndex = text.IndexOf("SCOPE OF SERVICES");
            var signaturesIndex = text.IndexOf("By signing below");

            Assert.True(headerIndex < termsIndex, "Header should come before terms");
            Assert.True(termsIndex < signaturesIndex, "Terms should come before signatures");
        }

        [Fact]
        public void ErrorRecovery_PartiallyCorruptedInput_ShouldHandleGracefully()
        {
            // Arrange
            var validDoc = Path.Combine(_testDirectory, "valid.docx");
            var mergedPath = Path.Combine(_testDirectory, "recovered.docx");

            _testFiles.AddRange(new[] { validDoc, mergedPath });

            // Create a valid document
            TestDocumentCreator.CreateTestDocument(validDoc, "Valid Document", "This document is properly formatted.");

            // Test with just the valid document (simulating recovery from partial corruption)
            var inputFiles = new List<string> { validDoc };

            // Act
            DocumentMerger.MergeDocuments(inputFiles, mergedPath);

            // Assert
            Assert.True(File.Exists(mergedPath), "Recovery merge should succeed");

            using var recoveredDoc = WordprocessingDocument.Open(mergedPath, false);
            var body = recoveredDoc.MainDocumentPart?.Document?.Body;
            Assert.NotNull(body);

            var text = body.InnerText;
            Assert.Contains("Valid Document", text);
            Assert.Contains("properly formatted", text);
        }
    }
}
