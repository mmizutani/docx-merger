using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxMerger.Tests
{
    public static class TestDocumentCreator
    {
        public static void CreateTestDocument(string filePath, string title, string content)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Create the main document part
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document tree
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Add title paragraph
                Paragraph titlePara = body.AppendChild(new Paragraph());
                Run titleRun = titlePara.AppendChild(new Run());
                RunProperties titleRunProps = titleRun.AppendChild(new RunProperties());
                titleRunProps.AppendChild(new Bold());
                titleRunProps.AppendChild(new FontSize() { Val = "28" });
                titleRun.AppendChild(new Text(title));

                // Add content paragraph
                Paragraph contentPara = body.AppendChild(new Paragraph());
                Run contentRun = contentPara.AppendChild(new Run());
                contentRun.AppendChild(new Text(content));

                // Save the document
                mainPart.Document.Save();
            }
        }

        public static void CreateTestDocuments()
        {
            CreateTestDocument("test1.docx", "Document 1", "This is the content of the first document. It contains some sample text for testing the merge functionality.");
            CreateTestDocument("test2.docx", "Document 2", "This is the content of the second document. It will be merged with the first document to create a combined output.");

            Console.WriteLine("Test documents created:");
            Console.WriteLine("- test1.docx");
            Console.WriteLine("- test2.docx");
        }
    }
}
