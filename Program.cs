using System;
using DocxMerger;

namespace DocxMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("DOCX Merger - OpenXML PowerTools");
            Console.WriteLine("================================");

            // Check for special command to create test documents
            if (args.Length == 1 && args[0] == "--create-test")
            {
                DocxMerger.Tests.TestDocumentCreator.CreateTestDocuments();
                return;
            }

            if (args.Length < 3)
            {
                ShowUsage();
                return;
            }

            // Parse command line arguments
            // Last argument is the output file
            string outputFile = args[args.Length - 1];

            // All other arguments are input files
            string[] inputFiles = new string[args.Length - 1];
            Array.Copy(args, inputFiles, args.Length - 1);

            try
            {
                Console.WriteLine($"Merging {inputFiles.Length} document(s):");
                foreach (var file in inputFiles)
                {
                    Console.WriteLine($"  - {file}");
                }
                Console.WriteLine($"Output: {outputFile}");
                Console.WriteLine();

                // Perform the merge
                DocumentMerger.MergeDocuments(inputFiles, outputFile);

                Console.WriteLine("✓ Documents merged successfully!");
                Console.WriteLine($"Output saved to: {outputFile}");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
                Environment.Exit(1);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
                Environment.Exit(1);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Unexpected error: {ex.Message}");
                Console.WriteLine($"Details: {ex}");
                Environment.Exit(1);
            }
        }

        private static void ShowUsage()
        {
            Console.WriteLine("Usage: DocxMerger <input1.docx> <input2.docx> [input3.docx ...] <output.docx>");
            Console.WriteLine("       DocxMerger --create-test    (creates test documents)");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  DocxMerger document1.docx document2.docx merged.docx");
            Console.WriteLine("  DocxMerger doc1.docx doc2.docx doc3.docx final.docx");
            Console.WriteLine("  DocxMerger --create-test");
            Console.WriteLine();
            Console.WriteLine("Arguments:");
            Console.WriteLine("  input*.docx  - Input Word documents to merge (minimum 2 required)");
            Console.WriteLine("  output.docx  - Output merged document file");
            Console.WriteLine("  --create-test - Create sample test documents (test1.docx, test2.docx)");
            Console.WriteLine();
            Console.WriteLine("Notes:");
            Console.WriteLine("  - All input files must exist and be valid .docx files");
            Console.WriteLine("  - Documents are merged in the order specified");
            Console.WriteLine("  - The first document's styles and formatting are preserved");
        }
    }
}
