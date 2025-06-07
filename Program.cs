using System;
using System.IO;
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

            // Check for special command to test compatibility mode
            if (args.Length == 1 && args[0] == "--test-compat")
            {
                TestCompatibilityMode();
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
            Console.WriteLine("       DocxMerger --test-compat    (tests compatibility mode)");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  DocxMerger document1.docx document2.docx merged.docx");
            Console.WriteLine("  DocxMerger doc1.docx doc2.docx doc3.docx final.docx");
            Console.WriteLine("  DocxMerger --create-test");
            Console.WriteLine("  DocxMerger --test-compat");
            Console.WriteLine();
            Console.WriteLine("Arguments:");
            Console.WriteLine("  input*.docx  - Input Word documents to merge (minimum 2 required)");
            Console.WriteLine("  output.docx  - Output merged document file");
            Console.WriteLine("  --create-test - Create sample test documents (test1.docx, test2.docx)");
            Console.WriteLine("  --test-compat - Test compatibility mode for merged documents");
            Console.WriteLine();
            Console.WriteLine("Notes:");
            Console.WriteLine("  - All input files must exist and be valid .docx files");
            Console.WriteLine("  - Documents are merged in the order specified");
            Console.WriteLine("  - The first document's styles and formatting are preserved");
        }

        /// <summary>
        /// Tests compatibility mode handling by creating a test document and processing it
        /// </summary>
        static void TestCompatibilityMode()
        {
            Console.WriteLine("Testing Compatibility Mode Handling");
            Console.WriteLine("===================================");
            
            try
            {
                // Create test files
                string compatFile = "test_compat_mode.docx";
                string normalFile = "test_normal.docx";
                string outputFile = "test_compat_merged.docx";
                
                Console.WriteLine("1. Creating compatibility mode test document...");
                CompatibilityModeTestHelper.CreateCompatibilityModeDocument(compatFile);
                
                Console.WriteLine("2. Creating normal test document...");
                DocxMerger.Tests.TestDocumentCreator.CreateTestDocument(normalFile, "Normal Document", "This is a normal document without compatibility mode.");
                
                Console.WriteLine("3. Merging documents (compatibility mode processing will be applied)...");
                DocumentMerger.MergeDocuments(new[] { compatFile, normalFile }, outputFile);
                
                Console.WriteLine();
                Console.WriteLine("✓ Compatibility mode test completed successfully!");
                Console.WriteLine($"✓ Output saved to: {outputFile}");
                Console.WriteLine();
                Console.WriteLine("The compatibility mode document was automatically upgraded to modern format during the merge process.");
                
                // Clean up test files
                Console.WriteLine();
                Console.WriteLine("Cleaning up test files...");
                if (File.Exists(compatFile)) File.Delete(compatFile);
                if (File.Exists(normalFile)) File.Delete(normalFile);
                Console.WriteLine("✓ Test files cleaned up.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Error during compatibility mode test: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
