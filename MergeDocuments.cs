using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System;

namespace DocxMerger
{
    public static class DocumentMerger
    {
        public static void MergeDocuments(string[] fileNames, string outputFilePath)
        {
            if (fileNames == null || fileNames.Length == 0)
                throw new ArgumentException("At least one input file must be specified.", nameof(fileNames));

            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path cannot be null or empty.", nameof(outputFilePath));

            // Verify all input files exist
            foreach (string fileName in fileNames)
            {
                if (!File.Exists(fileName))
                    throw new FileNotFoundException($"Input file not found: {fileName}");
            }

            var sources = new List<Source>();

            foreach (string fileName in fileNames)
            {
                // Create a Source from each document file
                var source = new Source(new WmlDocument(fileName), true);
                sources.Add(source);
            }

            // Use DocumentBuilder to merge all documents
            DocumentBuilder.BuildDocument(sources, outputFilePath);
        }
    }
}
