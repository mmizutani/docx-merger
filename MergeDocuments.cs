using Codeuctivity.OpenXmlPowerTools;
using Codeuctivity.OpenXmlPowerTools.DocumentBuilder;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;

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
                // Process the document to remove compatibility mode if present
                string processedFileName = ProcessCompatibilityMode(fileName);

                // Create a Source from each document file
                var source = new Source(new WmlDocument(processedFileName), true);
                sources.Add(source);
            }

            // Use DocumentBuilder to merge all documents
            DocumentBuilder.BuildDocument(sources, outputFilePath);
        }

        /// <summary>
        /// Processes a DOCX file to remove compatibility mode settings if present.
        /// Returns the path to the processed file (either the original if no changes needed,
        /// or a temporary file with compatibility mode removed).
        /// </summary>
        /// <param name="filePath">Path to the DOCX file to process</param>
        /// <returns>Path to the processed file</returns>
        private static string ProcessCompatibilityMode(string filePath)
        {
            try
            {
                bool hasCompatibilityMode = false;

                // Check if the document has compatibility mode settings
                using (var doc = WordprocessingDocument.Open(filePath, false))
                {
                    var settingsPart = doc.MainDocumentPart?.DocumentSettingsPart;
                    if (settingsPart?.Settings != null)
                    {
                        var compat = settingsPart.Settings.Elements<Compatibility>().FirstOrDefault();
                        hasCompatibilityMode = compat != null && compat.HasChildren;
                    }
                }

                // If no compatibility mode detected, return original file
                if (!hasCompatibilityMode)
                {
                    Console.WriteLine($"✓ No compatibility mode detected in: {Path.GetFileName(filePath)}");
                    return filePath;
                }

                // Create a temporary file to store the processed document
                string tempFilePath = Path.GetTempFileName();
                string tempDocxPath = Path.ChangeExtension(tempFilePath, ".docx");
                File.Delete(tempFilePath); // Remove the temp file created by GetTempFileName

                // Copy the original file to temp location
                File.Copy(filePath, tempDocxPath, true);

                // Remove compatibility mode from the temporary file
                RemoveCompatibilityMode(tempDocxPath);

                Console.WriteLine($"✓ Compatibility mode removed from: {Path.GetFileName(filePath)}");

                return tempDocxPath;
            }
            catch (Exception ex)
            {
                // If processing fails, log the error and return the original file
                Console.WriteLine($"Warning: Could not process compatibility mode for {filePath}: {ex.Message}");
                return filePath;
            }
        }

        /// <summary>
        /// Removes compatibility mode settings from a DOCX file.
        /// </summary>
        /// <param name="filePath">Path to the DOCX file to modify</param>
        private static void RemoveCompatibilityMode(string filePath)
        {
            using (var doc = WordprocessingDocument.Open(filePath, true))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart == null) return;

                // Ensure DocumentSettingsPart exists
                var settingsPart = mainPart.DocumentSettingsPart;
                if (settingsPart == null)
                {
                    settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                }

                var settings = settingsPart.Settings;
                if (settings == null)
                {
                    settings = new Settings();
                    settingsPart.Settings = settings;
                }

                // Remove compatibility settings
                var compatElements = settings.Elements<Compatibility>().ToList();
                foreach (var compat in compatElements)
                {
                    compat.Remove();
                }

                // Remove document protection if it exists (often related to compatibility mode)
                var docProtection = settings.Elements<DocumentProtection>().FirstOrDefault();
                if (docProtection != null)
                {
                    // Only remove if it's related to compatibility (forms protection)
                    if (docProtection.Edit?.Value == DocumentProtectionValues.Forms)
                    {
                        docProtection.Remove();
                    }
                }

                // Optionally add modern compatibility settings for the latest Word version
                var newCompat = new Compatibility();

                // Add compatibility settings for Word 2019/365 (version 16)
                newCompat.AppendChild(new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "16" // Word 2019/365 compatibility mode
                });

                settings.AppendChild(newCompat);

                // Save the changes
                settings.Save();
            }
        }
    }
}
