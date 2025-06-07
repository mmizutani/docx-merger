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
    /// <summary>
    /// Result of compatibility mode processing operation
    /// </summary>
    public class CompatibilityProcessingResult
    {
        public string ProcessedFilePath { get; set; }
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public bool WasProcessed { get; set; } // Whether any processing was actually performed

        public CompatibilityProcessingResult(string processedFilePath, bool success, string? errorMessage = null, bool wasProcessed = false)
        {
            ProcessedFilePath = processedFilePath;
            Success = success;
            ErrorMessage = errorMessage;
            WasProcessed = wasProcessed;
        }
    }

    public static class DocumentMerger
    {
        public static void MergeDocuments(string[] fileNames, string outputFilePath, bool failOnCompatibilityProcessingError = false)
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
            var tempFilesToDelete = new List<string>();

            try
            {
                foreach (string fileName in fileNames)
                {
                    try
                    {
                        // Process the document to remove compatibility mode if present
                        var processingResult = ProcessCompatibilityMode(fileName);

                        // Handle processing failures based on configuration
                        if (!processingResult.Success && failOnCompatibilityProcessingError)
                        {
                            throw new InvalidOperationException(
                                $"Failed to process compatibility mode for file '{fileName}': {processingResult.ErrorMessage}. " +
                                "This may cause issues during document merging. Use failOnCompatibilityProcessingError=false to proceed with unprocessed files.");
                        }

                        // Track temporary files for cleanup (only if processing created a temp file)
                        if (processingResult.ProcessedFilePath != fileName)
                        {
                            tempFilesToDelete.Add(processingResult.ProcessedFilePath);
                        }

                        // Create a Source from each document file
                        var source = new Source(new WmlDocument(processingResult.ProcessedFilePath), true);
                        sources.Add(source);
                    }
                    catch (Exception ex) when (!(ex is InvalidOperationException && ex.Message.Contains("Failed to process compatibility mode")))
                    {
                        // Handle other file-related errors (e.g., corrupted files, invalid formats)
                        if (failOnCompatibilityProcessingError)
                        {
                            throw new InvalidOperationException(
                                $"Failed to process file '{fileName}': {ex.Message}. " +
                                "The file may be corrupted or in an unsupported format. " +
                                "Use failOnCompatibilityProcessingError=false to skip problematic files.", ex);
                        }
                        else
                        {
                            Console.WriteLine($"Warning: Skipping file '{fileName}' due to error: {ex.Message}");
                            continue;
                        }
                    }
                }

                // Check if we have any valid sources to merge
                if (sources.Count == 0)
                {
                    throw new InvalidOperationException("No valid documents were found to merge. All input files were either corrupted, invalid, or could not be processed.");
                }

                // Use DocumentBuilder to merge all documents
                DocumentBuilder.BuildDocument(sources, outputFilePath);
            }
            finally
            {
                // Clean up temporary files
                foreach (var tempFile in tempFilesToDelete)
                {
                    try
                    {
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                            Console.WriteLine($"✓ Temporary file deleted: {Path.GetFileName(tempFile)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not delete temporary file {Path.GetFileName(tempFile)}: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Processes a DOCX file to remove compatibility mode settings if present.
        /// Returns detailed information about the processing result.
        /// </summary>
        /// <param name="filePath">Path to the DOCX file to process</param>
        /// <returns>Processing result with file path, success status, and error details</returns>
        private static CompatibilityProcessingResult ProcessCompatibilityMode(string filePath)
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
                    return new CompatibilityProcessingResult(filePath, true);
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

                return new CompatibilityProcessingResult(tempDocxPath, true, null, true);
            }
            catch (Exception ex)
            {
                // If processing fails, log the error and return detailed failure information
                string errorMessage = $"Could not process compatibility mode for {filePath}: {ex.Message}";
                Console.WriteLine($"Warning: {errorMessage}");

                return new CompatibilityProcessingResult(filePath, false, errorMessage);
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
