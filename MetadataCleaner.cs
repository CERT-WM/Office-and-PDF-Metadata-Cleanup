using System;
using System.IO;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace MetadataCleanerApp
{
    public static class MetadataCleaner
    {
        public static void CleanMetadata(string inputFile, string outputFolder)
        {
            string ext = Path.GetExtension(inputFile).ToLower();
            string fileName = Path.GetFileNameWithoutExtension(inputFile);
            string outputFile = Path.Combine(outputFolder, $"{fileName}_meta_clean{ext}");

            if (ext == ".docx")
            {
                CleanWordMetadata(inputFile, outputFile);
            }
            else if (ext == ".xlsx")
            {
                CleanExcelMetadata(inputFile, outputFile);
            }
            else if (ext == ".pptx")
            {
                CleanPowerPointMetadata(inputFile, outputFile);
            }
            else if (ext == ".pdf")
            {
                CleanPdfMetadata(inputFile, outputFile);
            }
            else
            {
                throw new NotSupportedException($"File extension {ext} is not supported.");
            }
        }

        private static void CleanWordMetadata(string inputFile, string outputFile)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(inputFile, false))
            {
                using (WordprocessingDocument newDoc = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document))
                {
                    foreach (var part in doc.Parts)
                    {
                        newDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                    }

                    if (newDoc.CoreFilePropertiesPart != null)
                    {
                        newDoc.DeletePart(newDoc.CoreFilePropertiesPart);
                    }

                    if (newDoc.ExtendedFilePropertiesPart != null)
                    {
                        newDoc.DeletePart(newDoc.ExtendedFilePropertiesPart);
                    }
                }
            }
        }

        private static void CleanExcelMetadata(string inputFile, string outputFile)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(inputFile, false))
            {
                using (SpreadsheetDocument newDoc = SpreadsheetDocument.Create(outputFile, SpreadsheetDocumentType.Workbook))
                {
                    foreach (var part in doc.Parts)
                    {
                        newDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                    }

                    if (newDoc.CoreFilePropertiesPart != null)
                    {
                        newDoc.DeletePart(newDoc.CoreFilePropertiesPart);
                    }

                    if (newDoc.ExtendedFilePropertiesPart != null)
                    {
                        newDoc.DeletePart(newDoc.ExtendedFilePropertiesPart);
                    }
                }
            }
        }

        private static void CleanPowerPointMetadata(string inputFile, string outputFile)
        {
            using (PresentationDocument doc = PresentationDocument.Open(inputFile, false))
            {
                using (PresentationDocument newDoc = PresentationDocument.Create(outputFile, PresentationDocumentType.Presentation))
                {
                    foreach (var part in doc.Parts)
                    {
                        newDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                    }

                    if (newDoc.CoreFilePropertiesPart != null)
                    {
                        newDoc.DeletePart(newDoc.CoreFilePropertiesPart);
                    }

                    if (newDoc.ExtendedFilePropertiesPart != null)
                    {
                        newDoc.DeletePart(newDoc.ExtendedFilePropertiesPart);
                    }
                }
            }
        }

        private static void CleanPdfMetadata(string inputFile, string outputFile)
        {
            const int maxRetries = 3;
            const int retryDelayMs = 500;

            try
            {
                // Validate input file
                if (!File.Exists(inputFile))
                {
                    throw new FileNotFoundException($"Input PDF file does not exist: {inputFile}");
                }

                // Check input file accessibility
                try
                {
                    using (var testStream = File.Open(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        // File is accessible
                    }
                }
                catch (IOException ex)
                {
                    throw new IOException($"Cannot access input PDF file {inputFile}: {ex.Message}", ex);
                }

                // Check if output path is writable
                string outputDir = Path.GetDirectoryName(outputFile);
                if (outputDir == null || string.IsNullOrEmpty(outputDir))
                {
                    throw new ArgumentException($"Output directory path is invalid: {outputFile}");
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Use a temporary file to avoid output file conflicts
                string tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}_temp.pdf");

                try
                {
                    // Delete existing output file if it exists
                    if (File.Exists(outputFile))
                    {
                        try
                        {
                            File.Delete(outputFile);
                        }
                        catch (IOException ex)
                        {
                            throw new IOException($"Cannot delete existing output file {outputFile}: {ex.Message}", ex);
                        }
                    }

                    // Process PDF with retries
                    bool success = false;
                    for (int attempt = 1; attempt <= maxRetries; attempt++)
                    {
                        try
                        {
                            using (PdfDocument document = PdfReader.Open(inputFile, PdfDocumentOpenMode.Modify))
                            {
                                // Clear metadata
                                document.Info.Author = "";
                                document.Info.Creator = "";
                                document.Info.Keywords = "";
                                document.Info.Subject = "";
                                document.Info.Title = "";

                                // Save to temporary file
                                document.Save(tempFile);
                                document.Close();
                            }

                            // Move temp file to final output
                            File.Move(tempFile, outputFile);
                            success = true;
                            break;
                        }
                        catch (PdfSharp.Pdf.IO.PdfReaderException ex)
                        {
                            throw new Exception($"PDF processing error for {inputFile}: {ex.Message} (Type: {ex.GetType().FullName}). Possible causes: PDF is encrypted, password-protected, corrupted, or has an unsupported format.", ex);
                        }
                        catch (IOException ex)
                        {
                            if (attempt == maxRetries)
                            {
                                throw new IOException($"Failed to process PDF {inputFile} to {outputFile} after {maxRetries} attempts: {ex.Message}. Possible causes: Output file is locked by another process (e.g., PDF viewer, antivirus). Ensure {outputFile} is not open and try a different output folder.", ex);
                            }
                            Thread.Sleep(retryDelayMs);
                        }
                    }

                    if (!success)
                    {
                        throw new Exception($"PDF processing failed after {maxRetries} attempts for {inputFile}");
                    }
                }
                finally
                {
                    // Clean up temp file if it exists
                    if (File.Exists(tempFile))
                    {
                        try
                        {
                            File.Delete(tempFile);
                        }
                        catch
                        {
                            // Ignore cleanup errors
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                throw new Exception($"File access error for {inputFile} or {outputFile}: {ex.Message} (Type: {ex.GetType().FullName})", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Unexpected error processing PDF {inputFile}: {ex.Message} (Type: {ex.GetType().FullName})", ex);
            }
        }
    }
}