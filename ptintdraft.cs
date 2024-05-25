using System;
using System.IO;
using SolidEdgeFramework;
using SolidEdgeDraft;

namespace SolidEdgePrintDraftsApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the folder containing the draft files.
            string folderPath = @"C:\Users\admin\Desktop\Solid edge part";

            // Check if the specified folder exists.
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"The folder path {folderPath} does not exist.");
                return;
            }

            // Get all draft files in the specified folder.
            string[] draftFiles = Directory.GetFiles(folderPath, "*.dft");

            if (draftFiles.Length == 0)
            {
                Console.WriteLine($"No draft files found in the folder {folderPath}.");
                return;
            }

            // Initialize Solid Edge.
            SolidEdgeFramework.Application application = null;

            try
            {
                application = (SolidEdgeFramework.Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
                application.Visible = false;

                foreach (string draftFile in draftFiles)
                {
                    Console.WriteLine($"Processing file: {draftFile}");
                    try
                    {
                        // Open the draft document.
                        SolidEdgeDraft.DraftDocument draftDocument = (SolidEdgeDraft.DraftDocument)application.Documents.Open(draftFile);

                        // Print the draft document.
                        draftDocument.PrintOut();
                        Console.WriteLine($"Printed: {draftFile}");

                        // Close the draft document.
                        draftDocument.Close(false);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file {draftFile}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing Solid Edge: {ex.Message}");
            }
            finally
            {
                if (application != null)
                {
                    // Clean up Solid Edge application.
                    application.Quit();
                }
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
