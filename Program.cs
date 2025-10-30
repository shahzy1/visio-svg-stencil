using System;
using System.IO;
using Visio = Microsoft.Office.Interop.Visio;

namespace AzureStencilBuilder
{
    class Program
    {
        // Desired master size (in inches)
        const double MASTER_WIDTH = 1.0;
        const double MASTER_HEIGHT = 1.0;

        static void Main(string[] args)
        {
            string baseFolder = @"C:\Dev\Temp\AzureSVGs";   // Root folder with category subfolders
            string outputFolder = @"C:\Dev\Temp\Stencils";  // Output folder for stencils

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var visioApp = new Visio.Application();
            visioApp.Visible = false;

            try
            {
                foreach (var categoryDir in Directory.GetDirectories(baseFolder))
                {
                    string categoryName = Path.GetFileName(categoryDir);
                    string stencilPath = Path.Combine(outputFolder, $"Azure-{categoryName}.vssx");

                    Console.WriteLine($"\n📘 Building stencil: {stencilPath}");

                    var stencilDoc = visioApp.Documents.Add("");
                    var tempDoc = visioApp.Documents.Add("");
                    var tempPage = tempDoc.Pages.Add();
                    tempPage.Name = "TempImport";

                    foreach (var svgFile in Directory.GetFiles(categoryDir, "*.svg"))
                    {
                        string baseName = Path.GetFileNameWithoutExtension(svgFile);
                        Console.WriteLine($"  → Importing: {baseName}");

                        try
                        {
                            // Import SVG
                            var shape = tempPage.Import(svgFile);

                            // Group if multiple shapes
                            if (tempPage.Shapes.Count > 1)
                            {
                                var shapes = tempPage.CreateSelection(Visio.VisSelectionTypes.visSelTypeAll);
                                shape = shapes.Group();
                            }

                            // Resize icon to fit master
                            double ratioW = MASTER_WIDTH / shape.CellsU["Width"].ResultIU;
                            double ratioH = MASTER_HEIGHT / shape.CellsU["Height"].ResultIU;
                            double scale = Math.Min(ratioW, ratioH);
                            shape.CellsU["Width"].ResultIU *= scale;
                            shape.CellsU["Height"].ResultIU *= scale;

                            // Center icon horizontally and move slightly up
                            shape.CellsU["PinX"].ResultIU = MASTER_WIDTH / 2;
                            shape.CellsU["PinY"].ResultIU = MASTER_HEIGHT * 0.6;

                            // Add a text label below the icon
                            var textShape = tempPage.DrawRectangle(0, 0, MASTER_WIDTH, 0.2); // small rectangle for text
                            textShape.Text = baseName;
                            textShape.CellsU["PinX"].ResultIU = MASTER_WIDTH / 2;
                            textShape.CellsU["PinY"].ResultIU = 0.1; // slightly below icon
                            textShape.CellsU["Char.Size"].ResultIU = 0.08; // font size
                            textShape.CellsU["Para.HorzAlign"].FormulaU = "0"; // center alignment

                            // Group icon + label
                            var grouped = tempPage.CreateSelection(Visio.VisSelectionTypes.visSelTypeAll).Group();

                            // Drop as master
                            var master = stencilDoc.Masters.Drop(grouped, 0, 0);
                            master.NameU = baseName;
                            master.Name = baseName;

                            // Cleanup
                            grouped.Delete();
                            tempPage.Delete((short)Visio.VisDeleteFlags.visDeleteNormal);
                            tempPage = tempDoc.Pages.Add();
                            tempPage.Name = "TempImport";
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Skipped {baseName}: {ex.Message}");
                        }
                    }

                    // Save stencil
                    stencilDoc.SaveAs(stencilPath);
                    stencilDoc.Close();
                    tempDoc.Close();

                    Console.WriteLine($"Completed stencil: {stencilPath}");
                }

                Console.WriteLine("\nAll stencils created successfully!");
            }
            finally
            {
                visioApp.Quit();
            }
        }
    }
}
