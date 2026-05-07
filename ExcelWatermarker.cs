namespace ExcelWatermarker
{
    class ExcelWatermarker
    {
        // static async Task Handle(string[] args)
        // {
        //     string excelFilePath = @"C:\temp\your_input_file.xlsx";
        //     string outputFilePath = @"C:\temp\watermarked_output_file.xlsx";

        //     if (!File.Exists(excelFilePath))
        //     {
        //         Console.WriteLine($"Error: Input file not found at '{excelFilePath}'");
        //         return;
        //     }

        //     try
        //     {
        //         // Load the existing Excel file
        //         var excelImporter = new ExcelImporter();
        //         var excelFile = await excelImporter.Import(excelFilePath);

        //         // Add a watermark to each sheet in the workbook
        //         foreach (var sheet in excelFile.Sheets)
        //         {
        //             // Add a text watermark using the built-in feature
        //             // This creates a shape with text that acts as a watermark
        //             sheet.AddWatermark("CONFIDENTIAL");
        //         }

        //         // Save the watermarked Excel file
        //         var excelExporter = new ExcelExporter();
        //         await excelExporter.Export(outputFilePath, excelFile);

        //         Console.WriteLine($"✅ Watermark added successfully!");
        //         Console.WriteLine($"📁 Output file saved to: {outputFilePath}");
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.WriteLine($"❌ An error occurred: {ex.Message}");
        //     }
        // }
   
    }
}