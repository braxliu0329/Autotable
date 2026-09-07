using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string licensePath = Path.Combine(AppContext.BaseDirectory, "Aspose.Total.NET.lic");
        
        if (!File.Exists(licensePath))
        {
             string projectRootLic = Path.Combine(AppContext.BaseDirectory, "../../../Aspose.Total.NET.lic");
             if (File.Exists(projectRootLic))
             {
                 licensePath = projectRootLic;
             }
        }

        Console.WriteLine($"Testing License from: {licensePath}");
        if (!File.Exists(licensePath))
        {
            Console.WriteLine("ERROR: License file not found!");
            return;
        }

        // Test Aspose.Words
        Console.WriteLine("\n--------------------------------------------------");
        Console.WriteLine("Testing Aspose.Words License...");
        try
        {
            var wordsLicense = new Aspose.Words.License();
            wordsLicense.SetLicense(licensePath);
            Console.WriteLine("Aspose.Words: License set successfully (No exception thrown).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Aspose.Words: Failed to set license.");
            Console.WriteLine("Error: " + ex.Message);
        }

        // Test Aspose.Slides
        Console.WriteLine("\n--------------------------------------------------");
        Console.WriteLine("Testing Aspose.Slides License...");
        try
        {
            var slidesLicense = new Aspose.Slides.License();
            slidesLicense.SetLicense(licensePath);
            Console.WriteLine("Aspose.Slides: License set successfully (No exception thrown).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Aspose.Slides: Failed to set license.");
            Console.WriteLine("Error: " + ex.Message);
        }

        // Test Aspose.PDF
        Console.WriteLine("\n--------------------------------------------------");
        Console.WriteLine("Testing Aspose.PDF License...");
        try
        {
            var pdfLicense = new Aspose.Pdf.License();
            pdfLicense.SetLicense(licensePath);
            Console.WriteLine("Aspose.PDF: License set successfully (No exception thrown).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Aspose.PDF: Failed to set license.");
            Console.WriteLine("Error: " + ex.Message);
        }
        Console.WriteLine("--------------------------------------------------\n");
    }
}
