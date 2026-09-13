using System;
using System.Globalization;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length != 4) return 2;
        try
        {
            double width = double.Parse(args[2], CultureInfo.InvariantCulture), height = double.Parse(args[3], CultureInfo.InvariantCulture);
            if (!(width > 0 && width <= 4032 && height > 0 && height <= 4032)) throw new ArgumentException("Invalid page dimensions.");
            using (PdfDocument pdf = PdfReader.Open(args[0], PdfDocumentOpenMode.Modify))
            {
                if (pdf.PageCount != 1) throw new InvalidDataException("Expected one PDF page.");
                PdfPage page = pdf.Pages[0];
                double pageHeight = page.MediaBox.Y2;
                var box = new PdfRectangle(new XPoint(0, pageHeight - height), new XPoint(width, pageHeight));
                page.MediaBox = box; page.CropBox = box; page.TrimBox = box; page.BleedBox = box; page.ArtBox = box;
                pdf.Info.Title = "obj2pdf";
                pdf.Save(args[1]);
            }
            return 0;
        }
        catch (Exception ex) { try { File.WriteAllText(args[1] + ".error", ex.Message); } catch { } return 1; }
    }
}
