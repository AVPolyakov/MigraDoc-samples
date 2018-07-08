using System;
using System.Diagnostics;
using PdfSharp.Pdf;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;

namespace HelloWorld
{
    /// <summary>
    /// This sample is the obligatory Hello World program for MigraDoc documents.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            M1(100);
            for (int i = 1; i <= 100; i++)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();

                var count = 1000*i;
                M1(count);

                stopwatch.Stop();

                Console.WriteLine($"{count} {stopwatch.ElapsedMilliseconds}");
            }
        }

        private static string M1(int count)
        {
// Create a MigraDoc document.
            var document = CreateDocument(count);

            // ----- Unicode encoding and font program embedding in MigraDoc is demonstrated here. -----

            // A flag indicating whether to create a Unicode PDF or a WinAnsi PDF file.
            // This setting applies to all fonts used in the PDF document.
            // This setting has no effect on the RTF renderer.
            const bool unicode = true;

            // Create a renderer for the MigraDoc document.
            var pdfRenderer = new PdfDocumentRenderer(unicode);

            // Associate the MigraDoc document with a renderer.
            pdfRenderer.Document = document;

            // Layout and render document to PDF.
            pdfRenderer.RenderDocument();

            // Save the document...
            var filename = $"HelloWorld{Guid.NewGuid()}.pdf";
            pdfRenderer.PdfDocument.Save(filename);
            return filename;
        }

        static readonly string text2 = "QWEQWE QWETQWTQ WRTWERT QWETQWET QWEQWR ASDASF".ToUpper();
        public const double BorderWidth = 0.5D;

        /// <summary>
        /// Creates an absolutely minimalistic document.
        /// </summary>
        /// <param name="count"></param>
        static Document CreateDocument(int count)
        {
            // Create a new MigraDoc document.
            var document = new Document();


            // Add a section to the document.
            var section = document.AddSection();

            section.PageSetup.TopMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(1);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(2);
            section.PageSetup.RightMargin = Unit.FromCentimeter(1);
            section.PageSetup.Orientation = Orientation.Landscape;

            var table = new Table();
            table.Borders.Width = BorderWidth;
            table.LeftPadding = table.RightPadding = 0;
            table.Format.Font.Name = "Times New Roman";
            table.Format.Font.Size = 8;
            var columnCount = 14;
            for (var i = 0; i < columnCount; i++)
                table.AddColumn(54.0607424071991);

            for (var j = 0; j < count; j++)
            {
                var r = table.AddRow();
                for (var i = 0; i < columnCount; i++)
                {
                    var paragraph = new Paragraph();
                    paragraph.Format.LeftIndent = Unit.FromCentimeter(0.05);
                    paragraph.Format.RightIndent = Unit.FromCentimeter(0.05);
                    paragraph.AddText(text2);
                    r[i].Add(paragraph);
                }
            }

            document.LastSection.Add(table);

            return document;
        }
    }
}
