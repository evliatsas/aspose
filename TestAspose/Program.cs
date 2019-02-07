using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Aspose.Pdf.Forms;
using System.Diagnostics;

namespace TestAspose
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            setLicense();

            Stopwatch sw = new Stopwatch();
            sw.Start();

            //GenerateRegistryReceipt();
            //ExportPdfFile("signed_document.pdf");
            //AddStampToPDF("signed_document.pdf", "Αρ. Πρωτ. : 4325/42");
            SignPdfDocument("test_pdf.pdf", "TestCA.pfx", "nigrita");
            
            //ExportWordFile("original\\_Παροχή οικονομικών στοιχείων (ΑΑ).doc");
            //ExportExcelFile("original\\_Στοιχεία Δήμων 2010-2012.xlsx");
            //ExportExcelFile("original\\_Έσοδα δήμων 2013-2017.xlsx");
            //ExportExcelFile("original\\Έξοδα δήμων 2013-2017.xlsx");

            //var originalFile = "test.doc";
            //var revisedFile = "test(1).docx";
            //CompareWordFiles(originalFile, revisedFile);

            sw.Stop();
            Console.WriteLine(string.Format("Elapsed: {0} ms", sw.Elapsed.TotalMilliseconds));
            Console.ReadLine();
        }

        private static void setLicense()
        {
            { //Enable Word support
                Aspose.Words.License license = new Aspose.Words.License();
                license.SetLicense("Aspose.Total.lic");
            }
            { //Enable Excel support
                Aspose.Cells.License license = new Aspose.Cells.License();
                license.SetLicense("Aspose.Total.lic");
            }
            { //Enable PDF support
                Aspose.Pdf.License license = new Aspose.Pdf.License();
                license.SetLicense("Aspose.Total.lic");
            }
        }

        internal static void ExportWordFile(string filename)
        {
            try
            {
                var path = $"original\\{filename}";
                Document doc = new Document(path);

                if (doc.ProtectionType != ProtectionType.NoProtection)
                    doc.Unprotect();

                if (doc.HasRevisions)
                {
                    doc.AcceptAllRevisions();
                    doc.TrackRevisions = false;
                }
                doc.HyphenationOptions.AutoHyphenation = false;
                doc.ViewOptions.ViewType = Aspose.Words.Settings.ViewType.PageLayout;

                var text = String.Format("Διανομή μέσω 'ΙΡΙΔΑ' από {0} την {1}(LT) με UID: {2}", "Aspose Test App", DateTime.Now.ToString("dd/MM/yy HH:mm"), "0042");
                DocumentBuilder builder = new DocumentBuilder(doc);
                // Create the footer.
                builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                builder.Writeln(text);

                var properties = doc.BuiltInDocumentProperties;
                properties["Title"].Value = "Δοκιμή Εγγράφου";
                properties["Author"].Value = "Aspose Test Word";

                var extIndex = filename.LastIndexOf(".");
                filename = filename.Remove(extIndex);
                filename = filename.Insert(extIndex, ".pdf");

                var outputPath = $"revised\\{filename}";
                doc.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine("Converted word document");
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }
        }

        internal static void ExportExcelFile(string filename)
        {
            try
            {
                var path = $"original\\{filename}";
                Document doc = new Document(path);

                // Check for password protection
                var info = Aspose.Cells.FileFormatUtil.DetectFileFormat(path);
                if (info.IsEncrypted)
                {
                    throw new Exception("The excel file is password protected...");
                }

                Aspose.Cells.Workbook book = new Aspose.Cells.Workbook(path);

                if (book.HasRevisions)
                {
                    book.AcceptAllRevisions();
                }

                // add footer to each sheet
                var footerText = string.Empty;
                footerText = String.Format("Διανομή μέσω 'ΙΡΙΔΑ' με UID: {0} στις {1}", filename, DateTime.Now.ToString("dd/MM/yy HH:mm"));
                var sheets = book.Worksheets;
                foreach (var sheet in sheets)
                {
                    var pSetup = sheet.PageSetup;
                    pSetup.SetFooter(1, footerText);
                }

                //set file metadata
                var properties = book.BuiltInDocumentProperties;
                properties["Title"].Value = "Δοκιμή Εγγράφου";
                properties["Author"].Value = "Aspose Test Excel";

                var extIndex = filename.LastIndexOf(".");
                filename = filename.Remove(extIndex);
                filename = filename.Insert(extIndex, ".pdf");

                var outputPath = $"revised\\{filename}";
                book.Save(outputPath, Aspose.Cells.SaveFormat.Pdf);

                Console.WriteLine("Converted excel document");
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }
        }

        internal static void GenerateRegistryReceipt()
        {
            try
            {
                Document doc = new Document("templates\\registry_receipt.docx");

                // This is the data for mail merge.
                String[] fieldNames = new String[] {
                        "inboundNo",
                        "entryDate",
                        "subject",
                        "recipients",
                        "sender",
                        "regNo",
                        "publicationDate",
                        "attachments"
                    };

                Object[] fieldValues = new Object[] {
                        "12345678",
                        DateTime.Now.Date.ToLocalTime().ToString("dd/MM/yyyy HH:mm"),
                        "my subject",
                        "ΥΠΕΣ/Δ. Ηλεκτρονικής",
                        "ΓΓΠ",
                        "2345/Σ.43/σαδφασδφ",
                        DateTime.Now.AddHours(-5).Date.ToLocalTime().ToString("dd/MM/yyyy HH:mm"),
                        "5 CD"
                    };

                // Execute the mail merge.
                doc.MailMerge.Execute(fieldNames, fieldValues);

                var properties = doc.BuiltInDocumentProperties;
                properties["Title"].Value = "Βεβαίωση";
                properties["Author"].Value = "Ίριδα";

                using (var final = new MemoryStream())
                {
                    doc.Save("revised\\receipt.pdf", SaveFormat.Pdf);
                }
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }
        }

        internal static void AddStampToPDF(string filename, string text)
        {
            var path = $"original\\{filename}";
            var formattedText = new Aspose.Pdf.Facades.FormattedText(
                text, 
                System.Drawing.Color.Black, 
                Aspose.Pdf.Facades.FontStyle.Helvetica, 
                Aspose.Pdf.Facades.EncodingType.Winansi, 
                true, 
                14);
            using (Aspose.Pdf.Document doc = new Aspose.Pdf.Document(path))
            {
                var fileStamp = new Aspose.Pdf.Facades.PdfFileStamp(doc);
                // Create stamp
                Aspose.Pdf.Facades.Stamp stamp = new Aspose.Pdf.Facades.Stamp();
                stamp.BindLogo(formattedText);

                stamp.Pages = new int[] { 1 };
                fileStamp.AddStamp(stamp);
                fileStamp.Save($"revised\\stamped_{filename}");
                fileStamp.Close();
            }

            Console.WriteLine("stamped....");
        }

        internal static void SignPdfDocument(string filename, string certificate, string password)
        {
            var path = $"original\\{filename}";
            var certPath = $"certificates\\{certificate}";

            var authority = "Ίριδα";
            var contactInfo = "Ευάγγελος Λιάτσας";
            var location = "ΥΠΕΘΑ";
            var reason = "Αρ.Π.:4322/12";

            using (Aspose.Pdf.Document doc = new Aspose.Pdf.Document(path))
            {                
                var signature = new Aspose.Pdf.Facades.PdfFileSignature(doc);
                // Create digital signature
                PKCS7 sig = new PKCS7(certPath, password); // Use PKCS7/PKCS7Detached objects
                sig.Authority = authority;
                sig.ContactInfo = contactInfo;
                sig.Location = location;
                sig.Reason = reason;
                sig.ShowProperties = false;
                // Set signature background image
                var lines = new List<string>() { authority, contactInfo, location, reason };
                signature.SignatureAppearanceStream = createSigningImage("emblem.png", lines);                
                // Set signature position
                var height = (int)(doc.PageInfo.Height - doc.PageInfo.Margin.Top);
                var width = (int)(doc.PageInfo.Width - doc.PageInfo.Margin.Right);
                var size = 50;
                var rect = new Aspose.Pdf.Rectangle(width - size, height - size, width, height);
                // Sign the document
                signature.Sign(1, true, rect.ToRect(), sig);
                // Save output PDF file
                var outputPath = $"revised\\signed_{filename}";
                signature.Save(outputPath);
            }

            Console.WriteLine("signed....");
        }

        internal static void CompareWordFiles(string originalFile, string revisedFile)
        {
            try
            {
                Document original = new Document($"original\\{originalFile}");
                Document revised = new Document($"original\\{revisedFile}");

                if (original.ProtectionType != ProtectionType.NoProtection)
                    original.Unprotect();

                if (revised.ProtectionType != ProtectionType.NoProtection)
                    revised.Unprotect();

                if (original.HasRevisions)
                    original.AcceptAllRevisions();

                if (revised.HasRevisions)
                    revised.AcceptAllRevisions();

                CompareOptions options = new CompareOptions();
                options.IgnoreFormatting = true;
                options.IgnoreHeadersAndFooters = true;
                original.TrackRevisions = true;
                original.HyphenationOptions.AutoHyphenation = false;
                // original now contains changes as revisions. 
                var author = "aspose tester";
                original.Compare(revised, author, DateTime.Now, options);

                var final = new MemoryStream();
                if (revisedFile.EndsWith(".docx"))
                {
                    original.Save(final, SaveFormat.Docx);
                }
                else if (revisedFile.EndsWith(".doc"))
                {
                    original.Save(final, SaveFormat.Doc);
                }
                else if (revisedFile.EndsWith(".docm"))
                {
                    original.Save(final, SaveFormat.Docm);
                }
                else if (revisedFile.EndsWith(".odt"))
                {
                    original.Save(final, SaveFormat.Odt); ;
                }

                using (FileStream file = new FileStream($"revised\\new_{revisedFile}", FileMode.Create, System.IO.FileAccess.Write))
                    final.WriteTo(file);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc);
            }
        }

        internal static void ExportPdfFile(string filename)
        {
            try
            {
                var path = $"original\\{filename}";

                var author = "Ευάγγελος Λιάτσας";
                var docid = Guid.NewGuid();
                var title = filename + " - Δοκιμή PDF διαχείριση";

                var document = new Aspose.Pdf.Document(path);

                // Check for password protection
                var info = new Aspose.Pdf.Facades.PdfFileInfo(document);
                // Determine if the source PDF is encrypted
                if (info.IsEncrypted)
                {
                    throw new Exception("The PDF file is password protected...");
                }
                var pdfSign = new Aspose.Pdf.Facades.PdfFileSignature(document);
                if (pdfSign.ContainsSignature())
                {
                    var signees = pdfSign.GetSignNames();
                    foreach (var signee in signees)
                    {
                        Console.WriteLine($"Is digitally signed by {signee} ...");
                    }
                }
                else
                {

                    // add footer to each sheet
                    var footerText = string.Empty;
                    if (author != null)
                        footerText = String.Format("Διανομή μέσω 'ΙΡΙΔΑ' με UID: {0} στις {1}", docid, DateTime.Now.ToString("dd/MM/yy HH:mm"));
                    // Create footer
                    Aspose.Pdf.TextStamp textStamp = new Aspose.Pdf.TextStamp(footerText);
                    // Set properties of the stamp
                    textStamp.BottomMargin = 10;
                    textStamp.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                    textStamp.VerticalAlignment = Aspose.Pdf.VerticalAlignment.Bottom;
                    // Add footer on all pages
                    foreach (Aspose.Pdf.Page page in document.Pages)
                    {
                        page.AddStamp(textStamp);
                    }

                    Console.WriteLine("added footer ...");
                }

                //set file metadata
                document.Info.Title = title;
                if (author != null)
                    document.Info.Author = author;

                var final = new MemoryStream();
                document.Save(final);
                using (FileStream file = new FileStream($"revised\\export_{filename}", FileMode.Create, System.IO.FileAccess.Write))
                    final.WriteTo(file);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc);
            }
        }

        private static Stream createSigningImage(string imageName, IEnumerable<string> lines)
        {
            var path = $"images\\{imageName}";
            var result = new MemoryStream();
            //Creates an instance of Image
            using (var file = new FileStream(path, FileMode.Open))
            {
                using (var image = new System.Drawing.Bitmap(file))
                {
                    //Creates and initialize an instance of Graphics class
                    using (var graphics = System.Drawing.Graphics.FromImage(image))
                    {
                        var font = new System.Drawing.Font("Helvetica", 18);
                        var brush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);

                        float x = 10;
                        float y = 50;
                        foreach (var line in lines)
                        {
                            //Draw a String
                            graphics.DrawString(line, font, brush, new System.Drawing.PointF(x, y));
                            y += 35;
                        }
                    }

                    // save all changes                
                    image.Save(result, System.Drawing.Imaging.ImageFormat.Png);
                }
            }

            return result;
        }
    }
}
