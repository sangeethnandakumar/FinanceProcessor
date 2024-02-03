using FinanceProcessor.Application.Handlers;
using FinanceProcessor.Core.Exceptions;
using FinanceProcessor.Core.Stetement;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Razor.Templating.Core;
using Serilog;
using Serilog.Context;
using System.Data;
using System.Diagnostics;
using System.Text;
using XMLToPDF;

namespace FinanceProcessor
{

    public partial class Main : Form
    {
        private readonly IExcelHandler excelHandler;
        private readonly IPDFHandler pdfHandler;

        public int ActivePage { get; set; } = 0;
        public ViewModel Model { get; set; } = new ViewModel();

        public Main(IExcelHandler excelHandler, IPDFHandler pdfHandler)
        {
            InitializeComponent();
            this.excelHandler = excelHandler;
            this.pdfHandler = pdfHandler;
        }

        private void BtnBrowseDetailExcel_Click(object sender, EventArgs e)
        {
            OpenExcelFileDialogue.ShowDialog(this);
            LocDetailedExcel.Text = OpenExcelFileDialogue.FileName;
            Model.FirstPage.DetailExcelLoc = OpenExcelFileDialogue.FileName;
        }

        private void BtnBrowseGroupExcel_Click(object sender, EventArgs e)
        {
            OpenExcelFileDialogue.ShowDialog(this);
            LocGroupExcel.Text = OpenExcelFileDialogue.FileName;
            Model.FirstPage.GroupExcelLoc = OpenExcelFileDialogue.FileName;
        }

        private void BtnBrowseSinglePDF_Click(object sender, EventArgs e)
        {
            SavePDFFileDialogue.ShowDialog(this);
            LocSinglePDF.Text = SavePDFFileDialogue.FileName;
            Model.FirstPage.SinglePDFLoc = SavePDFFileDialogue.FileName;
        }

        private void BtnBrowseMultiplePDF_Click(object sender, EventArgs e)
        {
            SavePDFFileDialogue.ShowDialog(this);
            LocMultiplePDF.Text = SavePDFFileDialogue.FileName;
            Model.FirstPage.MultiPDFLoc = SavePDFFileDialogue.FileName;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            var cpuCores = Environment.ProcessorCount;
            for (int i = cpuCores; i >= 1; i--)
            {
                CPUCores.Items.Add(i);
            }
            CPUCores.SelectedIndex = (CPUCores.Items.Count) / 2;
        }

        private async void BtnNext_Click(object sender, EventArgs e)
        {
        }


        private async Task ValidateAsync()
        {
            //Validate Excel
            ShowSildeLoader("Preparing...");
            await Task.Delay(100);

            if (Model.FirstPage.DetailExcelLoc == string.Empty)
            {
                MessageBox.Show("Please click browse to choose detail Excel file", "Validation failed");
                return;
            }
            if (Model.FirstPage.GroupExcelLoc == string.Empty)
            {
                MessageBox.Show("Please click browse to choose group Excel file", "Validation failed");
                return;
            }
            //if (Model.FirstPage.SinglePDFLoc == string.Empty)
            //{
            //    MessageBox.Show("Please click browse to choose a location to save single page PDF statements", "Validation failed");
            //    return;
            //}
            //if (Model.FirstPage.MultiPDFLoc == string.Empty)
            //{
            //    MessageBox.Show("Please click browse to choose a location to save multi page PDF statements", "Validation failed");
            //    return;
            //}

            try
            {
                ShowSildeLoader("Reading details excel...");
                await Task.Delay(100);
                Model.FirstPage.DetailTable = excelHandler.ParseExcelDataIntoDataTable(Model.FirstPage.DetailExcelLoc, 0);
            }
            catch (Exception)
            {
                throw new ExcelException(
                   $"Failed reading {Path.GetFileName(Model.FirstPage.DetailExcelLoc)}",
                   "The file might be invalid, inaccessible or locked",
                   "Ensure no other application opened Excel file, If opened it can lock the file. Close it and try if any",
                   "Make sure data is in Sheet-1",
                   "Make sure row/column structure is correct",
                   "Make sure Excel file is not created with a very old Excel",
                   "If using a downloaded Google Sheet, Choose xslx option when downloading");
            }
            finally
            {
                StopSildeLoader();
            }

            try
            {
                ShowSildeLoader("Reading group excel...");
                await Task.Delay(100);
                Model.FirstPage.GroupTable = excelHandler.ParseExcelDataIntoDataTable(Model.FirstPage.GroupExcelLoc, 0);
            }
            catch (Exception)
            {
                throw new ExcelException(
                    $"Failed reading {Path.GetFileName(Model.FirstPage.GroupExcelLoc)}",
                    "There is an issue in the Excel file",
                    "Make sure data is in Sheet-1",
                    "Make sure row/column structure is correct",
                    "Make sure Excel file is not created with a very old Excel",
                    "If using a downloaded Google Sheet, Choose xslx option when downloading");
            }
            finally
            {
                StopSildeLoader();
            }

            Pager.SelectTab(1);
        }

        private async Task DisplayErrorPage(GenericException ex)
        {
            ErrorViewer.Rtf = GenerateRTF(ex.Heading, ex.SubHeading, ex.Options);
            Pager.SelectTab(4);
            LogContext.PushProperty("ActivePage", ActivePage);
            LogContext.PushProperty("ViewModel", Model);
        }

        private void ShowSildeLoader(string status)
        {
            SideloadStatus.Text = status;
            SideLoader.Visible = true;
        }

        private void StopSildeLoader()
        {
            SideloadStatus.Text = "Done";
            SideLoader.Visible = false;
        }

        private string GenerateRTF(string heading, string subheading, List<string> bulletPoints)
        {
            string rtf = @"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Tahoma;}}
{\colortbl;\red255\green0\blue0;}
{\fs34\cf1 " + heading + @"}\par
{\fs28 " + subheading + @"}\par

\par{\fs20 Possible Reasons / Remedies}\par\par";

            for (int i = 0; i < bulletPoints.Count && i < 3; i++)
            {
                rtf += @"{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb}} \fs28 " + bulletPoints[i] + @"\par";
            }

            rtf += @"\pard\fi-360\li720\sl240\slmult0{\fs14\par}}";

            return rtf;
        }


        //public static bool IsChromeInstalled()
        //{
        //    // Try to check the registry key for Chrome
        //    try
        //    {
        //        string chromeKey = @"SOFTWARE\Google\Chrome";
        //        using (RegistryKey key = Registry.LocalMachine.OpenSubKey(chromeKey))
        //        {
        //            if (key != null)
        //            {
        //                return true;
        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        // Ignore any exceptions due to permission issues and fall back to searching for the Chrome executable
        //    }

        //    // If we couldn't check the registry, search for the Chrome executable in common installation folders
        //    string[] commonFolders = {
        //        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        //        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
        //        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        //    };
        //    foreach (string folder in commonFolders)
        //    {
        //        string chromePath = Path.Combine(folder, "Google", "Chrome", "Application", "chrome.exe");
        //        if (File.Exists(chromePath))
        //        {
        //            return true;
        //        }
        //    }

        //    // If we couldn't find Chrome, return false
        //    return false;
        //}

        private void ChkNotify_CheckedChanged(object sender, EventArgs e)
        {
            Model.SecondPage.NotifyWhenCompleted = ChkNotify.Checked;
        }

        private void ChkOpen_CheckedChanged(object sender, EventArgs e)
        {
            Model.SecondPage.OpenWhenCompleted = ChkOpen.Checked;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void BtnRestart_Click(object sender, EventArgs e)
        {
            Pager.SelectTab(0);
            Model.FirstPage.DetailExcelLoc = string.Empty;
            Model.FirstPage.GroupExcelLoc = string.Empty;
            Model.FirstPage.SinglePDFLoc = string.Empty;
            Model.FirstPage.MultiPDFLoc = string.Empty;
            Model.SecondPage.NotifyWhenCompleted = true;
            Model.SecondPage.OpenWhenCompleted = false;
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void BtnToFirstPage_Click(object sender, EventArgs e)
        {
            Pager.SelectTab(0);
        }

        private async void BtnToOptions_Click(object sender, EventArgs e)
        {
            try
            {
                await ValidateAsync();
            }
            catch (GenericException ex)
            {
                DisplayErrorPage(ex);
                Log.Error(ex, nameof(ExcelException));
            }
        }

        private void BtnToProcessing_Click(object sender, EventArgs e)
        {
            if (Model.SecondPage.CPUCores == Environment.ProcessorCount)
            {
                DialogResult result = MessageBox.Show(
                     $"You're allowing the app to use {Model.SecondPage.CPUCores} CPU cores on this PC for processing.\n\n" +
                     $"Allowing to use more CPU cores makes processing a lot faster. However, it's important to note that this can put your system to its peak limit.\n\n" +
                     "For better safety, please save all your works before starting.\n\n" +
                     "Are you sure you want to proceed with full throttle?",
                     "Save all your unsaved works",
                     MessageBoxButtons.OKCancel,
                     MessageBoxIcon.Information
                 );
                if (result == DialogResult.OK)
                {
                    Pager.SelectTab(2);
                    Worker.RunWorkerAsync();
                }
            }
            else
            {
                Pager.SelectTab(2);
                Worker.RunWorkerAsync();
            }
        }

        public void ResetOutputFolder()
        {
            try
            {
                var folderPath = "multi";
                if (Directory.Exists(folderPath))
                {
                    Directory.Delete(folderPath, true);
                }

                Directory.CreateDirectory(folderPath);
            }
            catch (Exception ex)
            {
            }
        }

        private void Worker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            ProgressBar.Invoke(() =>
            {
                ProgressBar.Style = ProgressBarStyle.Marquee;
                ProgressBar.MarqueeAnimationSpeed = 10;
            });

            var allStatements = new List<FinancialStatement>();

            //Take each detail
            foreach (DataRow group in Model.FirstPage.GroupTable.Rows)
            {
                //Take first item from groupTable
                var primaryKey = group["ReceiptID"];
                //Check if there is a match in detailsTable
                var matchingRows = Model.FirstPage.DetailTable.Rows.Cast<DataRow>()
                                    .Where(row => row["ReceiptID"].ToString() == primaryKey.ToString())
                                    .ToList();

                var ficialYear = group.Table.Columns.Contains("Fiscal Year") && int.TryParse(group["Fiscal Year"].ToString(), out int result) ?
                    result : DateTime.Now.Year;

                allStatements.Add(new FinancialStatement
                {
                    ReceiptID = group["ReceiptID"].ToString(),
                    FullName = group.Table.Columns.Contains("Full Name") ? group["Full Name"].ToString() : string.Empty,
                    AddressLine1 = group.Table.Columns.Contains("Address Line 1") ? group["Address Line 1"].ToString() : string.Empty,
                    City = group.Table.Columns.Contains("City") ? group["City"].ToString() : string.Empty,
                    State = group.Table.Columns.Contains("State") ? group["State"].ToString() : string.Empty,
                    ZIPCode = group.Table.Columns.Contains("ZIP Code") ? group["ZIP Code"].ToString() : string.Empty,
                    IMBarcode = group.Table.Columns.Contains("IM Barcode") ? group["IM Barcode"].ToString() : string.Empty,
                    QRContent = group.Table.Columns.Contains("QRcontent") ? group["QRcontent"].ToString() : string.Empty,
                    TraySort = group.Table.Columns.Contains("Tray-Sort") ? group["Tray-Sort"].ToString() : string.Empty,
                    Pages = group.Table.Columns.Contains("Pages") && int.TryParse(group["Pages"].ToString(), out int pages) ? pages : 0,
                    Total = group.Table.Columns.Contains("Total") && decimal.TryParse(group["Total"].ToString(), out decimal total) ? total : 0,
                    FicialYear = ficialYear,
                    CRST_Sort = group.Table.Columns.Contains("CRST_Sort") && int.TryParse(group["CRST_Sort"].ToString(), out int sort) ? sort : 0,

                    Payments = matchingRows.Select(row => new Payment
                    {
                        ReceiptID = row["ReceiptID"].ToString(),
                        Date = row.Table.Columns.Contains("Date") ? row["Date"].ToString() : string.Empty,
                        Check = row.Table.Columns.Contains("Check #") ? row["Check #"].ToString() : string.Empty,
                        Amount = row.Table.Columns.Contains("Amount") && decimal.TryParse(row["Amount"].ToString(), out decimal amount) ? amount : 0
                    }).ToList()
                });
            }

            ResetOutputFolder();

            var statements = allStatements.OrderBy(x => x.CRST_Sort).ToList();
            statements = statements.Take(2).ToList();

            MultiPageProcessingAndMerging(Model.FirstPage.MultiPDFLoc, statements);

            //Completed
            if (Model.SecondPage.NotifyWhenCompleted)
            {
                ProgressLog.Invoke(() =>
                {
                    TrayNotification.ShowBalloonTip(8000);
                });
            }
            if (Model.SecondPage.OpenWhenCompleted)
            {
                var folderPath = Path.GetDirectoryName(Model.FirstPage.DetailExcelLoc);
                if (Directory.Exists(folderPath))
                {
                    Process.Start("explorer.exe", folderPath);
                }
            }

            WriteProgressText("Merge completed...");
            WriteLog("Merge completed");

            Pager.Invoke(() =>
            {
                Pager.SelectTab(3);
            });
        }

        //private void MultiPageProcessing(List<FinancialStatement> multiPageStatements)
        //{
        //    ProgressBar.Invoke(() =>
        //    {
        //        ProgressBar.Style = ProgressBarStyle.Blocks;
        //    });
        //    int i = 1;
        //    Parallel.ForEach(multiPageStatements.ToList(),
        //                    new ParallelOptions
        //                    {
        //                        MaxDegreeOfParallelism = Environment.ProcessorCount
        //                    },
        //                    async s =>
        //                    {
        //                        WriteLog($"Scheduled '{s.FullName}'");
        //                        MoveProgress();

        //                        var model = new Core.Statement.PageResources
        //                        {
        //                            FinancialStatement = s,
        //                            Images = new Dictionary<string, string>()
        //                        };
        //                        model.Images.Add("logo", Convert.ToBase64String(File.ReadAllBytes("logo.jpg")));
        //                        model.Images.Add("qrCode", XMLProcessor.GenerateQRCode(model.FinancialStatement.QRContent));

        //                        var html = await RazorTemplateEngine.RenderAsync("/MultiPageTheme.cshtml", model);

        //                        if (!Directory.Exists($"multi"))
        //                        {
        //                            Directory.CreateDirectory($"multi");
        //                        }

        //                        var pdfName = $"multi/{s.CRST_Sort}.pdf";
        //                        PDFEngine.Generate(html, pdfName);
        //                        var noOfPages = GetNumberOfPages(pdfName);
        //                        if (!Directory.Exists($"multi/{noOfPages} Page"))
        //                        {
        //                            Directory.CreateDirectory($"multi/{noOfPages} Page");
        //                        }
        //                        File.Move(pdfName, $"multi/{noOfPages} Page/{Path.GetFileName(pdfName)}");


        //                        WriteLog($"Completed '{s.FullName}'");
        //                        WriteProgressText($"Working on '{s.FullName}'");
        //                        MoveProgress();
        //                    });
        //}

        private int GetNumberOfPages(string pdfName)
        {
            try
            {
                using (PdfDocument document = PdfReader.Open(pdfName, PdfDocumentOpenMode.ReadOnly))
                {
                    return document.PageCount;
                }
            }
            catch (Exception ex)
            {
                return 8;
            }
        }
        static string[] GetAllSubdirectories(string rootDirectory)
        {
            try
            {
                string[] subdirectories = Directory.GetDirectories(rootDirectory, "*", SearchOption.AllDirectories)
                                                    .OrderBy(d => d)
                                                    .ToArray();
                return subdirectories;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return null;
            }
        }

        static string GetLeafFolderName(string folderPath)
        {
            try
            {
                string[] pathParts = folderPath.Split('\\');
                string leafFolderName = pathParts[pathParts.Length - 1];
                return leafFolderName;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return null;
            }
        }

        //private void SinglePageMerging(string singlePdfLocation, List<FinancialStatement> statements)
        //{
        //    ProgressBar.Invoke(() =>
        //    {
        //        ProgressBar.Style = ProgressBarStyle.Marquee;
        //        ProgressBar.MarqueeAnimationSpeed = 10;
        //    });

        //    Parallel.Invoke(() =>
        //    {
        //        string[] inputFiles = Directory.GetFiles("outputs/", "*.pdf")
        //                             .OrderBy(f => f)
        //                             .ToArray();

        //        // Create a new PDF document object
        //        var outputDocument = new PdfDocument();

        //        // Register the encoding provider
        //        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        //        int j = 0;
        //        // Loop through all the input PDF files and add their pages to the output document
        //        foreach (string inputFile in inputFiles)
        //        {
        //            int i = 0;
        //            var recieptId = statements[j].ReceiptID.ToUpper();

        //            // Open the input PDF document
        //            var inputDocument = PdfReader.Open(inputFile, PdfDocumentOpenMode.Import);

        //            // Loop through each page in the input document and add it to the output document
        //            foreach (var inputPage in inputDocument.Pages)
        //            {
        //                // Create a new page in the output document
        //                var outputPage = outputDocument.AddPage(inputPage);

        //                // Get the PdfContentByte object for the page
        //                XGraphics gfx = XGraphics.FromPdfPage(outputPage);
        //                // Create the font and brush
        //                XFont font = new XFont("Times New Roman", 9, XFontStyle.Regular);
        //                XBrush brush = XBrushes.Black;

        //                var footerText = $"{recieptId}                                        Produced by Tzedokos App - https://ahavastzedokos.com                                        Page {i + 1} of {inputDocument.Pages.Count}";

        //                // Get the text size
        //                XSize size = gfx.MeasureString(footerText, font);
        //                // Create the formatter
        //                XTextFormatter tf = new XTextFormatter(gfx);
        //                // Calculate the x and y coordinates
        //                double x = (outputPage.Width - size.Width) / 2;
        //                double y = outputPage.Height - size.Height - 20;
        //                // Draw the text
        //                tf.DrawString(footerText, font, brush, new XRect(x, y, size.Width, size.Height), XStringFormats.TopLeft);

        //                i++;
        //            }

        //            // Close the input PDF document
        //            inputDocument.Close();
        //            j++;
        //        }

        //        // Save the merged PDF document to the output file
        //        if (statements.Any())
        //        {
        //            outputDocument.Save(singlePdfLocation);
        //        }

        //        // Close the output PDF document
        //        outputDocument.Close();
        //    });



        //    ProgressBar.Invoke(() =>
        //    {
        //        ProgressBar.Style = ProgressBarStyle.Blocks;
        //        ProgressBar.MarqueeAnimationSpeed = 10;
        //    });

        //    ResetOutputFolder();
        //}

        private void MultiPageProcessingAndMerging(string singlePdfLocation, List<FinancialStatement> statements)
        {
            var fileNo = 1;

            WriteProgressSubText($"Activating CPU cores");
            // Invoke the ProgressBar update on the UI thread
            ProgressBar.Invoke(() =>
            {
                ProgressBar.Style = ProgressBarStyle.Blocks;
                ProgressBar.Maximum = statements.Count + 1;
                ProgressBar.Value = fileNo;
            });

            // Register the encoding provider
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Use Parallel.ForEach for better load balancing
            Parallel.ForEach(statements.ToList(),
                            new ParallelOptions
                            {
                                MaxDegreeOfParallelism = Model.SecondPage.CPUCores
                            },
                            async s =>
                            {   // Create a new PDF document object
                                var outputDocument = new PdfDocument();

                                WriteProgressText($"{fileNo}/{statements.Count}: Working on '{s.FullName}'");
                                WriteLog($"Scheduled '{s.FullName}'");

                                WriteProgressSubText($"Generating QR code...");
                                var model = new Core.Statement.PageResources
                                {
                                    FinancialStatement = s,
                                    Images = new Dictionary<string, string>()
                                };
                                model.Images.Add("logo", Convert.ToBase64String(File.ReadAllBytes("logo.jpg")));
                                model.Images.Add("qrCode", XMLProcessor.GenerateQRCode(model.FinancialStatement.QRContent));

                                WriteProgressSubText($"Building PDF structure...");
                                var html = await RazorTemplateEngine.RenderAsync("/MultiPageTheme.cshtml", model);

                                if (!Directory.Exists($"multi"))
                                {
                                    Directory.CreateDirectory($"multi");
                                }

                                var pdfName = $"multi/{s.CRST_Sort}.pdf";

                                WriteProgressSubText($"Rendering PDF...");
                                PDFEngine.Generate(html, pdfName);

                                var noOfPages = GetNumberOfPages(pdfName);
                                if (!Directory.Exists($"multi/{noOfPages} Page"))
                                {
                                    Directory.CreateDirectory($"multi/{noOfPages} Page");
                                }

                                string destinationFile = $"multi/{noOfPages} Page/{Path.GetFileName(pdfName)}";
                                if (File.Exists(destinationFile))
                                {
                                    File.Delete(destinationFile);
                                }
                                File.Move(pdfName, destinationFile);

                                pdfName = $"multi/{noOfPages} Page/{Path.GetFileName(pdfName)}";

                                WriteProgressSubText($"Imprinting footer & page number...");
                                // Open the input PDF document
                                using (var inputDocument = PdfReader.Open(pdfName, PdfDocumentOpenMode.Import))
                                {

                                    // Loop through each page in the input document and add it to the output document
                                    for (int i = 0; i < inputDocument.Pages.Count; i++)
                                    {
                                        var inputPage = inputDocument.Pages[i];

                                        // Import the page from the input document to the output document
                                        var outputPage = outputDocument.AddPage(inputPage);

                                        // Get the PdfContentByte object for the page
                                        XGraphics gfx = XGraphics.FromPdfPage(outputPage);
                                        // Create the font and brush
                                        XFont font = new XFont("Times New Roman", 9, XFontStyle.Regular);
                                        XBrush brush = XBrushes.Black;

                                        var footerText = $"{s.ReceiptID.ToUpper()}                                        Produced by Tzedokos App - https://ahavastzedokos.com                                        Page {i + 1} of {inputDocument.Pages.Count}";

                                        // Get the text size
                                        XSize size = gfx.MeasureString(footerText, font);
                                        // Create the formatter
                                        XTextFormatter tf = new XTextFormatter(gfx);
                                        // Calculate the x and y coordinates
                                        double x = (outputPage.Width - size.Width) / 2;
                                        double y = outputPage.Height - size.Height - 20;
                                        // Draw the text
                                        tf.DrawString(footerText, font, brush, new XRect(x, y, size.Width, size.Height), XStringFormats.TopLeft);
                                    }

                                    // Save the output PDF document to the file
                                    lock (outputDocument)
                                    {
                                        outputDocument.Save(pdfName);
                                    }

                                    // Close the output PDF document
                                    outputDocument.Close();
                                }

                                WriteLog($"Completed '{s.FullName}'");
                                MoveProgress();
                                fileNo++;
                                WriteProgressSubText($"Done");
                            });

            // Get all subdirectories
            var pageDirectories = GetAllSubdirectories("multi");

            // Use Parallel.ForEach for better load balancing
            Parallel.ForEach(pageDirectories,
                new ParallelOptions
                {
                    MaxDegreeOfParallelism = Model.SecondPage.CPUCores
                },
                (pageDirectory) =>
            {
                // Get all PDF files in the directory, ordered by filename
                string[] inputFiles = Directory.GetFiles(pageDirectory, "*.pdf")
                                .OrderBy(f => f)
                                .ToArray();

                // Create a new PDF document object
                var outputDocument = new PdfDocument();

                // Register the encoding provider
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                // Loop through all the input PDF files and add their pages to the output document
                foreach (string inputFile in inputFiles)
                {
                    // Open the input PDF document
                    using (var inputDocument = PdfReader.Open(inputFile, PdfDocumentOpenMode.Import))
                    {
                        // Loop through each page in the input document and add it to the output document
                        for (int i = 0; i < inputDocument.Pages.Count; i++)
                        {
                            var inputPage = inputDocument.Pages[i];

                            // Import the page from the input document to the output document
                            var outputPage = outputDocument.AddPage(inputPage);
                        }
                    }
                }

                outputDocument.Save(Path.Combine(Path.GetDirectoryName(Model.FirstPage.DetailExcelLoc), $"Merged {GetLeafFolderName(pageDirectory)} PDFs - In CRST Order.pdf"));

                // Close the output PDF document
                outputDocument.Close();
            });

            ResetOutputFolder();
        }


        //private void MultiPageMerging(string singlePdfLocation, List<FinancialStatement> statements)
        //{
        //    // Invoke the ProgressBar update on the UI thread
        //    ProgressBar.Invoke(() =>
        //    {
        //        ProgressBar.Style = ProgressBarStyle.Marquee;
        //        ProgressBar.MarqueeAnimationSpeed = 10;
        //    });

        //    // Get all subdirectories
        //    var pageDirectories = GetAllSubdirectories("multi");

        //    // Use Parallel.ForEach for better load balancing
        //    Parallel.ForEach(pageDirectories, (pageDirectory) =>
        //    {
        //        // Get all PDF files in the directory, ordered by filename
        //        string[] inputFiles = Directory.GetFiles(pageDirectory, "*.pdf")
        //                        .OrderBy(f => f)
        //                        .ToArray();

        //        // Create a new PDF document object
        //        var outputDocument = new PdfDocument();

        //        // Register the encoding provider
        //        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        //        // Use another Parallel.ForEach to process each file in parallel
        //        Parallel.ForEach(inputFiles, (inputFile, loopState, j) =>
        //        {
        //            var recieptId = statements[(int)j].ReceiptID.ToUpper();

        //            // Open the input PDF document
        //            using (var inputDocument = PdfReader.Open(inputFile, PdfDocumentOpenMode.Import))
        //            {
        //                // Loop through each page in the input document and add it to the output document
        //                for (int i = 0; i < inputDocument.Pages.Count; i++)
        //                {
        //                    var inputPage = inputDocument.Pages[i];

        //                    // Create a new page in the output document
        //                    var outputPage = outputDocument.AddPage(inputPage);

        //                    // Get the PdfContentByte object for the page
        //                    XGraphics gfx = XGraphics.FromPdfPage(outputPage);
        //                    // Create the font and brush
        //                    XFont font = new XFont("Times New Roman", 9, XFontStyle.Regular);
        //                    XBrush brush = XBrushes.Black;

        //                    var footerText = $"{recieptId}                                        Produced by Tzedokos App - https://ahavastzedokos.com                                        Page {i + 1} of {inputDocument.Pages.Count}";

        //                    // Get the text size
        //                    XSize size = gfx.MeasureString(footerText, font);
        //                    // Create the formatter
        //                    XTextFormatter tf = new XTextFormatter(gfx);
        //                    // Calculate the x and y coordinates
        //                    double x = (outputPage.Width - size.Width) / 2;
        //                    double y = outputPage.Height - size.Height - 20;
        //                    // Draw the text
        //                    tf.DrawString(footerText, font, brush, new XRect(x, y, size.Width, size.Height), XStringFormats.TopLeft);
        //                }
        //            }

        //            // Save the merged PDF document to the output file
        //            if (statements.Any())
        //            {
        //                outputDocument.Save(Path.Combine(Path.GetDirectoryName(singlePdfLocation), $"{GetLeafFolderName(pageDirectory)} PDFs Merged In CRST Order.pdf"));
        //            }
        //        });

        //        // Close the output PDF document
        //        outputDocument.Close();
        //    });

        //    // Invoke the ProgressBar update on the UI thread
        //    ProgressBar.Invoke(() =>
        //    {
        //        ProgressBar.Style = ProgressBarStyle.Blocks;
        //        ProgressBar.MarqueeAnimationSpeed = 10;
        //    });

        //    ResetOutputFolder();
        //}


        //private void SinglePageProcessing(List<FinancialStatement> singlePageStatements)
        //{
        //    ProgressBar.Invoke(() =>
        //    {
        //        ProgressBar.Style = ProgressBarStyle.Blocks;
        //    });
        //    int i = 1;
        //    Parallel.ForEach(singlePageStatements.ToList(),
        //        new ParallelOptions
        //        {
        //            MaxDegreeOfParallelism = Model.SecondPage.CPUCores
        //        },
        //        async s =>
        //        {
        //            WriteLog($"Scheduled '{s.FullName}'");
        //            MoveProgress();

        //            var model = new Core.Statement.PageResources
        //            {
        //                FinancialStatement = s,
        //                Images = new Dictionary<string, string>()
        //            };
        //            model.Images.Add("logo", Convert.ToBase64String(File.ReadAllBytes("logo.jpg")));
        //            model.Images.Add("qrCode", XMLProcessor.GenerateQRCode(model.FinancialStatement.QRContent));

        //            var html = await RazorTemplateEngine.RenderAsync("/Theme.cshtml", model);

        //            PDFEngine.Generate(html, $"outputs/{s.TraySort.Replace("/", "_")}.pdf");

        //            WriteLog($"Completed '{s.FullName}'");
        //            WriteProgressText($"Working on '{s.FullName}'");
        //            MoveProgress();
        //        });
        //}

        private void MoveProgress()
        {
            try
            {
                if (ProgressBar.Value < ProgressBar.Maximum)
                {
                    ProgressBar.Invoke(() =>
                    {
                        ProgressBar.Value++;
                    });
                }
            }
            catch (Exception)
            {
            }
        }

        private void WriteLog(string message)
        {
            ProgressLog.Invoke(() =>
            {
                ProgressLog.Items.Insert(0, message);
            });
        }



        private void WriteProgressText(string message)
        {
            ProgressLog.Invoke(() =>
            {
                ProgressText.Text = message;
            });
        }

        private void WriteProgressSubText(string message)
        {
            ProgressLog.Invoke(() =>
            {
                ProgressSubText.Text = message;
            });
        }

        private void CPUCores_SelectedIndexChanged(object sender, EventArgs e)
        {
            Model.SecondPage.CPUCores = int.Parse(CPUCores.SelectedItem.ToString());
        }
    }
}