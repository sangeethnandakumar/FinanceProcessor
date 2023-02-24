using FinanceProcessor.Application.Handlers;
using FinanceProcessor.Core.Exceptions;
using FinanceProcessor.Core.Stetement;
using Microsoft.Win32;
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
            CPUCores.SelectedIndex = 0;
        }

        private async void BtnNext_Click(object sender, EventArgs e)
        {
        }


        private async Task ValidateAsync()
        {
            //Validate Excel
            ShowSildeLoader("Preparing...");
            await Task.Delay(500);

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
            if (Model.FirstPage.SinglePDFLoc == string.Empty)
            {
                MessageBox.Show("Please click browse to choose a location to save single page PDF statements", "Validation failed");
                return;
            }
            if (Model.FirstPage.MultiPDFLoc == string.Empty)
            {
                MessageBox.Show("Please click browse to choose a location to save multi page PDF statements", "Validation failed");
                return;
            }

            try
            {
                ShowSildeLoader("Reading details excel...");
                await Task.Delay(250);
                Model.FirstPage.DetailTable = excelHandler.ParseExcelDataIntoDataTable(Model.FirstPage.DetailExcelLoc, 0);
            }
            catch (Exception)
            {
                throw new ExcelException(
                   $"Failed reading {Path.GetFileName(Model.FirstPage.DetailExcelLoc)}",
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

            try
            {
                ShowSildeLoader("Reading group excel...");
                await Task.Delay(250);
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

            ShowSildeLoader("Checking compatibility...");
            await Task.Delay(250);
            if (!IsChromeInstalled())
            {
                throw new ExcelException(
                  $"Cannot detect Google Chrome in this computer",
                  "Is Chrome installed in this PC?",
                  "Try running this app as administrator",
                  "If not installed, Try installing Chrome",
                  "If already installed, Try updating Chrome");
            }

            StopSildeLoader();
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


        public static bool IsChromeInstalled()
        {
            // Try to check the registry key for Chrome
            try
            {
                string chromeKey = @"SOFTWARE\Google\Chrome";
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(chromeKey))
                {
                    if (key != null)
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
                // Ignore any exceptions due to permission issues and fall back to searching for the Chrome executable
            }

            // If we couldn't check the registry, search for the Chrome executable in common installation folders
            string[] commonFolders = {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            };
            foreach (string folder in commonFolders)
            {
                string chromePath = Path.Combine(folder, "Google", "Chrome", "Application", "chrome.exe");
                if (File.Exists(chromePath))
                {
                    return true;
                }
            }

            // If we couldn't find Chrome, return false
            return false;
        }

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
            DialogResult result = MessageBox.Show("The system is going to run very intensive tasks that can take sometime to complete. Please save all the works before proceeding.", "Save all your unsaved works", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            if (result == DialogResult.OK)
            {
                Pager.SelectTab(2);
                Worker.RunWorkerAsync();
            }
        }

        public void ResetOutputFolder()
        {
            string folderPath = "outputs/";
            if (!Directory.Exists(folderPath))
            {
                // Create the directory if it does not exist
                Directory.CreateDirectory(folderPath);
            }
            else
            {
                // Clear all files in the directory if it exists
                DirectoryInfo directory = new DirectoryInfo(folderPath);
                foreach (FileInfo file in directory.GetFiles())
                {
                    file.Delete();
                }
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
                var primaryKey = group.ItemArray[0];
                //Check if there is a match in detailsTable
                var matchingRows = Model.FirstPage.DetailTable.Rows.Cast<DataRow>()
                                    .Where(row => row.ItemArray[0].ToString() == primaryKey.ToString())
                                    .ToList();

                allStatements.Add(new FinancialStatement
                {
                    ReceiptID = group.ItemArray[0].ToString(),
                    FullName = group.ItemArray[1].ToString(),
                    AddressLine1 = group.ItemArray[2].ToString(),
                    City = group.ItemArray[3].ToString(),
                    State = group.ItemArray[4].ToString(),
                    ZIPCode = group.ItemArray[5].ToString(),
                    IMBarcode = group.ItemArray[6].ToString(),
                    QRContent = group.ItemArray[7].ToString(),
                    TraySort = group.ItemArray[8].ToString(),
                    Pages = int.Parse(group.ItemArray[9].ToString()),
                    Total = decimal.Parse(group.ItemArray[10].ToString()),

                    Payments = matchingRows.Select(row => new Payment
                    {
                        ReceiptID = row.ItemArray[0].ToString(),
                        Date = DateTime.Parse(row.ItemArray[1].ToString()),
                        Check = row.ItemArray[2].ToString(),
                        CheckNumber = row.ItemArray[3].ToString(),
                        Amount = decimal.Parse(row.ItemArray[4].ToString())
                    }).ToList()
                });
            }
            ResetOutputFolder();

            //DATA COLLECTION
            var singlePageStatements = allStatements.Where(s => s.Payments.Count <= 5).ToList();
            singlePageStatements.Sort((a, b) => string.Compare(a.TraySort, b.TraySort));
            singlePageStatements = singlePageStatements.Take(10).ToList();

            var multiPageStatements = allStatements.Where(s => s.Payments.Count > 5).ToList();
            multiPageStatements.Sort((a, b) => string.Compare(a.TraySort, b.TraySort));
            multiPageStatements = multiPageStatements.Take(10).ToList();

            ProgressBar.Invoke(() =>
            {
                ProgressBar.Style = ProgressBarStyle.Blocks;
                ProgressBar.Maximum = ((singlePageStatements.Count + singlePageStatements.Count) * 2) + 1;
            });

            //SINGLE PAGE PROCESSING    
            SinglePageProcessing(singlePageStatements);
            WriteProgressText("Merging PDF pages...");
            WriteLog("Merging PDF pages...");
            PageMergng(Model.FirstPage.SinglePDFLoc, singlePageStatements);

            //MULTI PAGE PROCESSING      
            MultiPageProcessing(multiPageStatements);
            WriteProgressText("Merging PDF pages...");
            WriteLog("Merging PDF pages...");
            PageMergng(Model.FirstPage.MultiPDFLoc, multiPageStatements);

            //Completed

            if (Model.SecondPage.NotifyWhenCompleted)
            {
                ProgressLog.Invoke(() =>
                {
                    TrayNotification.ShowBalloonTip(5000);
                });
            }
            if (Model.SecondPage.OpenWhenCompleted)
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(Model.FirstPage.SinglePDFLoc)
                {
                    UseShellExecute = true
                };
                Process.Start(startInfo);
                ProcessStartInfo startInfo2 = new ProcessStartInfo(Model.FirstPage.MultiPDFLoc)
                {
                    UseShellExecute = true
                };
                Process.Start(startInfo2);
            }

            WriteProgressText("Merge completed...");
            WriteLog("Merge completed");

            Pager.Invoke(() =>
            {
                Pager.SelectTab(3);
            });
        }

        private void MultiPageProcessing(List<FinancialStatement> multiPageStatements)
        {
            ProgressBar.Invoke(() =>
            {
                ProgressBar.Style = ProgressBarStyle.Blocks;
            });
            int i = 1;
            Parallel.ForEach(multiPageStatements.ToList(),
                            new ParallelOptions
                            {
                                MaxDegreeOfParallelism = Model.SecondPage.CPUCores
                            },
                            async s =>
                            {
                                WriteLog($"Scheduled '{s.FullName}'");
                                MoveProgress();

                                var model = new Core.Statement.PageResources
                                {
                                    FinancialStatement = s,
                                    Images = new Dictionary<string, string>()
                                };
                                model.Images.Add("logo", Convert.ToBase64String(File.ReadAllBytes("logo.jpg")));
                                model.Images.Add("qrCode", XMLProcessor.GenerateQRCode(model.FinancialStatement.QRContent));

                                var html = await RazorTemplateEngine.RenderAsync("/MultiPageTheme.cshtml", model);
                                PDFEngine.Generate(html, $"outputs/{i++}.pdf");

                                WriteLog($"Completed '{s.FullName}'");
                                WriteProgressText($"Working on '{s.FullName}'");
                                MoveProgress();
                            });
        }

        private void PageMergng(string singlePdfLocation, List<FinancialStatement> statements)
        {
            ProgressBar.Invoke(() =>
            {
                ProgressBar.Style = ProgressBarStyle.Marquee;
                ProgressBar.MarqueeAnimationSpeed = 10;
            });

            Parallel.Invoke(() =>
            {
                string[] inputFiles = Directory.GetFiles("outputs/", "*.pdf")
                                     .OrderBy(f => f)
                                     .ToArray();

                // Create a new PDF document object
                var outputDocument = new PdfDocument();

                // Register the encoding provider
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                int j = 0;
                int i = 0;
                // Loop through all the input PDF files and add their pages to the output document
                foreach (string inputFile in inputFiles)
                {
                    var recieptId = statements[i].ReceiptID.ToUpper();

                    // Open the input PDF document
                    var inputDocument = PdfReader.Open(inputFile, PdfDocumentOpenMode.Import);

                    // Loop through each page in the input document and add it to the output document
                    foreach (var inputPage in inputDocument.Pages)
                    {
                        // Create a new page in the output document
                        var outputPage = outputDocument.AddPage(inputPage);

                        // Get the PdfContentByte object for the page
                        XGraphics gfx = XGraphics.FromPdfPage(outputPage);
                        // Create the font and brush
                        XFont font = new XFont("Times New Roman", 9, XFontStyle.Regular);
                        XBrush brush = XBrushes.Black;
                        
                        var footerText = $"{recieptId}                                        Produced by Tzedokos App - https://ahavastzedokos.com                                        Page {i+1} of {inputDocument.Pages.Count}";
                        
                        // Get the text size
                        XSize size = gfx.MeasureString(footerText, font);
                        // Create the formatter
                        XTextFormatter tf = new XTextFormatter(gfx);
                        // Calculate the x and y coordinates
                        double x = (outputPage.Width - size.Width) / 2;
                        double y = outputPage.Height - size.Height - 20;
                        // Draw the text
                        tf.DrawString(footerText, font, brush, new XRect(x, y, size.Width, size.Height), XStringFormats.TopLeft);

                        i++;
                    }

                    // Close the input PDF document
                    inputDocument.Close();
                    j++;
                }

                // Save the merged PDF document to the output file
                outputDocument.Save(singlePdfLocation);

                // Close the output PDF document
                outputDocument.Close();
            });



            ProgressBar.Invoke(() =>
            {
                ProgressBar.Style = ProgressBarStyle.Blocks;
                ProgressBar.MarqueeAnimationSpeed = 10;
            });

            ResetOutputFolder();
        }

        private void SinglePageProcessing(List<FinancialStatement> singlePageStatements)
        {
            ProgressBar.Invoke(() =>
            {
                ProgressBar.Style = ProgressBarStyle.Blocks;
            });
            int i = 1;
            Parallel.ForEach(singlePageStatements.ToList(),
                new ParallelOptions
                {
                    MaxDegreeOfParallelism = Model.SecondPage.CPUCores
                },
                async s =>
                {
                    WriteLog($"Scheduled '{s.FullName}'");
                    MoveProgress();

                    var model = new Core.Statement.PageResources
                    {
                        FinancialStatement = s,
                        Images = new Dictionary<string, string>()
                    };
                    model.Images.Add("logo", Convert.ToBase64String(File.ReadAllBytes("logo.jpg")));
                    model.Images.Add("qrCode", XMLProcessor.GenerateQRCode(model.FinancialStatement.QRContent));

                    var html = await RazorTemplateEngine.RenderAsync("/Theme.cshtml", model);
                    PDFEngine.Generate(html, $"outputs/{i++}.pdf");

                    WriteLog($"Completed '{s.FullName}'");
                    WriteProgressText($"Working on '{s.FullName}'");
                    MoveProgress();
                });
        }

        private void MoveProgress()
        {
            ProgressBar.Invoke(() =>
            {
                ProgressBar.Value++;
            });
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

        private void CPUCores_SelectedIndexChanged(object sender, EventArgs e)
        {
            Model.SecondPage.CPUCores = int.Parse(CPUCores.SelectedItem.ToString());
        }
    }
}