using Aspose.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeOpenXml;
using Aspose.Pdf.Operators;


namespace PdftoExcel
{
    class UserFunctions
    {

        public static void WriteLog(string secureId, string request, string response, string servicename, string mfunctionName, [CallerMemberName] string callerName = "")
        {
            mfunctionName = callerName;
            string logFilePath = "C:\\Logs\\" + servicename + "\\";
            logFilePath = logFilePath + "Log-" + DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
            try
            {
                using (FileStream fileStream = new FileStream(logFilePath, FileMode.Append))
                {
                    FileInfo logFileInfo;


                    logFileInfo = new FileInfo(logFilePath);
                    DirectoryInfo logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
                    if (!logDirInfo.Exists) logDirInfo.Create();

                    DirectorySecurity dSecurity = logDirInfo.GetAccessControl();
                    dSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
                    logDirInfo.SetAccessControl(dSecurity);

                    StreamWriter log = new StreamWriter(fileStream);


                    if (!logFileInfo.Exists)
                    {
                        _ = logFileInfo.Create();
                    }
                    else
                    {
                        log.WriteLine(secureId);
                        log.WriteLine(DateTime.UtcNow.ToString());
                        log.WriteLine(request);
                        log.WriteLine(response);
                        log.WriteLine(mfunctionName);
                        log.WriteLine("_____________________________________________________________________________________");
                        log.Close();
                    }
                    fileStream.Close();
                }
            }
            catch (Exception)
            {

            }
        }

        public static bool ReadAllFiles(string sourcePath, out List<string> filePath, string extension)
        {
            bool worked = false;
            filePath = new List<string>();

            try
            {
                int counter = 0;

                foreach (string file in Directory.EnumerateFiles(sourcePath, "*." + extension))
                {
                    filePath.Add(file);
                    counter++;
                }

                if (counter > 0)
                {
                    filePath.RemoveAll(item => string.IsNullOrEmpty(item));
                    worked = true;
                }
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", sourcePath, ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }
            return worked;
        }

        public static void ConvertPDFtoEXCEL(string input, string output)
        {
            try
            {
                // load PDF with an instance of Document                        
                var document = new Document(input + "." + "pdf");

                // save document in Excel format
                document.Save(output + DateTime.Now.ToString("dd-MM-yyyy hh mm ss") + "." + "xlsx", Aspose.Pdf.SaveFormat.Excel);

                Console.WriteLine("PDF content has been successfully converted to Excel.");
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", input, ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }


        }



        public static void SplitPdfIntoFours(string input, string output)
        {
            try {
                {
                    string inputPdf = input + "." + "pdf";
                    int interval = 4;

                    using (PdfDocument pdfDocument = new PdfDocument(new PdfReader(inputPdf)))
                    {
                        FileInfo file = new FileInfo(inputPdf);
                        string fileName = file.Name;
                        int totalPages = pdfDocument.GetNumberOfPages();
                        int numberOfFiles = (int)Math.Ceiling((double)totalPages / interval);

                        for (int i = 0; i < numberOfFiles; i++)
                        {
                            int startIndex = i ==0  ? 1 :i * interval;
                            
                            int endIndex = Math.Min(startIndex + interval - 1, totalPages - 1);

                            string outputPdf = output+fileName+$"_{i + 1}.pdf";
                            using (PdfDocument newPdfDocument = new PdfDocument(new PdfWriter(outputPdf)))
                            {
                                for (int j = startIndex; j <= endIndex; j++)
                                {
                                    PdfPage page = pdfDocument.GetPage(j);
                                    newPdfDocument.AddPage(page.CopyTo(newPdfDocument));
                                }
                            }
                        }
                    }
                }
                Console.WriteLine("PDF content has been successfully split and saved.");
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", input, ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                Console.WriteLine("===> "+ ex.Message);
            }
        }

        public static void cleanUp(string input, string output)
        {
            {
                string sourceDirectory = input; 
                string destinationDirectory = output; 

                // Ensure the source directory exists
                if (!Directory.Exists(sourceDirectory))
                {
                    Console.WriteLine("Source directory does not exist.");
                    return;
                }

                // Ensure the destination directory exists; if not, create it
                if (!Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                // Get a list of all files in the source directory
                string[] files = Directory.GetFiles(sourceDirectory);

                foreach (string file in files)
                {
                    try
                    {
                        // Get the file name and build the destination path
                        string fileName = Path.GetFileName(file);
                        string destinationPath = Path.Combine(destinationDirectory, fileName);

                        // Move the file from the source directory to the destination directory
                        File.Move(file, destinationPath);
                        Console.WriteLine($"Moved: {file} to {destinationPath}");
                    }
                    catch (IOException e)
                    {
                        Console.WriteLine($"Error moving file: {e.Message}");
                    }
                }

                Console.WriteLine("All files moved successfully.");
            }


        }

        public static void CopyInwardOutwardToTemplate(string input, string output, string fileName)
        {
            string sourceFilePath = input+".xlsx";
            string destinationFilePath = output;

            using (var sourcePackage = new ExcelPackage(new FileInfo(sourceFilePath)))
            using (var destinationPackage = new ExcelPackage())
            {
                foreach (var sourceWorksheet in sourcePackage.Workbook.Worksheets)
                {
                    string GIPVal = sourceWorksheet.Cells["A4"].Text.Trim();

                    var destinationWorksheet = destinationPackage.Workbook.Worksheets.Add(sourceWorksheet.Name);

                    if (GIPVal.Equals("GIP INWARDS", StringComparison.OrdinalIgnoreCase))
                    {                    
                        destinationWorksheet.Cells["A6:I31"].LoadFromText(sourceWorksheet.Cells["A1:K40"].Text);
                    }
                    else if (GIPVal.Equals("GIP OUTWARDS", StringComparison.OrdinalIgnoreCase))
                    {
                        var secondDestinationWorksheet = destinationPackage.Workbook.Worksheets.Add(sourceWorksheet.Name + " OUTWARD");
                        secondDestinationWorksheet.Cells["A6:I31"].LoadFromText(sourceWorksheet.Cells["A1:K40"].Text);
                    }
                    else
                    {                   
                        Console.WriteLine($"Unsupported GIPVal '{GIPVal}' in worksheet '{sourceWorksheet.Name}'");
                    }
                }

                destinationPackage.SaveAs(new FileInfo(destinationFilePath+fileName+".xlsx"));
            }

            Console.WriteLine("Data copied and saved successfully.");
        }


    }
}
