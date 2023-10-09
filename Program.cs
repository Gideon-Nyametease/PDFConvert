
using System;
using System.Configuration;
using System.Reflection;
using Aspose.Pdf;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;


namespace PdftoExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {

			try
			{
                Console.WriteLine("----------------Start--------------------");

                Console.WriteLine("-----------------------------------------");

                Console.WriteLine("Start Time ------------->   " + DateTime.Now);
                string source = ConfigurationManager.AppSettings["inputPath"];
                string source2 = ConfigurationManager.AppSettings["splitPdfPath"];
                string splitExcelPath = ConfigurationManager.AppSettings["splitExcelPath"];
                string pdfBackup = ConfigurationManager.AppSettings["pdfBackup"];
                string finalDestination = ConfigurationManager.AppSettings["destination"];
                string output = ConfigurationManager.AppSettings["outputPath"];

                // Get files from source and split them
                if (!UserFunctions.ReadAllFiles(source, out List<string> filePath, "pdf"))
                {
                    Console.WriteLine("No data found in source path");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "No data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                }

                else
                {
                    Console.WriteLine("Data found in specified location");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog("Before SPLIT", "", "Data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                }

                foreach(var item in filePath)
                {
                    string fileName = Path.GetFileNameWithoutExtension(item);
                    UserFunctions.SplitPdfIntoFours(source + fileName, source2);
                    UserFunctions.cleanUp(source, pdfBackup);
                    //System.Threading.Thread.Sleep(1000);
                    continue;
                }


                // Process split files using Aspose
                if (!UserFunctions.ReadAllFiles(source2, out List<string> splitFilePath, "pdf"))
                {
                    Console.WriteLine("No data found in source path");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "No data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                }

                else
                {
                    Console.WriteLine("Data found in specified location");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog("After SPLIT", "", "Data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                }
                foreach (var item in splitFilePath)
                {
                    string fileName = Path.GetFileNameWithoutExtension(item);
                    UserFunctions.ConvertPDFtoEXCEL(source2 + fileName, splitExcelPath);
                    System.Threading.Thread.Sleep(800);
                    continue;
                }



                // Copy data from split files into new template
               // if (!UserFunctions.ReadAllFiles(splitExcelPath, out List<string> splitExcels, "xlsx"))
                //{
                 //   Console.WriteLine("No data found in source path");

                   // Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "No data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                //}

                //else
                //{
                  //  Console.WriteLine("Data found in specified location");

                    //Task.Factory.StartNew(() => UserFunctions.WriteLog("After EXCEL SPLIT", "", "Data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                //}
                //foreach (var item in splitExcels)
                //{
                  //  string fileName = Path.GetFileNameWithoutExtension(item);
                    //UserFunctions.CopyInwardOutwardToTemplate(splitExcelPath + fileName, finalDestination, fileName);
                    //System.Threading.Thread.Sleep(1800);
                    //continue;
                //}




            }
			catch (Exception ex)
			{
                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, "Error", string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                Console.WriteLine(" ");
                Console.WriteLine("Exception -------------------->    " + ex.Message + "  || " + ex.StackTrace);
            }

            Console.WriteLine("");
            Console.WriteLine("Process Completed @ " + DateTime.Now);

            System.Threading.Thread.Sleep(5000);
        }
    }
}