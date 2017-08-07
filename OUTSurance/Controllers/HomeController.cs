using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.IO.IsolatedStorage;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using OUTSurance.Models;


using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel; 

namespace OUTSurance.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/
        /// <summary>
        /// Method to read Data.csv file
        /// </summary>
        /// <returns>Reads Data.Csv File and Writes to Data1.txt and Data2.txt files and Display the values on to the View.</returns>
        public ActionResult Index()
        {
            return View(new List<ExcelDataModel>());
        }

 
        /// <summary>
        /// Method to read Data.csv file
        /// </summary>
        /// <param name="postedFile"></param>
        /// <returns>Reads Data.Csv File and Writes to Data1.txt and Data2.txt files</returns>
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {

            // Create Instance/Instantiate of ExcelDataModel Model/Class 
            List<ExcelDataModel> excelDataModel = new List<ExcelDataModel>();

            string filePath = string.Empty;
            if (postedFile != null)
            {

                // Excel file Path Location
                string path = Server.MapPath("~/Uploads/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                // Excel Doc  file path
                filePath = path + Path.GetFileName(postedFile.FileName);
                string extension = Path.GetExtension(postedFile.FileName);
               
                postedFile.SaveAs(filePath);

                //Read the contents of CSV file.
                string csvData = System.IO.File.ReadAllText(filePath);

                // Below code is to write Data.Csv file Query to the Text File2 on to the location (C:\\Data1). 
                
                try
                {         
                    // Execute a loop over the rows.
                    foreach (string row in csvData.Split('\n'))
                    {
                        if (!string.IsNullOrEmpty(row))
                        {
                            excelDataModel.Add(new ExcelDataModel
                            {
                                FirstName = row.Split(',')[0],
                                 LastName = row.Split(',')[1],
                                  Address = row.Split(',')[2],
                              PhoneNumber = row.Split(',')[3]
                            });
                        }
                    }

                    // Write the string to a file.
                    System.IO.StreamWriter file = new System.IO.StreamWriter("c:\\Data1.txt", true);
                    file.WriteLine(csvData);
                    file.Close();
                }
                
                // Catch exception
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }


                // Below code is to write Addrress Details to the Text File2 on to the location (C:\\Data2). 
                try
                {
                        
                    // Execute a loop over the rows.
                    foreach (string row in csvData.Split('\n'))
                    {
                        if (!string.IsNullOrEmpty(row))
                        {
                            excelDataModel.Add(new ExcelDataModel
                            {                   
                                Address = row.Split(',')[2],                     
                            });
                        }
                    }

                    // Write the string to a file.
                    System.IO.StreamWriter file = new System.IO.StreamWriter("c:\\Data2.txt", true);
                    file.WriteLine(csvData);
                    file.Close();
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                }            

               
            }
            
            return View(excelDataModel);  // returns the View
        }



    } // End of HomeController class
} // End of Namespace
