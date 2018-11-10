using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Reflection;
using OpenQA.Selenium.Support.UI;
using Moneta;
using System.Threading;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Expedia
{
    class Program
    {


        private static string build_number;
        private static string build_username;
        private static string environment;
        public static string ExcelFileName;
        private static string ExcelWorkSheet;
        private static string Browser;
        private static string baseURL;
        private static string destinationPath;
        private static IWebDriver driver;

       
        static MonetaHelper helpObj = new MonetaHelper();
        



        public static void SettingUp()
        {

            build_number = TpsWebDriver.initializeVariable("build_number");
            build_username = TpsWebDriver.initializeVariable("build_username");
            environment = TpsWebDriver.initializeVariable("environment");
            ExcelFileName = TpsWebDriver.initializeVariable("ExcelFileName");
            ExcelWorkSheet = TpsWebDriver.initializeVariable("ExcelWorkSheet");
            Browser = TpsWebDriver.initializeVariable("Browser");
            baseURL = TpsWebDriver.initializeVariable("baseURL");
            destinationPath = TpsWebDriver.initializeVariable("destinationPath");

        }


        public static void ExcelReading()
        {

            try
            {
                MonetaHelper MonHelp = new MonetaHelper();

                ImportModule();

                Thread.Sleep(5000);


                DataTable dt1 = ExcelLib.PopulateInCollection(ExcelFileName, ExcelWorkSheet);

                for (int i = 1; i <= dt1.Rows.Count; i++)
                {

                    driver = MonHelp.SendKeys(driver, By.XPath("//div/input[contains(@name,'search')]"), ExcelLib.ReadData(i, "People Name"));
                    TpsWebDriver.PageLoad(driver);

                    driver = MonHelp.ButtonClick(driver, By.XPath("//button[contains(@type,'submit')]"));
                    TpsWebDriver.PageLoad(driver);


                    try
                    {


                        IWebElement element = driver.FindElement(By.XPath("//p/i/a[contains(@title,'page does not exist')]"));
                        string elementHtml = element.Text;
                        Console.WriteLine("UI Message:" + elementHtml);


                   /*     IWebElement elementcross = driver.FindElement(By.XPath("//*[@id='mw - content - text']/div[4]/ul/li[1]/div[1]/a[contains(@title,'"+ExcelLib.ReadData(i, "People Name")+"')]"));
                        string elementHtmlnew = elementcross.Text;
                        string lowerelement = elementHtmlnew.ToLower();
                        string excellower = ExcelLib.ReadData(i, "People Name").ToLower();


                        if (lowerelement.Contains(excellower))
                        {

                            UpdateExcel("Sheet1", i + 1, 4, "Yes");
                            Console.WriteLine("Done");



                        }

                     */
                     
                        
                        //*[@id="mw-content-text"]/div[4]/ul/li[1]/div[1]/a[contains(@title,'page does not exist')]

                        //  driver.findElement(By.xpath("//span[@class='title']"));


                        if (elementHtml.Contains(ExcelLib.ReadData(i, "People Name")))
                        {

                            UpdateExcel("Sheet1", i + 1, 4, "No");
                            Console.WriteLine("Done");



                        }

                        else
                        {


                        }


                    }





                    catch (Exception ex)
                    {

                        try
                        {

                            string test = ExcelLib.ReadData(i, "People Name");
                            string lowertest = test.ToLower();
                            


                            string title = driver.FindElement(By.XPath("//*[@id='mw-content-text']/div[3]/ul/li[1]/div[1]/a")).Text;
                            string lowertitle = title.ToLower();
                            Console.WriteLine(lowertitle);

                            //  string substringtitle = lowertitle.Substring(0, 5);


                            // string element1 = driver.FindElement(By.XPath("//span[text()='" + words[0] + "']/parent::a[@title='" + ExcelLib.ReadData(i, "People Name") + "']")).Text;

                            //span[contains(text(),'" + ExcelLib.ReadData(i, "ExpectedMessage") + "')]")).Text
                            // string elementHtmlnew = element1.Text;
                            //span[text()='"+words[0]+"']/parent::a[@title='" + ExcelLib.ReadData(i, "People Name") + "']

                            //   Console.WriteLine(element1);

                            /*  if (elementHtml.Contains(ExcelLib.ReadData(i, "People Name")))
                              {

                                  UpdateExcel("Sheet1", i + 1, 4, "No");
                                  Console.WriteLine("Done");



                              }
                              */
                            if (lowertitle.Equals(lowertest))
                            {

                                UpdateExcel("Sheet1", i + 1, 4, "Yes");
                                Console.WriteLine("Done");


                            }

                            else
                            {

                             UpdateExcel("Sheet1", i + 1, 4, "Review");
                             Console.WriteLine("Done");
                            }

                        }

                        catch (Exception ex1)
                        {

                            

                            UpdateExcel("Sheet1", i + 1, 4, "Review");
                            Console.WriteLine("Done");

                        }



                       
                    }


                  
                        //a[@title='Peter Sullivan']/span[text()='Peter']




                     /*   if (elementHtml.Contains(ExcelLib.ReadData(i, "People Name")))
                            {

                                UpdateExcel("Sheet1", i + 1, 4, "No");
                                Console.WriteLine("Done");
                               


                            }

                            else if (element1.Equals(ExcelLib.ReadData(i, "People Name")))
                            {

                                UpdateExcel("Sheet1", i + 1, 4, "Review");
                                Console.WriteLine("Done");


                            }

                            else
                            {



                                UpdateExcel("Sheet1", i + 1, 4, "Yes");
                                Console.WriteLine("Done");
                            }

                        */
                       


                    driver = MonHelp.SendKeys(driver, By.XPath("//div/input[contains(@name,'search')]"), ExcelLib.ReadData(i, "Company Name"));
                    TpsWebDriver.PageLoad(driver);


                    driver = MonHelp.ButtonClick(driver, By.XPath("//button[contains(@type,'submit')]"));
                    TpsWebDriver.PageLoad(driver);

                    try
                    {


                        IWebElement element = driver.FindElement(By.XPath("//p/i/a[contains(@title,'page does not exist')]"));
                        string elementHtml = element.Text;
                        Console.WriteLine("UI Message:" + elementHtml);




                        if (elementHtml.Contains(ExcelLib.ReadData(i, "People Name")))
                        {

                            UpdateExcel("Sheet1", i + 1, 5, "No");
                            Console.WriteLine("Done");



                        }

                        else
                        {


                        }


                    }



                    catch (Exception ex)
                    {

                        try
                        {

                            string test = ExcelLib.ReadData(i, "People Name");
                            string lowertest = test.ToLower();



                            string title = driver.FindElement(By.XPath("//*[@id='mw-content-text']/div[3]/ul/li[1]/div[1]/a")).Text;
                            string lowertitle = title.ToLower();
                            Console.WriteLine(lowertitle);

                           
                            if (lowertitle.Equals(lowertest))
                            {

                                UpdateExcel("Sheet1", i + 1, 5, "Yes");
                                Console.WriteLine("Done");


                            }

                            else
                            {

                                UpdateExcel("Sheet1", i + 1, 5, "Review");
                                Console.WriteLine("Done");
                            }




                        }

                        catch (Exception ex1)
                        {

                            UpdateExcel("Sheet1", i + 1, 5, "Review");
                            Console.WriteLine("Done");

                        }




                    }





                }


            }



            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);

            }



        }



        public static void ImportModule()
        {

            driver = helpObj.setBrowser(Browser);

            driver.Navigate().GoToUrl(baseURL + "/");
            Console.WriteLine("Go to URL : " + baseURL);
            TpsWebDriver.PageLoad(driver);

            driver.Manage().Window.Maximize();
            Console.WriteLine("Window Maximized");
            TpsWebDriver.PageLoad(driver);




        }


        






        public static void UpdateExcel(string sheetName, int row, int col, string data)
        {
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                string path = Directory.GetCurrentDirectory();
                string target = path + "\\" + destinationPath;


             

                oWB = oXL.Workbooks.Open(target);
                oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                oSheet.Cells[row, col] = data;

                oWB.Save();


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                    oWB.Close();
            }

        }

        static public void TeardownTest()
        {

  
            driver.Close();

        }

        static void Main(string[] args)
        {

            SettingUp();
            ExcelReading();
            TeardownTest();


        }

    }
}
