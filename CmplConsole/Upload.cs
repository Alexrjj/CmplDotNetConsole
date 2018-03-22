using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = NetOffice.ExcelApi;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace CmplConsole
{
    class Upload
    {
        static void Main(string[] args)
        {
            //URL's
            string url = "http://gomnet.ampla.com/";

            //Informações de credenciais para login no GOMNET
            string credenciais = @"C:\gomnet.xlsx";
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook workbook = excel.Workbooks.Open(credenciais);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            string login = worksheet.Range("A1").Value.ToString();
            string senha = worksheet.Range("A2").Value.ToString();
            
            // close excel and dispose reference
            excel.Quit();
            excel.Dispose();

            //Inicia o webdriver do Chrome
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(url);

            //Loga no sistema
            IWebElement usrname = driver.FindElement(By.Id("txtBoxLogin"));
            IWebElement usrpass = driver.FindElement(By.Id("txtBoxSenha"));
            usrname.SendKeys(login);
            usrpass.SendKeys(senha);
            driver.FindElement(By.Id("ImageButton_Login")).Click();

            driver.Close();
            Console.ReadKey();
        }
    }
}
