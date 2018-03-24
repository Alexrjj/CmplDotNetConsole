using System;
using System.IO;
using System.Linq;
using Excel = NetOffice.ExcelApi;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace CmplConsole
{
    class ExtraiSobStatus
    {
        static void Main (string[] args)
        {
            //Variáveis
            string url = "http://gomnet.ampla.com/";
            string urlConsulta = "http://gomnet.ampla.com/ConsultaObra.aspx";
            string folder = Path.GetDirectoryName(Application.ExecutablePath);
            string sobs = folder + @"\sobs.txt";
            string log = folder + @"\Sobs&Status.txt";

            // ----  Chrome Settings ---- //
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless");
            //options.AddArgument("--window-size= 1600x900");
            options.AddArgument("start-maximized");
            options.LeaveBrowserRunning = true;
            // ----  Chrome Settings ---- //

            //Informações de credenciais para login no GOMNET
            string credenciais = @"C:\gomnet.xlsx";
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook workbook = excel.Workbooks.Open(credenciais);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            string login = worksheet.Range("A1").Value.ToString();
            string senha = worksheet.Range("A2").Value.ToString();

            // Fecha o Excel e o processo .exe
            excel.Quit();
            excel.Dispose();

            //Inicia o webdriver do Chrome
            IWebDriver driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl(url);

            //Loga no sistema
            IWebElement usrname = driver.FindElement(By.Id("txtBoxLogin"));
            IWebElement usrpass = driver.FindElement(By.Id("txtBoxSenha"));
            usrname.SendKeys(login);
            usrpass.SendKeys(senha);
            driver.FindElement(By.Id("ImageButton_Login")).Click();

            driver.Navigate().GoToUrl(urlConsulta);
            //var sob = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB"));

            string line;
            StreamReader file = new StreamReader(sobs);
            while ((line = file.ReadLine()) != null)
            {
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB")).Clear();
                var sob = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB"));
                sob.SendKeys(line);
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

                try
                {
                    var numSob = driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8][contains(text(), '" + line + "' )]"));
                    if (numSob.Displayed)
                    {
                        var numSobArquivo = driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8]")).Text;
                        var numStatusArquivo = driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[3]")).Text;
                        //if (!File.Exists(log))
                        //{
                            using (StreamWriter sw = File.AppendText(log))
                            {
                                sw.WriteLine(numSobArquivo + " " + numStatusArquivo);
                            }
                        //}
                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine(line + " not found.");
                    //file.Close();
                    continue;
                }
            }
            // Fecha o chromedriver
            driver.Close();
        }
    }
}
