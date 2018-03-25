using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = NetOffice.ExcelApi;

namespace CmplConsole
{
    public class Gomnet
    {
        public static string url;
        public static string urlConsulta;
        public static string folder;
        public static string sobs;
        public static string logSobsStatus;
        public static Excel.Application excel = new Excel.Application();
        public static string login;
        public static string senha;

        public static void Settings()
        {
            url = "http://gomnet.ampla.com/";
            urlConsulta = "http://gomnet.ampla.com/ConsultaObra.aspx";
            folder = Path.GetDirectoryName(Application.ExecutablePath);
            sobs = folder + @"\sobs.txt";
            logSobsStatus = folder + @"\Sobs&Status.txt";
        }

        public static void Login()
        {
            //Informações de credenciais para login no GOMNET
            string credenciais = @"C:\gomnet.xlsx";
            excel.DisplayAlerts = false;
            Excel.Workbook workbook = excel.Workbooks.Open(credenciais);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            login = worksheet.Range("A1").Value.ToString();
            senha = worksheet.Range("A2").Value.ToString();

            // Fecha o Excel e o processo .exe
            excel.Quit();
            excel.Dispose();

            //Inicia o webdriver do Chrome
            try
            {
                Chrome.driver.Navigate().GoToUrl(Gomnet.url);
            }
            catch (System.NullReferenceException) // Para o script, caso falhe a opção escolhida no ChromeInitializer
            {
                return;
            }

            //Loga no sistema
            IWebElement usrname = Chrome.driver.FindElement(By.Id("txtBoxLogin"));
            IWebElement usrpass = Chrome.driver.FindElement(By.Id("txtBoxSenha"));
            usrname.SendKeys(login);
            usrpass.SendKeys(senha);
            Chrome.driver.FindElement(By.Id("ImageButton_Login")).Click();
        }
    }
}
