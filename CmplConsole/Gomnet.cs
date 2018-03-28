﻿using OpenQA.Selenium;
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
        public static Excel.Application excel;
        public static string login;
        public static string senha;
        public static string googleForm;
        public static string urlUpload;
        public static string logUpload;

        public static void Settings()
        {
            url = "http://gomnet.ampla.com/";
            urlConsulta = "http://gomnet.ampla.com/ConsultaObra.aspx";
            urlUpload = "http://gomnet.ampla.com/Upload.aspx?numsob=";
            googleForm = "https://docs.google.com/forms/d/e/1FAIpQLSdqi7NxRSzKM0M3-ZQn5Fpn6rKriKJnw_0EPnlD3iScx18yXg/viewform";
            folder = Path.GetDirectoryName(Application.ExecutablePath);
            sobs = folder + @"\sobs.txt";
            logSobsStatus = folder + @"\LogSobs&Status.txt";
            logUpload = folder + @"\LogUpload.txt";

        }

        public static void Login()
        {
            //Informações de credenciais para login no GOMNET
            string credenciais = @"C:\gomnet.xlsx";
            Excel.Application excel = new Excel.Application();
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
