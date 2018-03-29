using System;
using System.IO;
using OpenQA.Selenium;

namespace CmplConsole
{
    class DownloadSGD
    {
        public static void SGD()
        {
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            string month = DateTime.Now.ToString("MM");
            string line;
            StreamReader file = new StreamReader(Gomnet.InSobDWG);
            while ((line = file.ReadLine()) != null)
            {
                Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + line);
                try
                {
                    try
                    {
                        var pdfs = Chrome.driver.FindElements(By.PartialLinkText("'-' " + month + " '-18'"));
                        foreach(int x in pdfs){
                            Chrome.driver.FindElements(By.PartialLinkText("'-' " + month + " '-18'"))[pdf].Click();
                        }
                    }
                }
            }
        }
    }
}

