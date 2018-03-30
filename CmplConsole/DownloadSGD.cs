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

            //string month = DateTime.Now.ToString("MM");
            Console.Write("Month Value: ");
            string month = Console.ReadLine();
            string line;
            StreamReader file = new StreamReader(Gomnet.InSobSGD);
            while ((line = file.ReadLine()) != null)
            {
                Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + line);
                try
                {
                    try
                    {
                        var pdfs = Chrome.driver.FindElements(By.PartialLinkText("-" + month + "-18")).Count;
                        for (int pdf = 0; pdf < pdfs; pdf++)
                        {
                            Chrome.driver.FindElements(By.PartialLinkText("-" + month + "-18"))[pdf].Click();
                            Gomnet.FechaJanela();
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                }
                catch (NoSuchElementException)
                {
                }
            }
            Console.Write("End of execution.");
        }
    }
}
