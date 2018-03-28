using System;
using System.IO;
using System.Windows.Forms;
using OpenQA.Selenium;

namespace CmplConsole
{
    class SobTrabalho
    {
        public static void ExtrairSobTrabalho()
        {
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            Chrome.driver.Navigate().GoToUrl(Gomnet.urlConsulta);

            string line;
            StreamReader file = new StreamReader(Gomnet.sobs);
            while ((line = file.ReadLine()) != null)
            {
                Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB")).Clear();
                var sob = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB"));
                sob.SendKeys(line);
                Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

                try
                {
                    var numSob = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8][contains(text(), '" + line + "' )]"));
                    if (numSob.Displayed)
                    {
                        var numSobArquivo = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8]")).Text;
                        var numTrabArquivo = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[4]")).Text;

                        using (StreamWriter sw = File.AppendText(Gomnet.logSobTrab))
                        {
                            sw.WriteLine(numSobArquivo + " " + numTrabArquivo);
                        }

                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine(line + " not found.");
                    continue;
                }
            }
            // Fecha o chromedriver
            Chrome.driver.Close();
        }
    }
}
