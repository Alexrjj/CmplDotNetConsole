using System;
using System.IO;
using System.Windows.Forms;
using OpenQA.Selenium;

namespace CmplConsole
{
    public class SobStatus
    {
        public static void ExtrairSobsStatus()
        {
            Gomnet.GomnetSettings();
            ChromeSettings.ChromeInitializer();
            Gomnet.LogaGomnet();

            ChromeSettings.driver.Navigate().GoToUrl(Gomnet.urlConsulta);

            string line;
            StreamReader file = new StreamReader(Gomnet.sobs);
            while ((line = file.ReadLine()) != null)
            {
                ChromeSettings.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB")).Clear();
                var sob = ChromeSettings.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB"));
                sob.SendKeys(line);
                ChromeSettings.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

                try
                {
                    var numSob = ChromeSettings.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8][contains(text(), '" + line + "' )]"));
                    if (numSob.Displayed)
                    {
                        var numSobArquivo = ChromeSettings.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8]")).Text;
                        var numStatusArquivo = ChromeSettings.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[3]")).Text;
                        
                        using (StreamWriter sw = File.AppendText(Gomnet.logSobsStatus))
                        {
                            sw.WriteLine(numSobArquivo + " " + numStatusArquivo);
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
            ChromeSettings.driver.Close();
        }
    }
}