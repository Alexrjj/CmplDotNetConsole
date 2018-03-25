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
            //Variáveis
            string url = "http://gomnet.ampla.com/";
            string urlConsulta = "http://gomnet.ampla.com/ConsultaObra.aspx";
            string folder = Path.GetDirectoryName(Application.ExecutablePath);
            string sobs = folder + @"\sobs.txt";
            string log = folder + @"\Sobs&Status.txt";

            ChromeSettings.ChromeInitializer();

            ExcelCredentials.LoginGomnet();
            // Fecha o Excel e o processo .exe
            ExcelCredentials.excel.Quit();
            ExcelCredentials.excel.Dispose();

            //Inicia o webdriver do Chrome
            try
            {
                ChromeSettings.driver.Navigate().GoToUrl(url);
            } catch (System.NullReferenceException) // Para o script, caso falhe a opção escolhida no ChromeInitializer
            {
                return;
            }
            
            //Loga no sistema
            IWebElement usrname = ChromeSettings.driver.FindElement(By.Id("txtBoxLogin"));
            IWebElement usrpass = ChromeSettings.driver.FindElement(By.Id("txtBoxSenha"));
            usrname.SendKeys(ExcelCredentials.login);
            usrpass.SendKeys(ExcelCredentials.senha);
            ChromeSettings.driver.FindElement(By.Id("ImageButton_Login")).Click();

            ChromeSettings.driver.Navigate().GoToUrl(urlConsulta);

            string line;
            StreamReader file = new StreamReader(sobs);
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
                        
                        using (StreamWriter sw = File.AppendText(log))
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