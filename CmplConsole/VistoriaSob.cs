using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace CmplConsole
{
    class VistoriaSob
    {
        public static void VincSupervisor()
        {
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();
            Chrome.driver.Navigate().GoToUrl(Gomnet.urlVincSup);

            string line;
            StreamReader file = new StreamReader(Gomnet.InSobSup);
            while ((line = file.ReadLine()) != null)
            {
                Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_SOB")).Clear();
                var sob = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_SOB"));
                sob.SendKeys(line);
                Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Pesquisar")).Click();

                try
                {
                    var supervisor = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_gridViewTarefas_ctl02_ddlSupervisor"));
                    var opcao = new SelectElement(supervisor);
                    opcao.SelectByText("MESSIAS JOSE DE FARIA");
                    Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Button_VincularSupervisor")).Click();
                    Chrome.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                    try
                    {
                        Chrome.driver.SwitchTo().Alert().Accept();
                        Console.WriteLine(line + " linked");
                    }
                    catch (NoAlertPresentException)
                    {
                        continue;
                    }
                } catch (NoSuchElementException)
                {
                    Console.WriteLine(line + " not linked.");
                    continue;
                }
            }

        }

        public static void Vistoria()
        {
            // TODO: Implementar vistoria de sob
            // Obs.: Verificar se a sob já foi vistoriada e informar no console
        }
    }
}
