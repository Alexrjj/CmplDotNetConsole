using System;
using System.IO;
using OpenQA.Selenium;

namespace CmplConsole
{
    class Download
    {
        public static void SGD() //TODO: Exibir um bloco de texto para inserção das sobs, além de uma caixa com escolha de data para baixar DWG (mais atual) + SGD (da data específica)
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

        public static void Orcamento()
        {
            //TODO: Implementar download de orçamento da sob
        }

        public static void DWG()
        {
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            string line;
            StreamReader file = new StreamReader(Gomnet.InSobDWG);
            while ((line = file.ReadLine()) != null)
            {
                Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + line);
                try //Busca pela versão mais atual do DWG da Sob
                {
                    try
                    {
                        var rev_09 = Chrome.driver.FindElement(By.PartialLinkText("REV.09.dwg"));
                        if (rev_09.Displayed)
                        {
                            rev_09.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_08 = Chrome.driver.FindElement(By.PartialLinkText("REV.08.dwg"));
                        if (rev_08.Displayed)
                        {
                            rev_08.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_07 = Chrome.driver.FindElement(By.PartialLinkText("REV.07.dwg"));
                        if (rev_07.Displayed)
                        {
                            rev_07.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_06 = Chrome.driver.FindElement(By.PartialLinkText("REV.06.dwg"));
                        if (rev_06.Displayed)
                        {
                            rev_06.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_05 = Chrome.driver.FindElement(By.PartialLinkText("REV.05.dwg"));
                        if (rev_05.Displayed)
                        {
                            rev_05.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_04 = Chrome.driver.FindElement(By.PartialLinkText("REV.04.dwg"));
                        if (rev_04.Displayed)
                        {
                            rev_04.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_03 = Chrome.driver.FindElement(By.PartialLinkText("REV.03.dwg"));
                        if (rev_03.Displayed)
                        {
                            rev_03.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_02 = Chrome.driver.FindElement(By.PartialLinkText("REV.02.dwg"));
                        if (rev_02.Displayed)
                        {
                            rev_02.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev_01 = Chrome.driver.FindElement(By.PartialLinkText("REV.01.dwg"));
                        if (rev_01.Displayed)
                        {
                            rev_01.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra9 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_09.dwg"));
                        if (revObra9.Displayed)
                        {
                            revObra9.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra8 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_08.dwg"));
                        if (revObra8.Displayed)
                        {
                            revObra8.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra7 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_07.dwg"));
                        if (revObra7.Displayed)
                        {
                            revObra7.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra6 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_06.dwg"));
                        if (revObra6.Displayed)
                        {
                            revObra6.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra5 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_05.dwg"));
                        if (revObra5.Displayed)
                        {
                            revObra5.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra4 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_04.dwg"));
                        if (revObra4.Displayed)
                        {
                            revObra4.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra3 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_03.dwg"));
                        if (revObra3.Displayed)
                        {
                            revObra3.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra2 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_02.dwg"));
                        if (revObra2.Displayed)
                        {
                            revObra2.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra1 = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA_01.dwg"));
                        if (revObra1.Displayed)
                        {
                            revObra1.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var revObra = Chrome.driver.FindElement(By.PartialLinkText("REVOBRA.dwg"));
                        if (revObra.Displayed)
                        {
                            revObra.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev9 = Chrome.driver.FindElement(By.PartialLinkText("REV09.dwg"));
                        if (rev9.Displayed)
                        {
                            rev9.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev8 = Chrome.driver.FindElement(By.PartialLinkText("REV08.dwg"));
                        if (rev8.Displayed)
                        {
                            rev8.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev7 = Chrome.driver.FindElement(By.PartialLinkText("REV07.dwg"));
                        if (rev7.Displayed)
                        {
                            rev7.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev6 = Chrome.driver.FindElement(By.PartialLinkText("REV06.dwg"));
                        if (rev6.Displayed)
                        {
                            rev6.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev5 = Chrome.driver.FindElement(By.PartialLinkText("REV05.dwg"));
                        if (rev5.Displayed)
                        {
                            rev5.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev4 = Chrome.driver.FindElement(By.PartialLinkText("REV04.dwg"));
                        if (rev4.Displayed)
                        {
                            rev4.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev3 = Chrome.driver.FindElement(By.PartialLinkText("REV03.dwg"));
                        if (rev3.Displayed)
                        {
                            rev3.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev2 = Chrome.driver.FindElement(By.PartialLinkText("REV02.dwg"));
                        if (rev2.Displayed)
                        {
                            rev2.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var rev1 = Chrome.driver.FindElement(By.PartialLinkText("REV01.dwg"));
                        if (rev1.Displayed)
                        {
                            rev1.Click();
                            Gomnet.FechaJanela();
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        var dwg = Chrome.driver.FindElement(By.PartialLinkText("dwg"));
                        if (dwg.Displayed)
                        {
                            dwg.Click();
                            Gomnet.FechaJanela();
                            continue;
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
            Console.WriteLine("End of execution.");
        }
    }
}
