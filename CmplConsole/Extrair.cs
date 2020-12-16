using System;
using System.IO;
using OpenQA.Selenium;
using OfficeOpenXml;
using OpenQA.Selenium.Support.UI;
using System.Linq;
using System.Text.RegularExpressions;
using Actions = OpenQA.Selenium.Interactions.Actions;

namespace CmplConsole
{
    class Extrair
    {
        public static void SobTrabalho()
        {
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            Chrome.driver.Navigate().GoToUrl(Gomnet.urlConsulta);

            string line;
            StreamReader file = new StreamReader(Gomnet.InSobTrab);
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

                        using (StreamWriter sw = File.AppendText(Gomnet.OutSobTrab))
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

        public static void SobStatus()
        {
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            Chrome.driver.Navigate().GoToUrl(Gomnet.urlConsulta);

            var pastaTrabalho = Gomnet.pasta.Worksheets.First();
            var linhas = pastaTrabalho.Dimension.End.Row;

            for (int linha = 2; linha <= linhas; linha++)
            {
                string x = pastaTrabalho.Cells[linha, 3].Value.ToString();
                if (x != null)
                {
                    string doisZeros = Regex.Replace(x, @"\A131|\A100", "00$&"); // Para sobs iniciadas em "100" ou "131", acrescenta dois zeros no início da mesma.
                    string cincoZeros = Regex.Replace(doisZeros, @"\A[1-4]", "00000$&"); // Para sobs iniciadas em 1-4, acrescenta cinco zeros no início da mesma.
                    string sobFinal = Regex.Split(cincoZeros, @"[\s\.]")[0]; // Delimita espaço quando há duas sobs na mesma célula.

                    try
                    {
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB")).Clear();
                        var sob = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB"));
                        sob.SendKeys(sobFinal);
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

                        try
                        {
                            var numSob = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8][contains(text(), '" + sobFinal + "' )]"));
                            if (numSob.Displayed)
                            {
                                var numSobArquivo = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8]")).Text;
                                var numStatusArquivo = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[3]")).Text;

                                //using (StreamWriter sw = File.AppendText(Gomnet.OutSobStatus))
                                //{
                                //    sw.WriteLine(numSobArquivo + " " + numStatusArquivo);
                                //}
                                Console.WriteLine(numSobArquivo + numStatusArquivo);

                            }
                        }
                        catch (NoSuchElementException)
                        {
                            Console.WriteLine(sobFinal + " não encontrada.");
                            continue;
                        }
                    }
                    catch (NullReferenceException)
                    {
                        Console.WriteLine("Não há mais sob para verificar...");
                        continue;
                    }
                }
            }
            // Fecha o chromedriver
            Chrome.driver.Close();
        }

            //    string line;
            //    StreamReader file = new StreamReader(Gomnet.InSobStatus);
            //    while ((line = file.ReadLine()) != null)
            //    {
            //        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB")).Clear();
            //        var sob = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox_NumSOB"));
            //        sob.SendKeys(line);
            //        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

            //        try
            //        {
            //            var numSob = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8][contains(text(), '" + line + "' )]"));
            //            if (numSob.Displayed)
            //            {
            //                var numSobArquivo = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[8]")).Text;
            //                var numStatusArquivo = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[3]")).Text;

            //                using (StreamWriter sw = File.AppendText(Gomnet.OutSobStatus))
            //                {
            //                    sw.WriteLine(numSobArquivo + " " + numStatusArquivo);
            //                }

            //            }
            //        }
            //        catch (NoSuchElementException)
            //        {
            //            Console.WriteLine(line + " not found.");
            //            continue;
            //        }
            //    }
            //    // Fecha o chromedriver
            //    Chrome.driver.Close();
            //}

            public static void SobEnergizacao()
        {
            //TODO: Implementar verificação de sob energizada.
        }
    }
}
