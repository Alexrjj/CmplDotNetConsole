using System;
using System.IO;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace CmplConsole
{

    public class Upload
    {
        public static void UploadArquivos()
        {

            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            var files = Directory.GetFiles(Gomnet.folder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(s => s.EndsWith(".pdf") || s.EndsWith(".PDF"));
            foreach (string file in files)
            {
                string f = Path.GetFileNameWithoutExtension(file);
                string ext = Path.GetExtension(file);
                if (f.StartsWith("SG_QUAL_"))
                {
                    var s2 = file.Split('_');
                    var s3 = String.Join("_", s2.Skip(1).Take(3));
                    int lastIndex = s3.LastIndexOf('_');
                    if (Char.IsDigit(s3[lastIndex + 1]))
                    {
                        string result = s3.Remove(lastIndex, "_".Length).Insert(lastIndex, "-");
                        Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + result);
                    }
                    else
                    {
                        Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + String.Join("_", s2.Skip(1).Take(2)));
                    }
                }
                else if (f.StartsWith("SG_REF") || f.StartsWith("SG_QUAL") || f.StartsWith("SG_INTER") || f.StartsWith("SG_RNT"))
                {
                    var s2 = file.Split('_');
                    Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_PQ"))
                {
                    var s2 = file.Split('_');
                    Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + String.Join("_", s2.Skip(1).Take(3)));
                }
                else if (f.StartsWith("SG_RRC_"))
                {
                    var s2 = file.Split('_');
                    Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_RRC"))
                {
                    var s2 = file.Split('_');
                    Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + String.Join("_", s2.Skip(1).Take(2)).Replace("_", "-"));
                }
                else if (f.StartsWith("SG_"))
                {
                    var s2 = file.Split('_');
                    Chrome.driver.Navigate().GoToUrl(Gomnet.urlUpload + String.Join("_", s2.Skip(1).Take(1)));
                }

                try //Verifica se a sob foi digitada corretamente
                {
                    var erro = Chrome.driver.FindElement(By.XPath("*//tr/td[contains(text(), 'Não existem dados para serem exibidos.')]"));
                    if (erro.Displayed)
                    {
                        Console.WriteLine("{0} not found.", f);
                    }
                }
                catch (NoSuchElementException)
                {
                    try //Verifica se o arquivo já foi anexado
                    {
                        var anexo = Chrome.driver.FindElement(By.XPath("*//a[contains(text(), '" + f + "')]"));
                        if (anexo.Displayed)
                        {
                            Console.WriteLine("{0} already attached. ", f);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        if (f.Contains("FORM_FISC"))
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("ENCERRAMENTO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("FORMULARIO DE FISCALIZACAO DE OBRA");
                        }
                        else if (f.Contains("AS_BUILT"))
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("ENCERRAMENTO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("AS BUILT");
                        }
                        else if (f.Contains("APOIO_TRANSITO"))
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("PROJETO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("CARTAS/OFICIOS");
                        }
                        else if (f.Contains("CCR"))
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("PROJETO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("CARTAS/OFICIOS");
                        }
                        else if (f.Contains("_SGD_"))
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("EXECUCAO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("SGD/TET");
                        }
                        else if (f.Contains("CLIENTE_VITAL"))
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("PROJETO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("CARTAS/OFICIOS");
                        }
                        else if (f.Contains("_LOCACAO_") || f.Contains("_APR_") || f.Contains("_DESENHO_"))
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("EXECUCAO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("LOCACAO");
                        }
                        else
                        {
                            SelectElement selectCat = new SelectElement(Chrome.driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("EXECUCAO");
                            SelectElement selectDoc = new SelectElement(Chrome.driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("PONTO DE SERVICO");
                        }
                        try
                        {
                            // Anexa o arquivo e clica no botão "Enviar..."
                            Chrome.driver.FindElement(By.Id("fileUPArquivo")).SendKeys(Gomnet.folder + @"\" + f + ext);
                            Chrome.driver.FindElement(By.Id("Button_Anexar")).Click();
                        }
                        catch (System.InvalidOperationException)
                        {
                            Console.WriteLine("{0} not found in folder.", f);
                            continue;
                        }
                        try
                        {
                            var status = Chrome.driver.FindElement(By.XPath("*//a[contains(text(), '" + f + ext + "')]"));
                            if (status.Displayed)
                            {
                                Console.WriteLine("File {0}  attached sucessfully.", f);
                                Screenshot ss = ((ITakesScreenshot)Chrome.driver).GetScreenshot();
                                ss.SaveAsFile(f + ".png", ScreenshotImageFormat.Tiff);
                            }
                        }
                        catch (NoSuchElementException)
                        {
                            if (!File.Exists(Gomnet.logUpload))
                            {
                                using (StreamWriter sw = File.AppendText(Gomnet.logUpload))
                                {
                                    sw.WriteLine(f);
                                    sw.Close();
                                }
                            }
                        }
                    }
                }
            }
            //driver.Close();
            Console.ReadKey();
        }
    }
}
