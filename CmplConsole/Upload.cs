using System;
using System.IO;
using System.Linq;
using Excel = NetOffice.ExcelApi;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace CmplConsole
{
    
    class Upload
    {
        static void Main(string[] args)
        {
            //Variáveis
            string url = "http://gomnet.ampla.com/";
            string urlUpload = "http://gomnet.ampla.com/Upload.aspx?numsob=";
            string folder = Path.GetDirectoryName(Application.ExecutablePath);
            string log = folder + @"\log.txt";

            // ----  Chrome Settings ---- //
                ChromeOptions options = new ChromeOptions();
                //options.AddArgument("--headless");
                options.AddArgument("--window-size= 1600x900");
            // ----  Chrome Settings ---- //

            //Informações de credenciais para login no GOMNET
            string credenciais = @"C:\gomnet.xlsx";
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook workbook = excel.Workbooks.Open(credenciais);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            string login = worksheet.Range("A1").Value.ToString();
            string senha = worksheet.Range("A2").Value.ToString();
            
            // Close excel and dispose reference
            excel.Quit();
            excel.Dispose();

            //Inicia o webdriver do Chrome
            IWebDriver driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl(url);

            //Loga no sistema
            IWebElement usrname = driver.FindElement(By.Id("txtBoxLogin"));
            IWebElement usrpass = driver.FindElement(By.Id("txtBoxSenha"));
            usrname.SendKeys(login);
            usrpass.SendKeys(senha);
            driver.FindElement(By.Id("ImageButton_Login")).Click();
            
            var files = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
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
                        driver.Navigate().GoToUrl(urlUpload + result);
                    }
                    else
                    {
                        driver.Navigate().GoToUrl(urlUpload + String.Join("_", s2.Skip(1).Take(2)));
                    }
                }
                else if (f.StartsWith("SG_REF") || f.StartsWith("SG_QUAL") || f.StartsWith("SG_INTER") || f.StartsWith("SG_RNT"))
                {
                    var s2 = file.Split('_');
                    driver.Navigate().GoToUrl(urlUpload + String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_PQ"))
                {
                    var s2 = file.Split('_');
                    driver.Navigate().GoToUrl(urlUpload + String.Join("_", s2.Skip(1).Take(3)));
                }
                else if (f.StartsWith("SG_RRC_"))
                {
                    var s2 = file.Split('_');
                    driver.Navigate().GoToUrl(urlUpload + String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_RRC"))
                {
                    var s2 = file.Split('_');
                    driver.Navigate().GoToUrl(urlUpload + String.Join("_", s2.Skip(1).Take(2)).Replace("_", "-"));
                }
                else if (f.StartsWith("SG_"))
                {
                    var s2 = file.Split('_');
                    driver.Navigate().GoToUrl(urlUpload + String.Join("_", s2.Skip(1).Take(1)));
                }

                try //Verifica se a sob foi digitada corretamente
                {
                    var erro = driver.FindElement(By.XPath("*//tr/td[contains(text(), 'Não existem dados para serem exibidos.')]"));
                    if (erro.Displayed)
                    {
                        Console.WriteLine("{0} not found.", f);
                    }
                }
                catch (NoSuchElementException)
                {
                    try //Verifica se o arquivo já foi anexado
                    {
                        var anexo = driver.FindElement(By.XPath("*//a[contains(text(), '" + f + "')]"));
                        if (anexo.Displayed)
                        {
                            Console.WriteLine("{0} already attached. ", f);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        if (f.Contains("FORM_FISC"))
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("ENCERRAMENTO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("FORMULARIO DE FISCALIZACAO DE OBRA");
                        }
                        else if (f.Contains("AS_BUILT"))
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("ENCERRAMENTO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("AS BUILT");
                        }
                        else if (f.Contains("APOIO_TRANSITO"))
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("PROJETO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("CARTAS/OFICIOS");
                        }
                        else if (f.Contains("CCR"))
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("PROJETO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("CARTAS/OFICIOS");
                        }
                        else if (f.Contains("_SGD_"))
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("EXECUCAO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("SGD/TET");
                        }
                        else if (f.Contains("CLIENTE_VITAL"))
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("PROJETO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("CARTAS/OFICIOS");
                        }
                        else if (f.Contains("_LOCACAO_") || f.Contains("_APR_") || f.Contains("_DESENHO_"))
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("EXECUCAO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("LOCACAO");
                        }
                        else
                        {
                            SelectElement selectCat = new SelectElement(driver.FindElement(By.Id("drpCategoria")));
                            selectCat.SelectByText("EXECUCAO");
                            SelectElement selectDoc = new SelectElement(driver.FindElement(By.Id("DropDownList1")));
                            selectDoc.SelectByText("PONTO DE SERVICO");
                        }
                        try
                        {
                            // Anexa o arquivo e clica no botão "Enviar..."
                            driver.FindElement(By.Id("fileUPArquivo")).SendKeys(folder + @"\" + f + ext);
                            driver.FindElement(By.Id("Button_Anexar")).Click();
                        }
                        catch (System.InvalidOperationException)
                        {
                            Console.WriteLine("{0} not found in folder.", f);
                            continue;
                        }
                        try
                        {
                            var status = driver.FindElement(By.XPath("*//a[contains(text(), '" + f + ext + "')]"));
                            if (status.Displayed)
                            {
                                Console.WriteLine("File {0}  attached sucessfully.", f);
                                Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
                                ss.SaveAsFile(f + ".png", ScreenshotImageFormat.Tiff);
                            }
                        }
                        catch (NoSuchElementException)
                        {
                            if (!File.Exists(log))
                            {
                                using (StreamWriter sw = File.AppendText(log))
                                {
                                    sw.WriteLine(f);
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
