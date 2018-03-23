using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

            //Headless Mode
            ChromeOptions options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", folder);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            //options.AddArgument("--headless");
            options.AddArgument("--window-size= 1600x900");

            //Informações de credenciais para login no GOMNET
            string credenciais = @"C:\gomnet.xlsx";
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook workbook = excel.Workbooks.Open(credenciais);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            string login = worksheet.Range("A1").Value.ToString();
            string senha = worksheet.Range("A2").Value.ToString();

            // close excel and dispose reference
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

            //Tira screenshot da página atual
            //Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            //ss.SaveAsFile("filename.png", ScreenshotImageFormat.Tiff);
            
            var files = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(s => s.EndsWith(".pdf") || s.EndsWith(".PDF"));
            foreach (string file in files)
            {
                string f = Path.GetFileNameWithoutExtension(file);
                if (f.StartsWith("SG_QUAL_"))
                {
                    var s2 = file.Split('_');
                    var s3 = String.Join("_", s2.Skip(1).Take(3));
                    int lastIndex = s3.LastIndexOf('_');
                    if (Char.IsDigit(s3[lastIndex + 1]))
                    {
                        string result = s3.Remove(lastIndex, "_".Length).Insert(lastIndex, "-");
                        driver.Navigate().GoToUrl(urlUpload + result);
                    } else
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
                } catch (NoSuchElementException)
                {
                    try //Verifica se o arquivo já foi anexado
                    {
                        var anexo = driver.FindElement(By.XPath("*//a[contains(text(), '" + f + "')]"));
                        if (anexo.Displayed)
                        {
                            Console.WriteLine("{0} already attached. ");
                        }
                    } catch (NoSuchElementException)
                    {
                        if (f.Contains("FORM_FISC"))
                        {
                            //var categoria =  driver.FindElement(By.Id("drpCategoria"));
                            //SelectElement selecCat = new SelectElement();
                            //
                            //
                            //
                            //categoria.select_by_visible_text('ENCERRAMENTO')
                            //documento = Select(driver.find_element_by_id('DropDownList1'))
                            //documento.select_by_visible_text('FORMULARIO DE FISCALIZACAO DE OBRA')
                        }
                    }
                }
            }

            //driver.Close();
            Console.ReadKey();
        }
    }
}
