using OfficeOpenXml;
using OpenQA.Selenium;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;


namespace CmplConsole
{
    public class Gomnet
    {
        public static string urlLogin;
        public static string urlConsulta;
        public static string urlUpload;
        public static string urlVincSup;
        public static string urlAcompObra;
        public static string folder;
        // public static Excel.Application excel;
        public static string login;
        public static string senha;
        public static string InSobStatus;
        public static string InSobTrab;
        public static string InSobDWG;
        public static string InSobSGD;
        public static string InSobSup;
        public static string OutUpload;
        public static string OutSobTrab;
        public static string OutSobStatus;

        public static void Settings()
        {
            urlLogin = "http://gomnet.ampla.com/";
            urlConsulta = "http://gomnet.ampla.com/ConsultaObra.aspx";
            urlUpload = "http://gomnet.ampla.com/Upload.aspx?numsob=";
            urlVincSup = "http://gomnet.ampla.com/vistoria/vincularSupervisor.aspx";
            urlAcompObra = "http://gomnet.ampla.com/AcompanhamentoObras.aspx";
            folder = Path.GetDirectoryName(Application.ExecutablePath);
            InSobStatus = folder + @"\InSobStatus.txt";
            InSobTrab = folder + @"\InSobTrab.txt";
            InSobDWG = folder + @"\InSobDWG.txt";
            InSobSGD = folder + @"\InSobSGD.txt";
            InSobSup = folder + @"\InSobSup.txt";
            OutSobStatus = folder + @"\OutSobStatus.txt";
            OutSobTrab = folder + @"\OutSobTrab.txt";
            OutUpload = folder + @"\OutUpload.txt";
        }

        public static void Menu()
        {
            Console.WriteLine("(1) Upload Files");
            Console.WriteLine("(2) Download DWG");
            Console.WriteLine("(3) Download SGD");
            Console.WriteLine("(4) Link Supervisor");
            Console.WriteLine("(5) Inspection Sob");
            Console.WriteLine("(6) Program Sob (Inspection)");
            Console.WriteLine("(7) Program Sob (Execution)");
            Console.WriteLine("(8) Extract Sob Status");
            Console.WriteLine("(9) Extract Sob Work Number");
            Console.WriteLine("(10) Extract Sob Status");
            Console.WriteLine("(11) Verify Energized Sob");
            Console.WriteLine("(12) DWG File Filter");
            Console.WriteLine("(13) PDF File Filter");
            Console.Write("\nChoose your option: ");
            var key = Console.ReadLine();
            switch (key)
            {
                case "1":
                    {
                        Upload.Arquivos();
                        break;
                    }
                case "2":
                    {
                        Download.DWG();
                        break;
                    }
                case "3":
                    {
                        Download.SGD();
                        break;
                    }
                case "4":
                    {
                        VistoriaSob.VincSupervisor();
                        break;
                    }
                case "5":
                    {

                        break;
                    }
                case "6":
                    {

                        break;
                    }
                case "7":
                    {
                        ProgramaSob.Execucao();
                        break;
                    }
                case "8":
                    {
                        Extrair.SobStatus();
                        break;
                    }
                case "9":
                    {
                        Extrair.SobTrabalho();
                        break;
                    }
                case "10":
                    {

                        break;
                    }
                case "11":
                    {

                        break;
                    }
                case "12":
                    {
                        FiltraArquivo.DWG();
                        break;
                    }
                case "13":
                    {
                        FiltraArquivo.PDF();
                        break;
                    }
            }
        }
        public static void Login()
        {
            // Cria varíavel do xlsx atribuindo endereço local.
            var arquivoXlsx = new FileInfo(@"C:\gomnet.xlsx");
            // Abre e lê o arquivo xlsx.
            using (var pacote = new ExcelPackage(arquivoXlsx))
            {
                // Cria a variável da pasta a ser trabalhada.
                var pasta = pacote.Workbook;
                if (pasta != null) // Verifica se há alguma pasta de trabalho existente no xlsx.
                {
                    if (pasta.Worksheets.Count > 0)
                    {
                        // Pega a primeira pasta de trabalho.
                        var pastaTrabalho = pasta.Worksheets.First();
                        login = pastaTrabalho.Cells["A1"].Text;
                        senha = pastaTrabalho.Cells["A2"].Text;

                        //Inicia o webdriver do Chrome
                        try
                        {
                            Chrome.driver.Navigate().GoToUrl(Gomnet.urlLogin);
                        }
                        catch (System.NullReferenceException) // Pára o script, caso falhe a opção escolhida no ChromeInitializer
                        {
                            return;
                        }

                        //Loga no sistema
                        IWebElement usrname = Chrome.driver.FindElement(By.Id("txtBoxLogin"));
                        IWebElement usrpass = Chrome.driver.FindElement(By.Id("txtBoxSenha"));
                        usrname.SendKeys(login);
                        usrpass.SendKeys(senha);
                        Chrome.driver.FindElement(By.Id("ImageButton_Login")).Click();
                    }
                }
            }
        }
        public static void FechaJanela()
        {
            Chrome.driver.SwitchTo().Window(Chrome.driver.WindowHandles.Last());
            Chrome.driver.Close();
            Chrome.driver.SwitchTo().Window(Chrome.driver.WindowHandles.First());
            Thread.Sleep(5000);
        }
    }
}