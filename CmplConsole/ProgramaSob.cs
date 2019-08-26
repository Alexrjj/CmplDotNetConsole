using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Actions = OpenQA.Selenium.Interactions.Actions;

namespace CmplConsole
{
    class ProgramaSob
    {
        public static void Vistoria()
        {
            // TODO: Implementar programação da sob em estado de vistoria
        }

        public static void GeraPedSAP()
        {
            // TODO: Implementar criação de pedido de materiais no SAP
            // TODO: Alterar quantidade de material solicitado de acordo com valores contidos numa planilha xls. Gerar erro caso a quantidade solicitada seja maior que orçada.
        }

        public static void Execucao()
        {
            // TODO: Consultar a sob antes de programar, para saber se o status está como Fechado ou Certificado, evitando nova programação da mesma. (Gerar retorno e pular para próxima sob)
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();
            WebDriverWait wait = new WebDriverWait(Chrome.driver, TimeSpan.FromSeconds(5));
            Actions action = new Actions(Chrome.driver);

            // string planilha = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\sobs.xlsm";
            // Cria varíavel do xlsx atribuindo endereço local.
            var arquivoXlsx = new FileInfo(@"Programação Filtrada.xlsx");
            // Abre e lê o arquivo xlsx.
            using (var pacote = new ExcelPackage(arquivoXlsx))
            {
                // Cria a variável da pasta a ser trabalhada.
                var pasta = pacote.Workbook;
                //int count = pastas.Worksheets.Count; // Pega o valor total de pastas

                if (pasta.Worksheets.Count > 0)
                    {
                    //var pasta = pastas.Worksheets[p]; // Cada pasta assume o número atual em "count", que deve ser iniciado em 1
                    var pastaTrabalho = pasta.Worksheets.First();
                    var linhas = pastaTrabalho.Dimension.End.Row;

                    for (int linha = 1; linha <= linhas; linha++)
                    {
                        string x = pastaTrabalho.Cells[linha, 3].Value.ToString();
                        string doisZeros = Regex.Replace(x, @"\A131|\A100", "00$&"); // Para sobs iniciadas em "100" ou "131", acrescenta dois zeros no início da mesma.
                        string cincoZeros = Regex.Replace(doisZeros, @"\A[1-4]", "00000$&"); // Para sobs iniciadas em 1-4, acrescenta cinco zeros no início da mesma.
                        string sobFinal = Regex.Split(cincoZeros, @"[\s\.]")[0]; // Delimita espaço quando há duas sobs na mesma célula.
                        Chrome.driver.Navigate().GoToUrl(Gomnet.urlAcompObra);

                        var sob = Chrome.driver.FindElement(By.Name("ctl00$ContentPlaceHolder1$TextBox_NumSOB"));
                        // Envia a variável apenas quando há dados na célula, do contrário, informa que não há mais Sobs.
                        try
                        {
                            sob.SendKeys(sobFinal); 
                        }
                        catch (NullReferenceException)
                        {
                            Console.WriteLine("Não há mais nada a programar.");
                            break;
                        }
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();
                        // Reconhece a linha onde foi despachado para a COMPEL, através do XPATH, e clica no ícone de programação da mesma. Se não houver, retorna a sob com erro.
                        try
                        {
                            IWebElement parceiraCompel = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[text() = 'COMPEL CONSTRUÇÕES MONTAGENS E']/following::td[31]")));
                            if (parceiraCompel.Displayed)
                            {
                                parceiraCompel.Click();
                            }
                        }
                        catch (WebDriverTimeoutException)
                        {
                            Console.WriteLine(sobFinal + " não há registro. Favor verificar.");
                            continue;
                        }

                        // Faz um loop desde a primeira linha com código baremo até a última.
                        for (int i = 1; i <= linhas; i++)
                        {
                            var baremo = pastaTrabalho.Cells[i, 3]; // Atribui a variável Sob à coluna 03 do arquivo, onde constam as Sobs.
                            var qtd = pastaTrabalho.Cells[i, 4]; // Atribui a variável Data à coluna 04 do arquivo, onde constam as Datas.
                            //                                     // Lê alternadamente a célula das duas colunas e cria uma tupla.
                            foreach (var valor in baremo.Zip(qtd, Tuple.Create))
                            {
                                // Verifica se a linha contém dados.
                                if (valor.Item1.Text != null && valor.Item2.Text != null)
                                {
                                    // Escreve na tela o valor (texto) de cada célula.
                                    Console.WriteLine(Convert.ToString(valor.Item1.Text) + " " + Convert.ToString(valor.Item2.Text));
                                }
                            }
                        }
                    }
                }
            }
        }
        public static void Energizacao()
        {
            //TODO: Implementar energização de sobs concluídas
        }
    }
}
