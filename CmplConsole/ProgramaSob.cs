using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Linq;
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

        [Obsolete]
        public static void Execucao()
        {
            // TODO: Consultar a sob antes de programar, para saber se o status está como Fechado ou Certificado, evitando nova programação da mesma. (Gerar retorno e pular para próxima sob)
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();
            WebDriverWait wait = new WebDriverWait(Chrome.driver, TimeSpan.FromSeconds(3));
            Actions action = new Actions(Chrome.driver);

            // string planilha = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\sobs.xlsm";
            // Cria varíavel do xlsx atribuindo endereço local.
            var arquivoXlsx = new FileInfo(@"C:\gomnet.xlsx");
            // Abre e lê o arquivo xlsx.
            using (var pacote = new ExcelPackage(arquivoXlsx))
            {
                // Cria a variável da pasta a ser trabalhada.
                var pastas = pacote.Workbook;
                int count = pastas.Worksheets.Count; // Pega o valor total de pastas

                for (int p = 1; p <= count; p++) // Inicia o loop em cada pasta
                {
                    var pasta = pastas.Worksheets[p]; // Cada pasta assume o número atual em "count", que deve ser iniciado em 1

                    Chrome.driver.Navigate().GoToUrl(Gomnet.urlAcompObra);

                    var sob = Chrome.driver.FindElement(By.Name("ctl00$ContentPlaceHolder1$TextBox_NumSOB"));

                    sob.SendKeys(pasta.Cells["A2"].Value.ToString());
                    Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

                    try
                    {
                        var compel = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[contains(text(), 'COMPEL CONSTRUÇÕES MONTAGENS E')]")));
                        if (compel.Displayed)
                        {
                            action.Click(compel).Perform();
                            int m = 0;
                            while (m <= 8)
                            {
                                action.SendKeys(Keys.Tab).Perform();
                                m += 1;
                            }
                            action.SendKeys(Keys.Space).Perform();
                        }
                    }
                    catch (NullReferenceException)
                    {
                        Console.WriteLine("Sob not available for Compel.");
                        break;
                    }
                    //break; // Pára o script ao chegar no preenchimendo de dados da programação

                    var linhas =  pasta.Dimension.End.Row;
                    // Faz um loop desde a primeira linha até a última.
                    for (int i = 1; i <= linhas; i++)
                    {
                        var baremo = pasta.Cells[i, 3]; // Atribui a variável Sob à coluna 03 do arquivo, onde constam as Sobs.
                        var qtd = pasta.Cells[i, 4]; // Atribui a variável Data à coluna 04 do arquivo, onde constam as Datas.
                                                              // Lê alternadamente a célula das duas colunas e cria uma tupla.
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
        public static void Energizacao()
        {
            //TODO: Implementar energização de sobs concluídas
        }
    }
}
