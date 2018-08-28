using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Excel = NetOffice.ExcelApi;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using OpenQA.Selenium.Support.UI;

namespace CmplConsole
{
    class ProgramaSob
    {
        public static void ProgramaVistoria()
        {
            // TODO: Implementar programação da sob em estado de vistoria
        }

        public static void GeraPedSAP()
        {
            // TODO: Implementar criação de pedido de materiais no SAP
            // TODO: Alterar quantidade de material solicitado de acordo com valores contidos numa planilha xls. Gerar erro caso a quantidade solicitada seja maior que orçada.
        }

        public static void ProgramaExecucao()
        {
            // TODO: Consultar a sob antes de programar, para saber se o status está como Fechado ou Certificado, evitando nova programação da mesma. (Gerar retorno e pular para próxima sob)
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            string planilha = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\sobs.xlsm";
            Excel.Application excelBaremos = new Excel.Application();
            excelBaremos.DisplayAlerts = false;
            Excel.Workbook wb = excelBaremos.Workbooks.Open(planilha);
            int count = wb.Worksheets.Count; // Pega o valor total de pastas

            for (int g = 1; g <= count; g++) // Inicia o loop em cada pasta
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[g]; // Cada pasta assume o número atual em "count", que deve ser iniciado em 1

                Chrome.driver.Navigate().GoToUrl(Gomnet.urlConsulta);

                var sob = Chrome.driver.FindElement(By.Name("ctl00$ContentPlaceHolder1$TextBox_NumSOB"));
                
                sob.SendKeys(ws.Range("A2").Value.ToString());
                Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

                Thread.Sleep(3000);
                var compel = Chrome.driver.FindElement(By.XPath("//*[contains(text(), 'COMPEL CONSTRUÇÕES MONTAGENS E')]"));
                if (compel.Displayed)
                {
                    Chrome.action.Click(compel).Perform();
                }
                
                int m = 0;
                while (m <= 8)
                {
                    Chrome.action.SendKeys(Keys.Tab).Perform();
                    m += 1;
                }
                
                
                var totalBaremos = ws.UsedRange.Columns["D"].Rows.Count; // Pega o total de linhas da coluna D (Baremos)

                var baremos = ws.UsedRange.Columns["D"]; // Pega todos os valores contidos na coluna D (Baremos)
                var qtds = ws.UsedRange.Columns["E"]; // Pega todos os valores contidos na coluna E (Quantidade)

                //Faz um loop através das duas colunas, utilizando o ZIP, e retorna uma lista contendo os valores das mesmas.
                for (int i = 2; i <= totalBaremos; i++) // O 'i' começa com '2' para ignorar a primeira linha "Baremos" e "Quantidade"
                {
                    Excel.Range baremo = baremos.Cells[i];
                    Excel.Range qtd = qtds.Cells[i];

                    foreach (var valor in baremo.Zip(qtd, Tuple.Create)) // Cria uma tupla contendo valor coluna1/coluna2 e inicia o loop
                    {                                                   // em todas as linhas.
                        if (valor.Item1.Value2 != null || valor.Item2.Value2 != null) // Ignora células vazias (null)
                        {
                            Console.WriteLine(Convert.ToString(valor.Item1.Value2) + " " + Convert.ToString(valor.Item2.Value2));
                        }
                    }
                }
            }

            //Fecha o excel e o processo .exe
            excelBaremos.Quit();
            excelBaremos.Dispose();

            Console.ReadKey();
        }

        public static void Energizacao()
        {
            //TODO: Implementar energização de sobs concluídas
        }
    }
}
