using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = NetOffice.ExcelApi;
using System.Windows.Forms;

namespace CmplConsole
{
    class ExcelReader
    {
        public static void LerBaremoQtd()
        {
            string planilha = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\sobs.xlsm";
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook wb = excel.Workbooks.Open(planilha);
            int count = wb.Worksheets.Count; // Pega o valor total de pastas
            
            for (int g = 1; g <= count; g++) // Inicia o loop em cada pasta
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[g]; // Cada pasta assume o número atual em "count", que deve ser iniciado em 1

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
                        if (valor.Item1.Value2 != null && valor.Item2.Value2 != null) // Ignora células vazias (null)
                        {
                            Console.WriteLine(Convert.ToString(valor.Item1.Value2) + " " + Convert.ToString(valor.Item2.Value2));
                        }
                    }
                }
            }

            //Fecha o excel e o processo .exe
            excel.Quit();
            excel.Dispose();

            Console.ReadKey();
        }
    }
}
