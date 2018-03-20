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
    class ExcelWriter
    {
        static void Main(string[] args)
        {
            string planilha = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\sobs.xlsm";
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook wb = excel.Workbooks.Open(planilha);
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;

            var totalBaremos = ws.UsedRange.Columns["D"].Rows.Count; // Pega o total de linhas da coluna D (Baremos)


            var baremos = ws.UsedRange.Columns["D"]; // Pega todos os valores contidos na coluna D (Baremos)
            var qtds = ws.UsedRange.Columns["E"]; // Pega todos os valores contidos na coluna E (Quantidade)

            //Faz um loop através das duas colunas, utilizando o ZIP, e retorna uma lista contendo os valores das mesmas.
            for (int i = 2; i <= totalBaremos; i++) // O 'i' começa com '2' para ignorar a primeira linha "Baremos" e "Quantidade"
            {
                Excel.Range baremo = baremos.Cells[i];
                Excel.Range qtd = qtds.Cells[i];

                foreach (var valor in baremo.Zip(qtd, Tuple.Create))
                {
                    Console.WriteLine(Convert.ToString(valor.Item1.Value2) + " " + Convert.ToString(valor.Item2.Value2));
                }
            }

            //Fecha o excel e o processo .exe
            excel.Quit();
            excel.Dispose();

            Console.ReadKey();
        }
    }
}
