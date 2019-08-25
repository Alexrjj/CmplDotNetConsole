using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace CmplConsole
{
    class ExcelReader
    {
        public static void LerBaremoQtd()
        {
            // Cria varíavel do xlsx atribuindo endereço local.
            var arquivoXlsx = new FileInfo(@"Programação Filtrada.xlsx");
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
                        // Pega o número total de linhas da pasta de trabalho (considera a coluna com maior quantidade).
                        var linhas = pastaTrabalho.Dimension.End.Row;
                        // Faz um loop desde a primeira linha até a última.
                        for (int i = 2; i <= linhas; i++)
                        {
                            var sob = pastaTrabalho.Cells[i, 3]; // Atribui a variável Sob à coluna 03 do arquivo, onde constam as Sobs.
                            var data = pastaTrabalho.Cells[i, 4]; // Atribui a variável Data à coluna 04 do arquivo, onde constam as Datas.
                            // Lê alternadamente a célula das duas colunas e cria uma tupla.
                            foreach (var valor in sob.Zip(data, Tuple.Create))
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
            Console.ReadKey();
        }
    }
}
