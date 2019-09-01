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

                    for (int linha = 2; linha <= linhas; linha++)
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

                        //Insere o valor da variável na textbox utilizando Javascript, otimizando a velocidade de exdecução
                        IWebElement turma = Chrome.driver.FindElement(By.CssSelector("#ctl00_ContentPlaceHolder1_txtBoxTurma"));
                        Chrome.js.ExecuteScript("arguments[0].value = 'COMP12'", turma);

                        IWebElement servico = Chrome.driver.FindElement(By.CssSelector("#ctl00_ContentPlaceHolder1_txtBoxRespServico"));
                        Chrome.js.ExecuteScript("arguments[0].value = 'A3279'", servico);

                        IWebElement servicoSup = Chrome.driver.FindElement(By.CssSelector("#ctl00_ContentPlaceHolder1_txtBoxServicoSuplente"));
                        Chrome.js.ExecuteScript("arguments[0].value = 'A3279'", servicoSup);

                        IWebElement dataInicial = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Control_DataHora_InicioPrevisto_TextBox6"));
                        Chrome.js.ExecuteScript("arguments[0].value = '31/08/2019 00:00'", dataInicial);

                        IWebElement dataFinal = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Control_DataHora_TerminoPrevisto_TextBox6"));
                        Chrome.js.ExecuteScript("arguments[0].value = '31/08/2019 00:05'", dataFinal);


                        // Adiciona técnicos à tarefa
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_imgBtnGravarTecnicos")).Click();
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_gridViewTecnicos_ctl01_ChkBoxAll")).Click();
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Button1")).Click();
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_btnVoltar")).Click();


                        // Identifica o menu "Tipo de Programação" e seleciona a opção "Execução de Obra" utilizando Javascript
                        IWebElement tipoProg = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_DropDownList_TipoProgramacao"));
                        Chrome.js.ExecuteScript("arguments[0].value = 'N'", tipoProg);

                        // Identifica o menu "Necessita Linha Viva" e seleciona a opção "Não" utilizando Javascript
                        IWebElement linhaViva = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_DropDownList_linhaViva"));
                        Chrome.js.ExecuteScript("arguments[0].value = 'N'", linhaViva);

                        // Identifica o menu "Necessita Desligamento" e seleciona a opção "Não" utilizando Javascript
                        IWebElement desliga = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_DropDownList_Desliga"));
                        Chrome.js.ExecuteScript("arguments[0].value = 'N'", desliga);

                        // Clica no botão "Programar"
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Button_ProgramarTarefa")).Click();

                        // Preenche o campo "Atividade" com o número da SOB utilizando Javascript
                        IWebElement atividade = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_txtBoxAtividade"));

                        // atividade.send_keys(sob_tratada)
                        Chrome.js.ExecuteScript("arguments[0].value = '" + sobFinal + "'", atividade);

                        // Identifica o menu "Horas" e insere o valor "01" utilizando Javascript
                        IWebElement hora = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_DropDownList_Hora"));
                        Chrome.js.ExecuteScript("arguments[0].value = '01'", hora);

                        // Identifica o menu "Minutos" e insere o valor "00" utilizando Javascript
                        IWebElement minuto = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_DropDownList_Minuto"));
                        Chrome.js.ExecuteScript("arguments[0].value = '00'", minuto);

                        // Identifica o menu "Viagem" e insere o valor "Não" utilizando Javascript
                        IWebElement viagem = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_DropDownList_Viagem"));
                        Chrome.js.ExecuteScript("arguments[0].value = 'N'", viagem);

                        // Clica no botão "Adicionar Programação"
                        Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_btnAdicionarProgramacao")).Click();
                        
                        // Insere pelo menos um baremo/material e confirma a programação
                        foreach (var valor in BaremosModelo.baremo.Zip(BaremosModelo.qtd, Tuple.Create))
                        {
                            try
                            {
                                IWebElement bar = Chrome.driver.FindElement(By.XPath("*//tr/td[contains(text(), '" + valor.Item1 + "')]/preceding-sibling::td/input"));
                                if (bar.Displayed)
                                {
                                    bar.Click();
                                    action.SendKeys(Keys.Tab).Perform();
                                    action.SendKeys(valor.Item2).Perform();
                                    break;
                                }
                            }
                            catch (NoSuchElementException)
                            {
                                continue;
                            }
                            continue;
                        }
                        // Ao fim do loop de inserção de baremos, clica no botão "registrar programação"
                        // Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_btnEnviarItens']")).Click();
                        Console.WriteLine(sobFinal + " programada com êxito.");
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
