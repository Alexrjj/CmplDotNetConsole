using OfficeOpenXml;
using OpenQA.Selenium;
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
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            // Abre e lê o arquivo xlsx.
            using (Gomnet.pacoteTrabalho)
            {
                if (Gomnet.pasta.Worksheets.Count > 0)
                {
                    var pastaTrabalho = Gomnet.pasta.Worksheets.First();
                    var linhas = pastaTrabalho.Dimension.End.Row;

                    for (int linha = 2; linha <= linhas; linha++)
                    {
                        var cell_in = (double)pastaTrabalho.Cells[linha, 7].Value;
                        var cell_fin = (double)pastaTrabalho.Cells[linha, 8].Value;
                        var hr_in = DateTime.FromOADate(cell_in).TimeOfDay.ToString().Split(':');
                        var hr_fin = DateTime.FromOADate(cell_fin).TimeOfDay.ToString().Split(':');
                        string x = pastaTrabalho.Cells[linha, 3].Value.ToString();
                        string doisZeros = Regex.Replace(x, @"\A131|\A100", "00$&"); // Para sobs iniciadas em "100" ou "131", acrescenta dois zeros no início da mesma.
                        string cincoZeros = Regex.Replace(doisZeros, @"\A[1-4]", "00000$&"); // Para sobs iniciadas em 1-4, acrescenta cinco zeros no início da mesma.
                        string sobFinal = Regex.Split(cincoZeros, @"[\s\.]")[0]; // Delimita espaço quando há duas sobs na mesma célula.
                        
                        try
                        {
                            // Verifica o status da Sob antes de programar
                            Chrome.driver.Navigate().GoToUrl(Gomnet.urlConsulta);
                            var consultaStatusSob = Chrome.driver.FindElement(By.Name("ctl00$ContentPlaceHolder1$TextBox_NumSOB"));
                            consultaStatusSob.SendKeys(sobFinal);
                            Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();

                            IWebElement sobDespachada = Chrome.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/form/table/tbody/tr[4]/td/div[3]/table/tbody/tr[2]/td[8][contains(text(), '" + sobFinal + "')]")));
                            if (sobDespachada.Displayed)
                            {
                                // Se o status for diferente de "Vistoria", informa a sob já programada e pula para a próxima. Do contrário, acessa a página referente à programação para dar prosseguimento.
                                string statusSob = Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_Gridview_GomNet1']/tbody/tr[2]/td[3]")).Text;
                                if (statusSob != "Vistoria")
                                {
                                    Console.WriteLine(sobFinal + " já programada.");
                                    continue;
                                } else {
                                    //Acessa a página de programação
                                    Chrome.driver.Navigate().GoToUrl(Gomnet.urlProgObraVist);
                                    var sob = Chrome.driver.FindElement(By.Name("ctl00$ContentPlaceHolder1$TextBox_NumSOB"));
                                    // Envia a variável apenas quando há dados na célula, do contrário, informa que não há mais Sobs.
                                    try
                                    {
                                        sob.SendKeys(sobFinal);
                                    }
                                    catch (NullReferenceException)
                                    {
                                        // Se não houver mais célula com dados na Programação Filtrada, informa a mensagem abaixo e finaliza o programa.
                                        Console.WriteLine("Não há mais nada a programar.");
                                        break;
                                    }
                                    Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ImageButton_Enviar")).Click();
                                    // Reconhece a linha onde foi despachado para a COMPEL, através do XPATH, e clica no ícone de programação da mesma. Se não houver, retorna a sob com erro.
                                    try
                                    {
                                        IWebElement parceiraCompel = Chrome.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//*[text() = 'COMPEL CONSTRUÇÕES MONTAGENS E']/following::td[17]")));
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
                                    Chrome.js.ExecuteScript("arguments[0].value = '" + DateTime.Now.ToString("dd/MM/yyyy ").ToString() + hr_in[0] + ":" + hr_in[1] + "'", dataInicial);

                                    IWebElement dataFinal = Chrome.driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Control_DataHora_TerminoPrevisto_TextBox6"));
                                    Chrome.js.ExecuteScript("arguments[0].value = '" + DateTime.Now.ToString("dd/MM/yyyy ").ToString() + hr_fin[0] + ":" + hr_fin[1] + "'", dataFinal);

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
                                            IWebElement barQtd = Chrome.driver.FindElement(By.XPath("*//tr/td[text() = '" + valor.Item1 + "']/following::td[2]/input[@type='text']"));
                                            if (bar.Displayed)
                                            {
                                                bar.Click();
                                                barQtd.Clear();
                                                barQtd.SendKeys(valor.Item2);
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
                                    //Chrome.driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_btnEnviarItens']")).Click();
                                    Console.WriteLine(sobFinal + " programada com êxito.");
                                }
                            }
                        }
                        catch (WebDriverTimeoutException)
                        {
                            Console.WriteLine(sobFinal + " ainda não vistoriada. Impossível programá-la.");
                            continue;
                        }
                    }
                }
            }
        }

        public static void GeraPedSAP()
        {
            // TODO: Implementar criação de pedido de materiais no SAP
            // TODO: Alterar quantidade de material solicitado de acordo com valores contidos numa planilha xls. Gerar erro caso a quantidade solicitada seja maior que orçada.
            Gomnet.Settings();
            Chrome.Initializer();
            Gomnet.Login();

            // Abre e lê o arquivo xlsx.
            using (Gomnet.pacoteTrabalho)
            {
                if (Gomnet.pasta.Worksheets.Count > 0)
                {
                    var pastaTrabalho = Gomnet.pasta.Worksheets.First();
                    var linhas = pastaTrabalho.Dimension.End.Row;

                    for (int linha = 2; linha <= linhas; linha++)
                    {
                        var cell_in = (double)pastaTrabalho.Cells[linha, 7].Value;
                        var cell_fin = (double)pastaTrabalho.Cells[linha, 8].Value;
                        var hr_in = DateTime.FromOADate(cell_in).TimeOfDay.ToString().Split(':');
                        var hr_fin = DateTime.FromOADate(cell_fin).TimeOfDay.ToString().Split(':');
                        string x = pastaTrabalho.Cells[linha, 3].Value.ToString();
                        string doisZeros = Regex.Replace(x, @"\A131|\A100", "00$&"); // Para sobs iniciadas em "100" ou "131", acrescenta dois zeros no início da mesma.
                        string cincoZeros = Regex.Replace(doisZeros, @"\A[1-4]", "00000$&"); // Para sobs iniciadas em 1-4, acrescenta cinco zeros no início da mesma.
                        string sobFinal = Regex.Split(cincoZeros, @"[\s\.]")[0]; // Delimita espaço quando há duas sobs na mesma célula.

                    }
                }
            }
        }

        public static void FinalizaGomMob()
        {
            //TODO: Implementar finalização de sobs programadas
        }

        public static void EnergizaGomMob()
        {
            //TODO: Implementar energização de sobs concluídas
        }
    }
}
