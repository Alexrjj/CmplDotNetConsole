using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace CmplConsole
{
    class GoogleForms
    {
        public static void FormsGoogle()
        {

            Gomnet.Settings();
            Chrome.Initializer();
            WebDriverWait wait = new WebDriverWait(Chrome.driver, TimeSpan.FromSeconds(10));
            Chrome.driver.Navigate().GoToUrl(Gomnet.googleForm);

            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[2]/div/div[2]/div[3]/div[2]/content/span"))).Click();
            var email = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("identifierId")));
            email.SendKeys("alexsoares2517");
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='identifierNext']/content/span"))).Click();
            var senha = wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("password")));
            senha.SendKeys("1712887ab");
            Thread.Sleep(1000);
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("passwordNext"))).Click();
            Thread.Sleep(1000);
            //Data
            //var data = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='mG61Hd']/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div[1]/div/div[1]/input")));
            var data = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='mG61Hd']/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div[1]/div/div[1]/input")));
            Thread.Sleep(1000);
            data.SendKeys("26032018");

            //Base
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='mG61Hd']/div/div[2]/div[2]/div[3]/div[2]/div/content/div/label[6]/div/div[1]/div[3]/div"))).Click();

            //Botão Próximo
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[contains(text(), 'Próxima')]"))).Click();

            //Supervisor
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='mG61Hd']/div/div[2]/div[2]/div[2]/div[2]/div/content/div/label[19]/div/div[2]/div/span[contains(text(), 'THIAGO PASSOS')]"))).Click();

            //Botão Próximo
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[contains(text(), 'Próxima')]"))).Click();

            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='i1'][../following-sibling::'OBRAS']"))).Click();
        }
    }
}

