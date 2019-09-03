using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Windows.Forms;
using Actions = OpenQA.Selenium.Interactions.Actions;

namespace CmplConsole
{
    public class Chrome
    {
        public static ChromeOptions options = new ChromeOptions();
        public static IWebDriver driver;
        public static IJavaScriptExecutor js;
        public static WebDriverWait wait;
        public static Actions action;

        public static void NonHeadless()
        {
            options.AddArgument("start-maximized");
            options.AddUserProfilePreference("download.default_directory", Path.GetDirectoryName(Application.ExecutablePath));
            options.AddUserProfilePreference("download.prompt_for_download", false);
        }

        public static void Headless()
        {
            options.AddArgument("--headless");
            options.AddArgument("--window-size= 1600x900");
            options.AddArgument("start-maximized");
            options.AddUserProfilePreference("download.default_directory", Path.GetDirectoryName(Application.ExecutablePath));
            options.AddUserProfilePreference("download.prompt_for_download", false);
        }

        public static void Initializer()
        {
            try
            {
                if (Form1.Headless.Checked == true)
                {
                    Console.WriteLine("\n\nInitializing in Headless Mode...\n");
                    Headless();
                    driver = new ChromeDriver(options);
                    js = (IJavaScriptExecutor)driver;
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    action = new Actions(driver);
                }
                if (Form1.Headless.Checked == false)
                {
                    Console.WriteLine("\n\nInitializing in Normal Mode...");
                    NonHeadless();
                    driver = new ChromeDriver(options);
                    js = (IJavaScriptExecutor)driver;
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    action = new Actions(driver);
                }
            }
            catch
            {
                Console.WriteLine("Impossível continuar.");
            }
        }
    }
}