using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
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
            char mode = '\0';
            Console.WriteLine("Headless Mode? (y/n)");
            mode = Convert.ToChar(Console.ReadLine());
            if (mode == 'y')
            {
                Chrome.Headless();
                driver = new ChromeDriver(options);
                Actions action = new Actions(driver);
    }
            else if (mode == 'n')
            {
                Chrome.NonHeadless();
                driver = new ChromeDriver(options);
                Actions action = new Actions(driver);
            } else
            {
                Console.WriteLine("Invalid option.");
                Console.ReadKey();
            }
        }
    }
}
