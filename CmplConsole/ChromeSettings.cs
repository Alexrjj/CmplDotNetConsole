using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Windows.Forms;

namespace CmplConsole
{
    public class ChromeSettings
    {
        public static ChromeOptions options = new ChromeOptions();
        public static IWebDriver driver;

        public static void ChromeNonHeadless()
        {
            options.AddArgument("start-maximized");
            options.AddUserProfilePreference("download.default_directory", Path.GetDirectoryName(Application.ExecutablePath));
            options.AddUserProfilePreference("download.prompt_for_download", false);
        }

        public static void ChromeHeadless()
        {
            options.AddArgument("--headless");
            options.AddArgument("--window-size= 1600x900");
            options.AddArgument("start-maximized");
            options.AddUserProfilePreference("download.default_directory", Path.GetDirectoryName(Application.ExecutablePath));
            options.AddUserProfilePreference("download.prompt_for_download", false);
        }

        public static void ChromeInitializer()
        {
            char mode = '\0';
            Console.WriteLine("Headless Mode? (y/n)");
            mode = Convert.ToChar(Console.ReadLine());
            if (mode == 'y')
            {
                ChromeSettings.ChromeHeadless();
                driver = new ChromeDriver(options);
            }
            else if (mode == 'n')
            {
                ChromeSettings.ChromeNonHeadless();
                driver = new ChromeDriver(options);
            } else
            {
                Console.WriteLine("Invalid option.");
                Console.ReadKey();
            }
        }
    }
}
