using System;

namespace CmplConsole
{
    class Principal
    {
        public static void Main (string[] args)
        {
            Console.Write("Choose Option: \n");
            Console.WriteLine("(1) Upload Files");
            Console.WriteLine("(2) Download DWG");
            Console.WriteLine("(3) Download SGD");
            Console.WriteLine("(4) Link Supervisor");
            Console.WriteLine("(5) Inspection Sob");
            Console.WriteLine("(6) Program Sob (Inspection)");
            Console.WriteLine("(7) Program Sob (Execution)");
            Console.WriteLine("(8) Extract Sob Status");
            Console.WriteLine("(9) Extract Sob Work Number");
            Console.WriteLine("(10) Extract Sob Status");
            Console.WriteLine("(11) Verify Energized Sob");
            Console.WriteLine("(12) File Filter \n");
            var key = Console.ReadKey().KeyChar.ToString();
            switch (key)
            {
                case "1":
                    {
                        Upload.Arquivos();
                        break;
                    }
                case "2":
                    {
                        Download.DWG();
                        break;
                    }
                case "3":
                    {
                        Download.SGD();
                        break;
                    }
                case "4":
                    {
                        
                        break;
                    }
                case "5":
                    {
                        
                        break;
                    }
                case "6":
                    {

                        break;
                    }
                case "7":
                    {

                        break;
                    }
                case "8":
                    {
                        Extrair.SobStatus();
                        break;
                    }
                case "9":
                    {
                        Extrair.SobTrabalho();
                        break;
                    }
                case "10":
                    {

                        break;
                    }
                case "11":
                    {

                        break;
                    }
                case "12":
                    {

                        break;
                    }
            }
        }
    }
}
