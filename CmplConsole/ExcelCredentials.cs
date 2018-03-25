using Excel = NetOffice.ExcelApi;

namespace CmplConsole
{
    public class ExcelCredentials
    {
        public static Excel.Application excel = new Excel.Application();
        public static string login;
        public static string senha;

        public static void LoginGomnet()
        {
            //Informações de credenciais para login no GOMNET
            string credenciais = @"C:\gomnet.xlsx";
            excel.DisplayAlerts = false;
            Excel.Workbook workbook = excel.Workbooks.Open(credenciais);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            login = worksheet.Range("A1").Value.ToString();
            senha = worksheet.Range("A2").Value.ToString();
        }
    }
}
