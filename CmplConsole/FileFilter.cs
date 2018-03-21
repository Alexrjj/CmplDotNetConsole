using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CmplConsole
{
    class FileFilter
    {
        static void Main(string[] args)
        {
            string folder = Path.GetDirectoryName(Application.ExecutablePath);
            var files = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(s => s.EndsWith(".pdf") || s.EndsWith(".PDF"));
            foreach (string file in files)
            {
                string f = Path.GetFileNameWithoutExtension(file);
                if (f.StartsWith("SG_REF"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(Path.GetFileNameWithoutExtension(String.Join("_", s2.Skip(1).Take(2))));
                }
                else if (f.StartsWith("SG_RRC"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(Path.GetFileNameWithoutExtension(String.Join("_", s2.Skip(1).Take(2)).Replace("_", "-")));
                }
                else if (f.StartsWith("SG_"))
                {
                    var s2 = file.Split('_');  
                    Console.WriteLine(Path.GetFileNameWithoutExtension(String.Join("_", s2.Skip(1).Take(1))));
                }
                
                else
                {
                    Console.WriteLine("File not found");
                }
            }
                
                
            Console.ReadKey();
        }
    }
}
