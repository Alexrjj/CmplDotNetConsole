using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CmplConsole
{
    class FiltraArquivo
    {
        public static void DWG()
        {
            string folder = Path.GetDirectoryName(Application.ExecutablePath);
            var files = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(s => s.EndsWith(".dwg") || s.EndsWith(".DWG"));
            foreach (string file in files)
            {
                string f = Path.GetFileNameWithoutExtension(file);
                if (f.StartsWith("SG_QUAL_"))
                {
                    var s2 = file.Split('_');
                    var s3 = String.Join("_", s2.Skip(1).Take(3));
                    int lastIndex = s3.LastIndexOf('_');
                    if (Char.IsDigit(s3[lastIndex + 1]))
                    {
                        string result = s3.Remove(lastIndex, "_".Length).Insert(lastIndex, "-");
                        Console.WriteLine(result);
                    }
                    else
                    {
                        Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)));
                    }
                }
                else if (f.StartsWith("SG_REF") || f.StartsWith("SG_QUAL") || f.StartsWith("SG_INTER") || f.StartsWith("SG_RNT"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_PQ"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(3)));
                }
                else if (f.StartsWith("SG_RRC_"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_RRC"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)).Replace("_", "-"));
                }
                else if (f.StartsWith("SG_"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(1)));
                }
                else
                {
                    Console.WriteLine("{0} not found", f);
                }
            }
            Console.ReadKey();
        }
        public static void PDF()
        {
            string folder = Path.GetDirectoryName(Application.ExecutablePath);
            var files = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(s => s.EndsWith(".pdf") || s.EndsWith(".PDF"));
            foreach (string file in files)
            {
                string f = Path.GetFileNameWithoutExtension(file);
                if (f.StartsWith("SG_QUAL_"))
                {
                    var s2 = file.Split('_');
                    var s3 = String.Join("_", s2.Skip(1).Take(3));
                    int lastIndex = s3.LastIndexOf('_');
                    if (Char.IsDigit(s3[lastIndex + 1]))
                    {
                        string result = s3.Remove(lastIndex, "_".Length).Insert(lastIndex, "-");
                        Console.WriteLine(result);
                    }
                    else
                    {
                        Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)));
                    }
                }
                else if (f.StartsWith("SG_REF") || f.StartsWith("SG_QUAL") || f.StartsWith("SG_INTER") || f.StartsWith("SG_RNT"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_PQ"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(3)));
                }
                else if (f.StartsWith("SG_RRC_"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)));
                }
                else if (f.StartsWith("SG_RRC"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(2)).Replace("_", "-"));
                }
                else if (f.StartsWith("SG_"))
                {
                    var s2 = file.Split('_');
                    Console.WriteLine(String.Join("_", s2.Skip(1).Take(1)));
                }
                else
                {
                    Console.WriteLine("{0} not found", f);
                }
            }
            Console.ReadKey();
        }
    }
}
