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

            var files = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories)
                .Where(s => s.EndsWith(".pdf") || s.EndsWith(".PDF"));
            foreach (string file in files)
            {
                var s1 = file.Split(new string[] { "_" }, StringSplitOptions.None);
                var s2 = s1[0].Split('_');
                Console.WriteLine(Path.GetFileNameWithoutExtension(String.Join("_", s2)));
            }
            Console.ReadKey();
        }
    }
}
