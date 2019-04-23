using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace vsgoto
{
    class Program
    {
        public static readonly string SCHEMA = "vsgoto:";
        public static readonly Regex PARSER = new Regex("(.*):(\\d+)");

        static void Main(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
                Console.WriteLine($"%{i}: " + args[i]);

            try
            {

                if (args.Length < 1)
                {
                    ExportRegFile();
                    Console.WriteLine("Reg-File Generated...");
                    throw new Exception("Missing Arguments!");
                }

                if (!args[0].StartsWith(SCHEMA))
                    throw new Exception("Wrong Schema!");

                var uriFile = args[0].Substring(SCHEMA.Length, args[0].Length - SCHEMA.Length);

                var m = PARSER.Match(uriFile);

                var file = uriFile;
                var line = 1;

                if(m.Success)
                {
                    file = m.Groups[1].Value;
                    line = int.Parse(m.Groups[2].Value);
                }

                OpenFile(file, line);
            } catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        public static void OpenFile(string file, int line)
        {
            EnvDTE80.DTE2 dte2;
            dte2 = (EnvDTE80.DTE2)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE");
            dte2.MainWindow.Activate();
            EnvDTE.Window w = dte2.ItemOperations.OpenFile(file);
            ((EnvDTE.TextSelection)dte2.ActiveDocument.Selection).GotoLine(line, true);
        }

        public static void ExportRegFile()
        {
            File.WriteAllText("vsgoto.reg", $@"Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\vsgoto]
""URL Protocol""=""""
@= ""URL:vsgoto""

[HKEY_CLASSES_ROOT\vsgoto\shell]

[HKEY_CLASSES_ROOT\vsgoto\shell\open]

[HKEY_CLASSES_ROOT\vsgoto\shell\open\command]
@= ""\""{typeof(Program).Assembly.Location.Replace("\\", "\\\\")}\"" \""%1\""""
");
        }
    }
}
