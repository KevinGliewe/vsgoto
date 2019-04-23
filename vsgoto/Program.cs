using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vsgoto
{
    class Program
    {
        public static readonly string CODE = "code.cmd";
        public static readonly string SCHEMA = "vsgoto:";

        static void Main(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
                Console.WriteLine($"%{i}: " + args[i]);

            try
            {
                var code = GetFullPath(CODE);

                if (code is null)
                    throw new Exception($"Can't find '{CODE}' in PATH");
                else
                    Console.WriteLine($"Found code Command: '{code}'");

                if (args.Length < 1)
                {
                    ExportRegFile();
                    Console.WriteLine("Reg-File Generated...");
                    throw new Exception("Missing Arguments!");
                }

                if (!args[0].StartsWith(SCHEMA))
                    throw new Exception("Wrong Schema!");


                var uriFile = args[0].Substring(SCHEMA.Length, args[0].Length - SCHEMA.Length);
                var vsargs = "-g \"" + uriFile + "\"";

                Console.WriteLine("ARGS: " + vsargs);

                var result = ExFile(code, vsargs);
                Console.WriteLine(result);
            } catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
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

        public static string ExFile(string file, string args)
        {
            var process = new Process()
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = file,
                    Arguments = args,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                }
            };
            process.Start();
            string result = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            return result.Trim();
        }

        public static string GetFullPath(string fileName)
        {
            if (File.Exists(fileName))
                return Path.GetFullPath(fileName);

            var values = Environment.GetEnvironmentVariable("PATH");
            foreach (var path in values.Split(';'))
            {
                var fullPath = Path.Combine(path, fileName);
                if (File.Exists(fullPath))
                    return fullPath;
            }
            return null;
        }
    }
}
