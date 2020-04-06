using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
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

        public static Dictionary<string, string> VSVersions = new Dictionary<string, string>()
        {
            {"2013", "12.0"},
            {"2015", "14.0"},
            {"2017", "15.0"},
            {"2019", "16.0"},
        };

        static void Main(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
                Console.WriteLine($"%{i}: \"{Escape(args[i])}\"");

            string file = "";
            int line = 1;

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
                uriFile = Uri.UnescapeDataString(uriFile);

                var m = PARSER.Match(uriFile);

                file = uriFile;
                line = 1;

                if(m.Success)
                {
                    file = m.Groups[1].Value;
                    line = int.Parse(m.Groups[2].Value);
                }

                foreach (DictionaryEntry env in Environment.GetEnvironmentVariables()) {
                    file = InjectSingleValue(file, env.Key.ToString(), env.Value);
                }



                OpenFile(file, line);
            } catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine($"file: {Escape(file)}");
                Console.WriteLine($"line: {line}");
                Console.ReadLine();
            }
        }

        public static void OpenFile(string file, int line)
        {
            EnvDTE80.DTE2 dte2;

            dte2 = FindVSInstance();

            if(dte2 is null)
                throw new Exception("No Visual Studio version found");
            
            dte2.MainWindow.Activate();
            EnvDTE.Window w = dte2.ItemOperations.OpenFile(file);
            ((EnvDTE.TextSelection)dte2.ActiveDocument.Selection).GotoLine(line, true);
        }

        public static EnvDTE80.DTE2 FindVSInstance()
        {
            try {
                return (EnvDTE80.DTE2)System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "VisualStudio.DTE");
            } catch (Exception ex) { }

            foreach (var vsVersion in VSVersions.Reverse()) {
                try
                {
                    return (EnvDTE80.DTE2) System.Runtime.InteropServices.Marshal.GetActiveObject(
                        "VisualStudio.DTE." + vsVersion.Value);
                }catch(Exception ex) { }
            }

            return null;
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

        /// <summary>
        /// Replaces all instances of a 'key' (e.g. {foo} or {foo:SomeFormat}) in a string with an optionally formatted value, and returns the result.
        /// </summary>
        /// <param key="formatString">The string containing the key; unformatted ({foo}), or formatted ({foo:SomeFormat})</param>
        /// <param key="key">The key key (foo)</param>
        /// <param key="replacementValue">The replacement value; if null is replaced with an empty string</param>
        /// <returns>The input string with any instances of the key replaced with the replacement value</returns>
        public static string InjectSingleValue(string formatString, string key, object replacementValue) {
            string result = formatString;
            //regex replacement of key with value, where the generic key format is:
            //Regex foo = new Regex("{(foo)(?:}|(?::(.[^}]*)}))");
            Regex attributeRegex = new Regex("{(" + key + ")(?:}|(?::(.[^}]*)}))", RegexOptions.IgnoreCase);  //for key = foo, matches {foo} and {foo:SomeFormat}

            //loop through matches, since each key may be used more than once (and with a different format string)
            foreach (Match m in attributeRegex.Matches(formatString)) {
                string replacement = m.ToString();
                if (m.Groups[2].Length > 0) //matched {foo:SomeFormat}
                {
                    //do a double string.Format - first to build the proper format string, and then to format the replacement value
                    string attributeFormatString = string.Format(CultureInfo.InvariantCulture, "{{0:{0}}}", m.Groups[2]);
                    replacement = string.Format(CultureInfo.CurrentCulture, attributeFormatString, replacementValue);
                } else //matched {foo}
                {
                    replacement = (replacementValue ?? string.Empty).ToString();
                }
                //perform replacements, one match at a time
                result = result.Replace(m.ToString(), replacement);  //attributeRegex.Replace(result, replacement, 1);
            }
            return result;

        }

        public static string Escape(string s) {
            if (s == null || s.Length == 0) {
                return "";
            }

            char c = '\0';
            int i;
            int len = s.Length;
            StringBuilder sb = new StringBuilder(len + 4);
            String t;

            for (i = 0; i < len; i += 1) {
                c = s[i];
                switch (c) {
                    case '\\':
                    case '"':
                        sb.Append('\\');
                        sb.Append(c);
                        break;
                    case '/':
                        sb.Append('\\');
                        sb.Append(c);
                        break;
                    case '\b':
                        sb.Append("\\b");
                        break;
                    case '\t':
                        sb.Append("\\t");
                        break;
                    case '\n':
                        sb.Append("\\n");
                        break;
                    case '\f':
                        sb.Append("\\f");
                        break;
                    case '\r':
                        sb.Append("\\r");
                        break;
                    default:
                        if (c < ' ') {
                            t = "000" + String.Format("X", c);
                            sb.Append("\\u" + t.Substring(t.Length - 4));
                        } else {
                            sb.Append(c);
                        }
                        break;
                }
            }
            return sb.ToString();
        }
    }
}
