using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using Johan.Extensions;

namespace GenerateVbaEnumHelpers
{
    class Program
    {
        private const string _title = "Generate VBA enum helpers";
        private const string _settingsFile = "dllPaths.xml";

        private static string[] _dllPaths = null;

        static void Main(string[] args)
        {
            Console.Title = _title;
            ConsoleEx.PrintTitle(_title);

            try
            {
                LoadDllPaths();

                foreach (var item in _dllPaths)
                    foreach (Type type in Assembly.LoadFrom(item).GetTypes().Where(t => t.IsEnum && t.IsPublic))
                        GenerateModule(type);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            Console.WriteLine();
            Console.WriteLine("Done executing!");
            Console.ReadLine();
        }

        private static void LoadDllPaths()
        {
            if (File.Exists(_settingsFile))
                _dllPaths = XmlSerialization.DeserializeFromFile<string[]>(_settingsFile);

            else //standard settings
            {
                _dllPaths = new string[] {
                    @"C:\Windows\assembly\GAC_MSIL\Office\14.0.0.0__71e9bce111e9429c\Office.dll",
                    @"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook\14.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Outlook.dll",
                    @"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\14.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll",
                    @"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\14.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll",
                    @"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Publisher\14.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Publisher.dll",
                    @"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\14.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
                };
                XmlSerialization.Serialize(_dllPaths, _settingsFile);
            }
        }

        private static void GenerateModule(Type enumType)
        {
            Directory.CreateDirectory("output");
            var filename = Path.Combine("output", enumType.Name + ".bas");

            try
            {
                if (!enumType.IsEnum)
                    throw new Exception(enumType + " is not an enum");

                var vbaCode = GenerateVba(enumType);


                File.WriteAllText(filename, vbaCode);

                Console.WriteLine("- Generated {0}", filename);
            }
            catch {
                Console.WriteLine("- Failed to generate {0}", filename);
            }
        }

        private static string GenerateVba(Type type)
        {
            var code = new StringBuilder();

            code.AppendFormat("Attribute VB_Name = \"w{0}\"{1}", type.Name, Environment.NewLine);

            BuildFromString(type, code);

            code.AppendLine();

            AddToString(type, code);

            return code.ToString();
        }

        private static StringBuilder AddToString(Type type, StringBuilder code)
        {
            var functionName = type.Name + "ToString";
            code.AppendFormat("Function {0}(value As {1}) As String{2}", functionName, type.Name, Environment.NewLine);

            code.AppendFormat(@"    Select Case value{1}", functionName, Environment.NewLine);

            var names = Enum.GetNames(type);

            foreach (var name in names)
            {
                code.AppendFormat(@"        Case {0}: {1} = ""{0}""{2}", name, functionName, Environment.NewLine);
            }

            code.AppendLine("    End Select");
            code.AppendLine("End Function");

            return code;
        }

        private static StringBuilder BuildFromString(Type type, StringBuilder code)
        {
            var functionName = type.Name + "FromString";
            code.AppendFormat("Function {0}(value As String) As {1}{2}", functionName, type.Name, Environment.NewLine);

            code.AppendFormat(@"    If IsNumeric(value) Then
        {0} = CInt(value)
        Exit Function
    End If{1}{1}    Select Case value{1}", functionName, Environment.NewLine);

            var names = Enum.GetNames(type);

            foreach (var name in names)
            {
                code.AppendFormat(@"        Case ""{0}"": {1} = {0}{2}", name, functionName, Environment.NewLine);
            }

            code.AppendLine("    End Select");
            code.AppendLine("End Function");

            return code;
        }
    }
}
