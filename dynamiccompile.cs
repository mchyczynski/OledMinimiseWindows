using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System.Windows.Forms;

public static class DisplayFusionFunction
{
    public static void Run(IntPtr windowHandle)
    {
        string scriptPath = @"C:\Users\mikolaj\Documents\OledMinimiseWindows\minimisewindows.cs";

        try
        {
            // MessageBox.Show($"About to compile");
            string code = File.ReadAllText(scriptPath);

            // Pobranie domyślnych referencji .NET
            var references = new List<MetadataReference>
            {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location),       // System.Private.CoreLib.dll
                MetadataReference.CreateFromFile(typeof(Console).Assembly.Location),      // System.Console.dll
                MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Linq").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Drawing").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Drawing.Primitives").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Windows.Forms").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Collections").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Collections.Concurrent").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Collections.Specialized").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Text.RegularExpressions").Location),
                MetadataReference.CreateFromFile(@"C:\Program Files\DisplayFusion\DisplayFusion.dll"),
                MetadataReference.CreateFromFile(@"C:\Program Files\DisplayFusion\DisplayFusionScripting.dll")
            };

            // Konfiguracja kompilatora Roslyn
            var compilation = CSharpCompilation.Create(
                "DynamicAssembly",
                new[] { CSharpSyntaxTree.ParseText(code) },
                references,
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
                .WithSpecificDiagnosticOptions(new Dictionary<string, ReportDiagnostic>
                                                {
                                                    {"CS1701", ReportDiagnostic.Suppress}
                                                })
            );

            using (var ms = new MemoryStream())
            {
                var result = compilation.Emit(ms);
                // Pobranie błędów kompilacji
                var errors = string.Join("\n", result.Diagnostics.Select(d => d.ToString()));
                File.WriteAllText(@"C:\Users\mikolaj\Documents\MinimizerLogs\Błędy.log", $"Błędy kompilacji:\n{errors}");
                if (!result.Success)
                {
                    MessageBox.Show($"Compiler errors:\n\n{errors}");
                    return;
                }

                // Wczytaj skompilowany kod
                ms.Seek(0, SeekOrigin.Begin);
                var assembly = Assembly.Load(ms.ToArray());

                // Znajdź i uruchom metodę Run
                Type type = assembly.GetType("DisplayFusionFunction");
                MethodInfo method = type.GetMethod("Run", new[] { typeof(IntPtr) });
                method.Invoke(null, new object[] { windowHandle });
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Exception:\n\n{ex}");
            File.WriteAllText(@"C:\Users\mikolaj\Documents\MinimizerLogs\Błędy.log", $"Błąd: {ex}");
        }
    }
}