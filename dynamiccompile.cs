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
    private const string SCRIPT_REPOSITORY_DIRECTORY = @"C:\Users\mikolaj\Documents\OledMinimiseWindows";
    private const string LOG_DIRECTORY = @"C:\Users\mikolaj\Documents\MinimizerLogs";
    private const string COMPILATION_LOGS_FILENAME = @"CompilationResults.log";
    private const string SCRIPT_FILENAME = @"minimisewindows.cs";
    private const string DLL_FILENAME = @"minimisewindows.dll";
    public static void Run(IntPtr windowHandle)
    {
        string scriptPath = Path.Combine(SCRIPT_REPOSITORY_DIRECTORY, SCRIPT_FILENAME);
        string assemblyPath = Path.Combine(SCRIPT_REPOSITORY_DIRECTORY, DLL_FILENAME);
        string compilerLogPath = Path.Combine(LOG_DIRECTORY, COMPILATION_LOGS_FILENAME);

        try
        {
            Assembly assembly = null;
            bool compileRequired = true;

            if (File.Exists(assemblyPath))
            {
                DateTime sourceTime = File.GetLastWriteTime(scriptPath);
                DateTime assemblyTime = File.GetLastWriteTime(assemblyPath);
                if (assemblyTime >= sourceTime) compileRequired = false;
            }

            if (compileRequired)
            {
                string code = File.ReadAllText(scriptPath);

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

                var compilationOptions = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
                    .WithSpecificDiagnosticOptions(new Dictionary<string, ReportDiagnostic>
                    {
                        { "CS1701", ReportDiagnostic.Suppress }
                    });

                // Code compilation
                var compilation = CSharpCompilation.Create(
                    "DynamicAssembly",
                    new[] { CSharpSyntaxTree.ParseText(code) },
                    references,
                    compilationOptions
                );

                // Save compilation result to DLL file
                using (var fs = new FileStream(assemblyPath, FileMode.Create, FileAccess.Write))
                {
                    var emitResult = compilation.Emit(fs);

                    // Store compilation result to log file
                    var compilationLogs = string.Join("\n", emitResult.Diagnostics.Select(d => d.ToString()));
                    File.WriteAllText(compilerLogPath, $"Compilation logs:\n{compilationLogs}");

                    if (!emitResult.Success)
                    {
                        MessageBox.Show($"Compilation errors:\n\n{compilationLogs}");
                        return;
                    }
                }
            } // compileRequired

            var assemblyBytes = File.ReadAllBytes(assemblyPath);
            assembly = Assembly.Load(assemblyBytes);

            // Find and invode metod Run from dynamically compiled code
            Type type = assembly.GetType("DisplayFusionFunction");
            MethodInfo method = type.GetMethod("Run", new[] { typeof(IntPtr) });
            method.Invoke(null, new object[] { windowHandle });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Exception:\n\n{ex}");
            File.WriteAllText(compilerLogPath, $"Exception: {ex}");
        }
    }
}
