// Add assembly to references:
// C:\Program Files\DisplayFusion\Microsoft.CodeAnalysis.dll
// C:\Program Files\DisplayFusion\Microsoft.CodeAnalysis.CSharp.dll

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System.Windows.Forms;
using System.Runtime.Loader;

public class CollectibleALC : AssemblyLoadContext
{
    public CollectibleALC() : base(isCollectible: true) { }
    protected override Assembly Load(AssemblyName assemblyName) => null;
}

public static class DisplayFusionFunction
{
    private const string SCRIPT_REPOSITORY_DIRECTORY = @"C:\Users\mikolaj\Documents\OledMinimiseWindows";
    private const string LOG_DIRECTORY = @"C:\Users\mikolaj\Documents\MinimizerLogs";
    private const string COMPILATION_LOGS_FILENAME = @"CompilationResults.log";
    private const string SCRIPT_FILENAME = @"minimisewindows.cs";
    private const string DLL_FILENAME = @"minimisewindows.dll";

    private static Assembly cachedAssembly = null;
    private static CollectibleALC currentALC = null;
    private static readonly object assemblyLock = new object();

    public static void Run(IntPtr windowHandle)
    {
        string scriptPath = Path.Combine(SCRIPT_REPOSITORY_DIRECTORY, SCRIPT_FILENAME);
        string assemblyPath = Path.Combine(SCRIPT_REPOSITORY_DIRECTORY, DLL_FILENAME);
        string compilerLogPath = Path.Combine(LOG_DIRECTORY, COMPILATION_LOGS_FILENAME);

        try
        {
            Assembly assembly = null;
            bool compileRequired = true;

            lock (assemblyLock)
            {
                // Skip compilation if DLL exists and source file was not modified in meantime
                if (File.Exists(assemblyPath))
                {
                    DateTime sourceTime = File.GetLastWriteTime(scriptPath);
                    DateTime assemblyTime = File.GetLastWriteTime(assemblyPath);
                    if (assemblyTime >= sourceTime)
                        compileRequired = false;
                }

                if (compileRequired)
                {
                    // If context was loaded unload it before compiling new DLL
                    if (currentALC != null)
                    {
                        currentALC.Unload();
                        currentALC = null;
                        cachedAssembly = null;
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }


                    var references = GetReferences();
                    if (!CompileAssembly(scriptPath, assemblyPath, compilerLogPath))
                        return;

                } // End of compileRequired
                  // else MessageBox.Show($"Compilation skipped");


                // Use cached assembly if available and new wasn't just recompiled
                if (!compileRequired && cachedAssembly != null)
                {
                    // MessageBox.Show($"Loading assembly skipped");
                    assembly = cachedAssembly;
                }
                else // There is no cached assembly or it was just compiled
                {
                    currentALC = new CollectibleALC();

                    // Load assembly from file into byte array and load it from it so that file is not locked
                    byte[] assemblyBytes = File.ReadAllBytes(assemblyPath);
                    assembly = currentALC.LoadFromStream(new MemoryStream(assemblyBytes));

                    // Cache newly loaded assembly
                    cachedAssembly = assembly;
                }


                // Find and invoke metod Run from dynamically compiled code
                Type type = assembly.GetType("DisplayFusionFunction");
                MethodInfo method = type.GetMethod("Run", new[] { typeof(IntPtr) });
                method.Invoke(null, new object[] { windowHandle });
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Exception:\n\n{ex}");
            File.WriteAllText(compilerLogPath, $"Exception: {ex}");
        }
    } // Run

    private static bool CompileAssembly(string scriptPath, string assemblyPath, string compilerLogPath)
    {
        string code = File.ReadAllText(scriptPath);
        var references = GetReferences();
        var compilationOptions = new CSharpCompilationOptions(
                 OutputKind.DynamicallyLinkedLibrary,
                 warningLevel: 4, // Max warning level
                 generalDiagnosticOption: ReportDiagnostic.Warn // Treat warnings as warnings
        )
        .WithSpecificDiagnosticOptions(new Dictionary<string, ReportDiagnostic>
        {
            { "CS1701", ReportDiagnostic.Suppress }, // Rectangle older version
        });

        var compilation = CSharpCompilation.Create(
            "DynamicAssembly",
            new[] { CSharpSyntaxTree.ParseText(code) },
            references,
            compilationOptions
        );

        using (var fs = new FileStream(assemblyPath, FileMode.Create, FileAccess.Write))
        {
            var emitResult = compilation.Emit(fs);
            var compilationLogs = string.Join("\n", emitResult.Diagnostics.Select(d => d.ToString()));
            File.WriteAllText(compilerLogPath, $"Compilation logs:\n{compilationLogs}");

            if (!emitResult.Success)
            {
                MessageBox.Show($"Compilation errors:\n\n{compilationLogs}");
                return false;
            }
        }
        return true;
    }

    private static List<MetadataReference> GetReferences()
    {
        return new List<MetadataReference>
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
    }
}
