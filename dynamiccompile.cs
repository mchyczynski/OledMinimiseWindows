// Add assembly references:
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
    // Instead of a single file, we compile only the files specified in the list below.
    private static readonly string[] SourceFileNames = new string[]
    {
        "minimisewindows.cs",
        "log.cs",
        "windowsutils.cs"
    };
    private const string DLL_FILENAME = @"minimisewindows.dll";

    private static Assembly cachedAssembly = null;
    private static CollectibleALC currentALC = null;
    private static readonly object assemblyLock = new object();

    public static void Run(IntPtr windowHandle)
    {
        string assemblyPath = Path.Combine(SCRIPT_REPOSITORY_DIRECTORY, DLL_FILENAME);
        string compilerLogPath = Path.Combine(LOG_DIRECTORY, COMPILATION_LOGS_FILENAME);

        try
        {
            Assembly assembly = null;
            bool compileRequired = true;

            // Check if compilation is required by comparing the last write times of the source files and the DLL.
            DateTime sourceTime = SourceFileNames
                .Select(file => File.GetLastWriteTime(Path.Combine(SCRIPT_REPOSITORY_DIRECTORY, file)))
                .Max();
            DateTime assemblyTime = File.Exists(assemblyPath) ? File.GetLastWriteTime(assemblyPath) : DateTime.MinValue;
            if (assemblyTime >= sourceTime)
                compileRequired = false;

            if (compileRequired)
            {
                // If an AssemblyLoadContext was already loaded, unload it before compiling the new DLL.
                if (currentALC != null)
                {
                    currentALC.Unload();
                    currentALC = null;
                    cachedAssembly = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                if (!CompileAssembly(SCRIPT_REPOSITORY_DIRECTORY, assemblyPath, compilerLogPath))
                    return;
            }
            // else MessageBox.Show("Compilation skipped");

            lock (assemblyLock)
            {
                // Use the cached assembly if available and if compilation was not performed.
                if (!compileRequired && cachedAssembly != null)
                {
                    assembly = cachedAssembly;
                }
                else // Otherwise, load the assembly from the file (via a byte array to avoid locking the file)
                {
                    currentALC = new CollectibleALC();
                    byte[] assemblyBytes = File.ReadAllBytes(assemblyPath);
                    assembly = currentALC.LoadFromStream(new MemoryStream(assemblyBytes));
                    cachedAssembly = assembly;
                }
            }

            // Find the class 'DisplayFusionFunction' in the compiled assembly.
            Type type = assembly.GetType("DisplayFusionFunction");
            if (type == null)
            {
                string message = "Class 'DisplayFusionFunction' not found in compiled code.";
                File.AppendAllText(compilerLogPath, message);
                MessageBox.Show(message);
                return;
            }

            // Find the 'Run' method with the signature (IntPtr).
            MethodInfo method = type.GetMethod("Run", new[] { typeof(IntPtr) });
            if (method == null)
            {
                string message = "Method 'Run(IntPtr)' not found in compiled code.";
                File.AppendAllText(compilerLogPath, message);
                MessageBox.Show(message);
                return;
            }

            try
            {
                method.Invoke(null, new object[] { windowHandle });
            }
            catch (TargetInvocationException tie)
            {
                string message = $"Error during invocation of compiled code:\n\n{tie.InnerException}";
                File.AppendAllText(compilerLogPath, message);
                MessageBox.Show(message);
            }
            catch (Exception ex)
            {
                string message = $"Unknown error during execution of compiled code:\n\n{ex}";
                File.AppendAllText(compilerLogPath, message);
                MessageBox.Show(message);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Exception:\n\n{ex}");
            File.WriteAllText(compilerLogPath, $"Exception: {ex}");
        }
    } // Run

    private static bool CompileAssembly(string sourceDirectory, string assemblyPath, string compilerLogPath)
    {
        // Compile only the files specified in the SourceFileNames list.
        List<SyntaxTree> syntaxTrees = new List<SyntaxTree>();
        foreach (var file in SourceFileNames)
        {
            string fullPath = Path.Combine(sourceDirectory, file);
            try
            {
                string code = File.ReadAllText(fullPath);
                syntaxTrees.Add(CSharpSyntaxTree.ParseText(code, path: fullPath));
            }
            catch (Exception ex)
            {
                string message = $"Error reading file '{fullPath}':\n\n{ex}";
                File.AppendAllText(compilerLogPath, message);
                MessageBox.Show(message);
                return false;
            }
        }

        var references = GetReferences();

        var compilationOptions = new CSharpCompilationOptions(
            OutputKind.DynamicallyLinkedLibrary,
            warningLevel: 4, // Maximum warning level
            generalDiagnosticOption: ReportDiagnostic.Warn // Treat warnings as warnings
        )
        .WithSpecificDiagnosticOptions(new Dictionary<string, ReportDiagnostic>
        {
            { "CS1701", ReportDiagnostic.Suppress }, // Rectangle older version
        });

        var compilation = CSharpCompilation.Create(
            "DynamicAssembly",
            syntaxTrees,
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
