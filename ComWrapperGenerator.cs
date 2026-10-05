// ======================================================================
// ComWrapperGenerator - Full Extended Version
//
// Output: a complete Class Library project ready to build
//   <output>/ 
//     GeneratedComWrapper.csproj
//     ComWrapper.cs
//     build.bat / build.sh (optional helpers)
//
// Features:
//   - Parameterless & parameterized constructors (Factory Create)
//   - EnsureInitialized protection
//   - List<T>, IEnumerable<T>, IList<T>, ICollection<T>, arrays <-> object[]
//   - out / ref parameters
//   - Events with arguments (including complex objects)
//   - Complex objects from the same assembly (recursive wrapping)
//   - Static methods & static properties
//   - Method overload renaming
//   - Indexers as GetItem/SetItem (with overload renaming)
// ======================================================================

using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;

if (args.Length < 1)
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  ComWrapperGenerator <path-to-dll> [output-folder] [namespace] [target-framework]");
    Console.WriteLine();
    Console.WriteLine("Arguments:");
    Console.WriteLine("  path-to-dll       Source managed DLL (required)");
    Console.WriteLine("  output-folder     Where to write the project (default: <dll-dir>/ComWrapper)");
    Console.WriteLine("  namespace         Wrapper namespace (default: GeneratedComWrapper)");
    Console.WriteLine("  target-framework  net48 | net8.0-windows (default: net48)");
    Console.WriteLine();
    Console.WriteLine("Example:");
    Console.WriteLine("  ComWrapperGenerator.exe \"C:\\Libs\\MyLib.dll\" \"C:\\Output\" \"MyCompany.Wrapper\" net48");
    return 1;
}

string dllPath = Path.GetFullPath(args[0]);
string outputFolder = args.Length > 1 ? Path.GetFullPath(args[1]) : Path.Combine(Path.GetDirectoryName(dllPath)!, "ComWrapper");
string wrapperNamespace = args.Length > 2 ? args[2] : "GeneratedComWrapper";
string tfm = args.Length > 3 ? args[3] : "net48";

if (!File.Exists(dllPath))
{
    Console.WriteLine($"File not found: {dllPath}");
    return 1;
}

bool isNetFramework = tfm.StartsWith("net4", StringComparison.OrdinalIgnoreCase);
// Assembly/project name from namespace argument (e.g. MyCompany.Wrapper)
string projectName = string.Join(".", wrapperNamespace
    .Split(new[] { '.', '+', ' ' }, StringSplitOptions.RemoveEmptyEntries)
    .Select(p => new string(p.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray()))
    .Where(p => p.Length > 0));
if (string.IsNullOrEmpty(projectName))
    projectName = "GeneratedComWrapper";


Directory.CreateDirectory(outputFolder);

string dllDir = Path.GetDirectoryName(dllPath)!;

// Collect probe directories for dependency resolution / copy
var probeDirs = new List<string> { dllDir, AppContext.BaseDirectory };
string? nugetPackages = GetNuGetPackagesRoot();
if (nugetPackages != null)
    probeDirs.Add(nugetPackages);

// Shared map: assembly simple name -> full path of resolved file
var resolvedDeps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

ResolveEventHandler resolver = (sender, args) =>
{
    try
    {
        var an = new AssemblyName(args.Name);
        string simple = an.Name!;
        if (resolvedDeps.TryGetValue(simple, out var known) && File.Exists(known))
            return Assembly.LoadFrom(known);

        string? found = FindAssemblyFile(simple, an.Version, probeDirs);
        if (found != null)
        {
            resolvedDeps[simple] = found;
            return Assembly.LoadFrom(found);
        }
    }
    catch { /* ignore */ }
    return null;
};
AppDomain.CurrentDomain.AssemblyResolve += resolver;

Console.WriteLine($"Loading: {dllPath}");
Console.WriteLine($"Dependency probe dirs:");
foreach (var d in probeDirs.Take(5))
    Console.WriteLine($"  - {d}");
if (probeDirs.Count > 5)
    Console.WriteLine($"  - ... and {probeDirs.Count - 5} more (NuGet subfolders scanned on demand)");

Assembly asm;
try
{
    asm = Assembly.LoadFrom(dllPath);
}
catch (Exception ex)
{
    Console.WriteLine($"Failed to load assembly: {ex.Message}");
    return 1;
}

// Proactively resolve & copy referenced assemblies into output folder (and beside source if missing)
Console.WriteLine("Resolving dependencies...");
var depsToCopy = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
CollectAndResolveDependencies(asm, dllDir, probeDirs, resolvedDeps, depsToCopy, depth: 0);

foreach (var depPath in depsToCopy.OrderBy(p => p))
{
    string name = Path.GetFileName(depPath);
    // Copy next to source DLL if missing there
    string besideSource = Path.Combine(dllDir, name);
    if (!File.Exists(besideSource))
    {
        try
        {
            File.Copy(depPath, besideSource, overwrite: false);
            Console.WriteLine($"  [COPY] -> source dir: {name}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  [WARN] Cannot copy {name} to source dir: {ex.Message}");
        }
    }
    // Always ensure in output project folder
    string besideOutput = Path.Combine(outputFolder, name);
    if (!string.Equals(Path.GetFullPath(depPath), Path.GetFullPath(besideOutput), StringComparison.OrdinalIgnoreCase))
    {
        try
        {
            File.Copy(depPath, besideOutput, overwrite: true);
            Console.WriteLine($"  [COPY] -> output: {name}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  [WARN] Cannot copy {name} to output: {ex.Message}");
        }
    }
}

Dictionary<string, Type> allPublicTypes;
try
{
    allPublicTypes = asm.GetExportedTypes()
        .Where(t => t.IsClass && !t.IsAbstract && t.IsPublic && !t.IsGenericTypeDefinition)
        .ToDictionary(t => t.FullName!, t => t);
}
catch (ReflectionTypeLoadException ex)
{
    Console.WriteLine("Warning: some types failed to load:");
    foreach (var le in ex.LoaderExceptions ?? Array.Empty<Exception>())
        if (le != null) Console.WriteLine($"  - {le.Message}");
    allPublicTypes = (ex.Types ?? Array.Empty<Type>())
        .Where(t => t != null && t.IsClass && !t.IsAbstract && t.IsPublic && !t.IsGenericTypeDefinition)
        .ToDictionary(t => t!.FullName!, t => t!);
}

Console.WriteLine($"Public non-generic classes: {allPublicTypes.Count}");

// Heuristic: if assembly references System.Runtime 5+, warn about net4x TFM
try
{
    var runtimeRef = asm.GetReferencedAssemblies()
        .FirstOrDefault(a => a.Name == "System.Runtime");
    if (runtimeRef?.Version != null && runtimeRef.Version.Major >= 5
        && tfm.StartsWith("net4", StringComparison.OrdinalIgnoreCase))
    {
        Console.WriteLine();
        Console.WriteLine("WARNING: Source assembly appears to target modern .NET (System.Runtime "
            + runtimeRef.Version + ").");
        Console.WriteLine($"         TFM '{tfm}' will likely fail with CS1705.");
        Console.WriteLine("         Use e.g. net6.0-windows / net8.0-windows / net10.0-windows instead of net48/net481.");
        Console.WriteLine();
    }
}
catch { /* ignore */ }

var typesToWrap = new HashSet<Type>();
foreach (var t in allPublicTypes.Values)
{
    try
    {
        CollectTypesToWrap(t, typesToWrap, allPublicTypes);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  [WARN] Skipping type {t.FullName}: {ex.Message}");
    }
}

var orderedTypes = typesToWrap.OrderBy(t => t.Name).ToList();
Console.WriteLine($"Wrappers to generate: {orderedTypes.Count}");

// ---------- Generate ComWrapper.cs ----------
var sb = new StringBuilder();
sb.AppendLine("// ======================================================================");
sb.AppendLine("// Auto-generated COM Wrapper");
sb.AppendLine($"// Source        : {Path.GetFileName(dllPath)}");
sb.AppendLine($"// Generated     : {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
sb.AppendLine("// Features      : Factory, Collections, out/ref, Events+args,");
sb.AppendLine("//                 Complex objects, Static members, Overload renaming, Indexers");
sb.AppendLine("// Generator     : ComWrapperGenerator (full extended)");
sb.AppendLine("// ======================================================================");
sb.AppendLine();
sb.AppendLine("#nullable enable");
sb.AppendLine("using System;");
sb.AppendLine("using System.Collections.Generic;");
sb.AppendLine("using System.Linq;");
sb.AppendLine("using System.Runtime.InteropServices;");
sb.AppendLine();
sb.AppendLine($"namespace {wrapperNamespace}");
sb.AppendLine("{");

var wrapperMap = orderedTypes.ToDictionary(t => t.FullName!, t => "I" + t.Name);

int generated = 0;
foreach (var type in orderedTypes)
{
    if (GenerateWrapper(sb, type, wrapperNamespace, wrapperMap, allPublicTypes))
        generated++;
}

sb.AppendLine("}");

string csFile = Path.Combine(outputFolder, "ComWrapper.cs");
File.WriteAllText(csFile, sb.ToString(), Encoding.UTF8);

// ---------- Generate .csproj ----------
string csprojPath = Path.Combine(outputFolder, $"{projectName}.csproj");
WriteCsproj(csprojPath, projectName, tfm, isNetFramework, dllPath, wrapperNamespace, outputFolder);

// ---------- Generate build helpers ----------
WriteBuildBat(outputFolder, projectName, isNetFramework);
WriteBuildSh(outputFolder, projectName);
WriteReadmeSnippet(outputFolder, projectName, wrapperNamespace, dllPath, tfm);

Console.WriteLine();
Console.WriteLine($"Done. Generated wrappers: {generated}");
Console.WriteLine($"Project folder: {outputFolder}");
Console.WriteLine($"  - {projectName}.csproj");
Console.WriteLine($"  - ComWrapper.cs");
Console.WriteLine($"  - build.bat / build.sh");
Console.WriteLine();
Console.WriteLine("Build:");
Console.WriteLine($"  cd \"{outputFolder}\"");
Console.WriteLine("  dotnet build -c Release");
if (isNetFramework)
{
    Console.WriteLine();
    Console.WriteLine("Register (admin, matching Office bitness):");
    Console.WriteLine($"  regasm bin\\Release\\{tfm}\\{projectName}.dll /codebase /tlb");
}
else
{
    Console.WriteLine();
    Console.WriteLine("Modern .NET COM host:");
    Console.WriteLine($"  regsvr32 bin\\Release\\{tfm}\\{projectName}.comhost.dll");
}
Console.WriteLine($"Output assembly: {projectName}.dll");
return 0;

// =====================================================================
// Project file generation
// =====================================================================

static void WriteCsproj(string path, string projectName, string tfm, bool isNetFramework, string sourceDllPath, string wrapperNamespace, string outputFolder)
{
    // Copy source DLL next to project for reliable relative reference
    string localDllName = Path.GetFileName(sourceDllPath);
    string localDllPath = Path.Combine(Path.GetDirectoryName(path)!, localDllName);
    if (!string.Equals(Path.GetFullPath(sourceDllPath), Path.GetFullPath(localDllPath), StringComparison.OrdinalIgnoreCase))
    {
        try { File.Copy(sourceDllPath, localDllPath, overwrite: true); }
        catch { /* best effort */ }
    }

    var sb = new StringBuilder();
    sb.AppendLine("<Project Sdk=\"Microsoft.NET.Sdk\">");
    sb.AppendLine();
    sb.AppendLine("  <PropertyGroup>");
    sb.AppendLine($"    <TargetFramework>{tfm}</TargetFramework>");
    sb.AppendLine("    <ImplicitUsings>disable</ImplicitUsings>");
    sb.AppendLine("    <Nullable>enable</Nullable>");
    sb.AppendLine("    <LangVersion>latest</LangVersion>");
    sb.AppendLine("    <OutputType>Library</OutputType>");
    sb.AppendLine($"    <AssemblyName>{projectName}</AssemblyName>");
    sb.AppendLine($"    <RootNamespace>{wrapperNamespace}</RootNamespace>");
    sb.AppendLine("    <GenerateAssemblyInfo>true</GenerateAssemblyInfo>");
    sb.AppendLine("    <!-- SDK already includes **/*.cs — do not list Compile explicitly -->");
    sb.AppendLine("    <EnableDefaultCompileItems>true</EnableDefaultCompileItems>");
    sb.AppendLine();
    sb.AppendLine("    <!-- COM Interop -->");
    sb.AppendLine("    <EnableComHosting>true</EnableComHosting>");
    if (isNetFramework)
    {
        sb.AppendLine("    <RegisterForComInterop>false</RegisterForComInterop>");
        sb.AppendLine("    <ComVisible>true</ComVisible>");
        sb.AppendLine("    <!-- Register with regasm after build; auto-register is fragile in CI -->");
    }
    else
    {
        sb.AppendLine("    <ComVisible>true</ComVisible>");
    }
    sb.AppendLine("    <!-- Avoid MSB3277 noise from mixed dependency versions -->");
    sb.AppendLine("    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>");
    sb.AppendLine("    <GenerateBindingRedirectsOutputType>true</GenerateBindingRedirectsOutputType>");
    sb.AppendLine("  </PropertyGroup>");
    sb.AppendLine();
    sb.AppendLine("  <ItemGroup>");
    sb.AppendLine($"    <Reference Include=\"{Path.GetFileNameWithoutExtension(localDllName)}\">");
    sb.AppendLine($"      <HintPath>{localDllName}</HintPath>");
    sb.AppendLine("      <Private>true</Private>");
    sb.AppendLine("    </Reference>");

    // Reference other managed deps that may define types used in public API (not framework facades)
    var depDlls = Directory.GetFiles(outputFolder, "*.dll")
        .Select(f => Path.GetFileName(f)!)
        .Where(n => !n.Equals(localDllName, StringComparison.OrdinalIgnoreCase) &&
                    !n.Equals(projectName + ".dll", StringComparison.OrdinalIgnoreCase))
        .OrderBy(n => n)
        .ToList();

    foreach (var dep in depDlls)
    {
        string simple = Path.GetFileNameWithoutExtension(dep)!;
        if (IsFrameworkAssembly(simple))
            continue;
        // Skip facades / packages that cause MSB3277 version conflicts
        if (simple.StartsWith("Microsoft.Win32", StringComparison.OrdinalIgnoreCase))
            continue;
        if (simple.Equals("Newtonsoft.Json", StringComparison.OrdinalIgnoreCase))
            continue; // copy to output only; KkmDriver already brings the right version
        sb.AppendLine($"    <Reference Include=\"{simple}\">");
        sb.AppendLine($"      <HintPath>{dep}</HintPath>");
        sb.AppendLine("      <Private>true</Private>");
        sb.AppendLine("    </Reference>");
    }
    sb.AppendLine("  </ItemGroup>");
    sb.AppendLine();

    if (depDlls.Count > 0)
    {
        sb.AppendLine("  <!-- Ensure all deps land in output folder -->");
        sb.AppendLine("  <ItemGroup>");
        foreach (var dep in depDlls)
        {
            sb.AppendLine($"    <None Include=\"{dep}\">");
            sb.AppendLine("      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>");
            sb.AppendLine("    </None>");
        }
        sb.AppendLine("  </ItemGroup>");
        sb.AppendLine();
    }
    sb.AppendLine("</Project>");

    File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
}

static void WriteBuildBat(string folder, string projectName, bool isNetFramework)
{
    var sb = new StringBuilder();
    sb.AppendLine("@echo off");
    sb.AppendLine("setlocal");
    sb.AppendLine("cd /d \"%~dp0\"");
    sb.AppendLine();
    sb.AppendLine("echo Building...");
    sb.AppendLine("dotnet build -c Release");
    sb.AppendLine("if errorlevel 1 exit /b 1");
    sb.AppendLine();
    if (isNetFramework)
    {
        sb.AppendLine("echo.");
        sb.AppendLine("echo To register for COM (run as Administrator, match Office bitness):");
        sb.AppendLine($"echo   regasm bin\\Release\\net48\\{projectName}.dll /codebase /tlb");
        sb.AppendLine("echo.");
        sb.AppendLine("echo 32-bit Office:");
        sb.AppendLine($"echo   %%windir%%\\Microsoft.NET\\Framework\\v4.0.30319\\regasm.exe bin\\Release\\net48\\{projectName}.dll /codebase /tlb");
        sb.AppendLine("echo 64-bit Office:");
        sb.AppendLine($"echo   %%windir%%\\Microsoft.NET\\Framework64\\v4.0.30319\\regasm.exe bin\\Release\\net48\\{projectName}.dll /codebase /tlb");
    }
    else
    {
        sb.AppendLine("echo.");
        sb.AppendLine($"echo For .NET 5+ COM: regsvr32 ...\\{projectName}.comhost.dll");
        sb.AppendLine("echo Generate TLB with dscom if needed.");
    }
    sb.AppendLine("endlocal");
    File.WriteAllText(Path.Combine(folder, "build.bat"), sb.ToString(), Encoding.UTF8);
}

static void WriteBuildSh(string folder, string projectName)
{
    var sb = new StringBuilder();
    sb.AppendLine("#!/usr/bin/env bash");
    sb.AppendLine("set -e");
    sb.AppendLine("cd \"$(dirname \"$0\")\"");
    sb.AppendLine("echo Building...");
    sb.AppendLine("dotnet build -c Release");
    sb.AppendLine("echo Done.");
    File.WriteAllText(Path.Combine(folder, "build.sh"), sb.ToString(), Encoding.UTF8);
}

static void WriteReadmeSnippet(string folder, string projectName, string ns, string sourceDll, string tfm)
{
    var sb = new StringBuilder();
    sb.AppendLine("# Generated COM Wrapper Project");
    sb.AppendLine();
    sb.AppendLine($"- Source DLL : `{Path.GetFileName(sourceDll)}`");
    sb.AppendLine($"- Namespace  : `{ns}`");
    sb.AppendLine($"- Assembly   : `{projectName}.dll`");
    sb.AppendLine($"- TFM        : `{tfm}`");
    sb.AppendLine();
    sb.AppendLine("## Build");
    sb.AppendLine();
    sb.AppendLine("```bash");
    sb.AppendLine("dotnet build -c Release");
    sb.AppendLine("# or: build.bat / ./build.sh");
    sb.AppendLine("```");
    sb.AppendLine();
    sb.AppendLine("## Register (Windows, .NET Framework / net48)");
    sb.AppendLine();
    sb.AppendLine("```bat");
    sb.AppendLine($"regasm bin\\Release\\net48\\{projectName}.dll /codebase /tlb");
    sb.AppendLine("```");
    sb.AppendLine();
    sb.AppendLine("Match regasm bitness with Office (Framework vs Framework64).");
    sb.AppendLine();
    sb.AppendLine("## VBA");
    sb.AppendLine();
    sb.AppendLine("Tools → References → select the generated type library, then:");
    sb.AppendLine();
    sb.AppendLine("```vb");
    sb.AppendLine($"Dim obj As {ns}.IYourClass");
    sb.AppendLine($"Set obj = New {ns}.YourClass");
    sb.AppendLine("' if constructor had parameters:");
    sb.AppendLine("Set obj = obj.Create(...)");
    sb.AppendLine("```");
    File.WriteAllText(Path.Combine(folder, "README_GENERATED.md"), sb.ToString(), Encoding.UTF8);
}

// =====================================================================
// Dependency resolution & copy
// =====================================================================

static string? GetNuGetPackagesRoot()
{
    string? home = Environment.GetEnvironmentVariable("NUGET_PACKAGES");
    if (!string.IsNullOrEmpty(home) && Directory.Exists(home))
        return home;

    string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    string defaultCache = Path.Combine(userProfile, ".nuget", "packages");
    if (Directory.Exists(defaultCache))
        return defaultCache;

    return null;
}

static string? FindAssemblyFile(string simpleName, Version? version, List<string> probeDirs)
{
    string fileName = simpleName + ".dll";

    // 1) Direct probe directories
    foreach (var dir in probeDirs)
    {
        if (dir.EndsWith("packages", StringComparison.OrdinalIgnoreCase))
            continue; // NuGet root handled below

        string candidate = Path.Combine(dir, fileName);
        if (File.Exists(candidate))
            return candidate;
    }

    // 2) NuGet global packages cache: packages/<id>/<version>/lib/**/<name>.dll
    string? nugetRoot = probeDirs.FirstOrDefault(d =>
        d.EndsWith("packages", StringComparison.OrdinalIgnoreCase) ||
        d.Contains($"{Path.DirectorySeparatorChar}.nuget{Path.DirectorySeparatorChar}packages"));
    if (nugetRoot == null)
        nugetRoot = GetNuGetPackagesRoot();

    if (nugetRoot != null && Directory.Exists(nugetRoot))
    {
        // Package folder names are lowercase
        string pkgId = simpleName.ToLowerInvariant();
        string pkgDir = Path.Combine(nugetRoot, pkgId);
        if (Directory.Exists(pkgDir))
        {
            // Prefer version match, else highest version
            var versionDirs = Directory.GetDirectories(pkgDir)
                .Select(d => new DirectoryInfo(d))
                .OrderByDescending(d => d.Name)
                .ToList();

            if (version != null)
            {
                var exact = versionDirs.FirstOrDefault(d =>
                    d.Name.StartsWith(version.ToString(), StringComparison.OrdinalIgnoreCase) ||
                    d.Name.Equals(version.ToString(), StringComparison.OrdinalIgnoreCase));
                if (exact != null)
                {
                    string? hit = FindDllUnderLib(exact.FullName, fileName);
                    if (hit != null) return hit;
                }
            }

            foreach (var vd in versionDirs)
            {
                string? hit = FindDllUnderLib(vd.FullName, fileName);
                if (hit != null) return hit;
            }
        }

        // Sometimes package id differs slightly (e.g. still system.drawing.common)
        // Already using simpleName.ToLowerInvariant which matches NuGet id for System.Drawing.Common
    }

    // 3) Runtime shared framework folders (best-effort)
    string? dotnetRoot = Environment.GetEnvironmentVariable("DOTNET_ROOT")
        ?? (OperatingSystem.IsWindows()
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet")
            : "/usr/share/dotnet");
    if (Directory.Exists(dotnetRoot))
    {
        foreach (var pattern in new[] { "shared/Microsoft.NETCore.App", "shared/Microsoft.WindowsDesktop.App", "shared/Microsoft.AspNetCore.App" })
        {
            string shared = Path.Combine(dotnetRoot, pattern.Replace('/', Path.DirectorySeparatorChar));
            if (!Directory.Exists(shared)) continue;
            foreach (var verDir in Directory.GetDirectories(shared).OrderByDescending(d => d))
            {
                string candidate = Path.Combine(verDir, fileName);
                if (File.Exists(candidate))
                    return candidate;
            }
        }
    }

    return null;
}

static string? FindDllUnderLib(string packageVersionDir, string fileName)
{
    string lib = Path.Combine(packageVersionDir, "lib");
    if (!Directory.Exists(lib))
    {
        // some packages put dlls at root or in runtimes
        string direct = Path.Combine(packageVersionDir, fileName);
        if (File.Exists(direct)) return direct;
        return Directory.Exists(packageVersionDir)
            ? Directory.GetFiles(packageVersionDir, fileName, SearchOption.AllDirectories).FirstOrDefault()
            : null;
    }

    // Prefer netstandard2.0 / netcoreapp3.0 / net48 / net6.0 order
    string[] preferred = { "netstandard2.0", "netstandard2.1", "netcoreapp3.0", "netcoreapp3.1",
        "net462", "net47", "net48", "net6.0", "net7.0", "net8.0", "net9.0", "net10.0" };

    foreach (var tfm in preferred)
    {
        string candidate = Path.Combine(lib, tfm, fileName);
        if (File.Exists(candidate))
            return candidate;
    }

    return Directory.GetFiles(lib, fileName, SearchOption.AllDirectories).FirstOrDefault();
}

static void CollectAndResolveDependencies(
    Assembly asm,
    string dllDir,
    List<string> probeDirs,
    Dictionary<string, string> resolvedDeps,
    HashSet<string> depsToCopy,
    int depth)
{
    if (depth > 8) return;

    AssemblyName[] refs;
    try { refs = asm.GetReferencedAssemblies(); }
    catch { return; }

    foreach (var an in refs)
    {
        string simple = an.Name!;
        if (IsFrameworkAssembly(simple))
            continue;

        if (resolvedDeps.ContainsKey(simple))
        {
            if (resolvedDeps.TryGetValue(simple, out var p) && File.Exists(p))
                depsToCopy.Add(p);
            continue;
        }

        string? found = FindAssemblyFile(simple, an.Version, probeDirs);
        if (found == null)
        {
            // try already next to source
            string local = Path.Combine(dllDir, simple + ".dll");
            if (File.Exists(local))
                found = local;
        }

        if (found == null)
        {
            Console.WriteLine($"  [MISS] {simple} (version {an.Version})");
            continue;
        }

        resolvedDeps[simple] = found;
        depsToCopy.Add(found);
        Console.WriteLine($"  [FOUND] {simple} <= {found}");

        try
        {
            var depAsm = Assembly.LoadFrom(found);
            CollectAndResolveDependencies(depAsm, dllDir, probeDirs, resolvedDeps, depsToCopy, depth + 1);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  [WARN] Cannot recurse into {simple}: {ex.Message}");
        }
    }
}

static bool IsFrameworkAssembly(string simpleName)
{
    // Do not copy / reference core runtime & boxed assemblies
    if (simpleName.Equals("mscorlib", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.Equals("netstandard", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.Equals("System", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.Equals("System.Core", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.Equals("System.Runtime", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.Equals("WindowsBase", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.StartsWith("System.", StringComparison.OrdinalIgnoreCase))
    {
        // Must copy on modern .NET / netstandard bridges
        if (simpleName.Equals("System.Drawing.Common", StringComparison.OrdinalIgnoreCase))
            return false;
        if (simpleName.Equals("System.Drawing", StringComparison.OrdinalIgnoreCase))
            return false;
        return true;
    }
    // Package facades that cause MSB3277 if referenced explicitly alongside app refs
    if (simpleName.Equals("Microsoft.Win32.Registry", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.Equals("Microsoft.Win32.SystemEvents", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.StartsWith("Microsoft.CSharp", StringComparison.OrdinalIgnoreCase)) return true;
    if (simpleName.StartsWith("Microsoft.VisualBasic", StringComparison.OrdinalIgnoreCase)) return true;
    return false;
}

// =====================================================================
// Collect types recursively
// =====================================================================


static void CollectTypesToWrap(Type type, HashSet<Type> result, Dictionary<string, Type> allPublic)
{
    if (type.IsGenericTypeDefinition) return;
    if (!result.Add(type)) return;

    try
    {
        foreach (var ctor in type.GetConstructors(BindingFlags.Public | BindingFlags.Instance))
        {
            try
            {
                foreach (var p in ctor.GetParameters())
                    ConsiderType(p.ParameterType, result, allPublic);
            }
            catch (FileNotFoundException) { /* missing dependency in signature */ }
            catch (TypeLoadException) { }
        }

        foreach (var prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
        {
            try
            {
                if (prop.GetIndexParameters().Length == 0)
                    ConsiderType(prop.PropertyType, result, allPublic);
            }
            catch (FileNotFoundException) { }
            catch (TypeLoadException) { }
        }

        foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
        {
            if (method.IsSpecialName) continue;
            try
            {
                ConsiderType(method.ReturnType, result, allPublic);
                foreach (var p in method.GetParameters())
                    ConsiderType(p.ParameterType, result, allPublic);
            }
            catch (FileNotFoundException) { }
            catch (TypeLoadException) { }
        }

        foreach (var ev in type.GetEvents(BindingFlags.Public | BindingFlags.Instance))
        {
            try
            {
                if (ev.EventHandlerType == null) continue;
                var invoke = ev.EventHandlerType.GetMethod("Invoke");
                if (invoke == null) continue;
                foreach (var p in invoke.GetParameters())
                    ConsiderType(p.ParameterType, result, allPublic);
            }
            catch (FileNotFoundException) { }
            catch (TypeLoadException) { }
        }
    }
    catch (FileNotFoundException ex)
    {
        Console.WriteLine($"  [WARN] Skip dependencies of {type.Name}: {ex.FileName}");
    }
    catch (TypeLoadException ex)
    {
        Console.WriteLine($"  [WARN] Type load issue for {type.Name}: {ex.Message}");
    }
}

static void ConsiderType(Type t, HashSet<Type> result, Dictionary<string, Type> allPublic)
{
    t = Unwrap(t);

    if (IsCollectionType(t, out var itemType) && itemType != null)
    {
        ConsiderType(itemType, result, allPublic);
        return;
    }

    if (t.IsArray)
    {
        ConsiderType(t.GetElementType()!, result, allPublic);
        return;
    }

    if (allPublic.TryGetValue(t.FullName ?? "", out var found))
        CollectTypesToWrap(found, result, allPublic);
}

static Type Unwrap(Type t) => t.IsByRef ? t.GetElementType()! : t;

static bool IsCollectionType(Type t, out Type? itemType)
{
    itemType = null;
    t = Unwrap(t);

    if (t.IsGenericType)
    {
        var def = t.GetGenericTypeDefinition();
        if (def == typeof(List<>) ||
            def == typeof(IEnumerable<>) ||
            def == typeof(IList<>) ||
            def == typeof(ICollection<>) ||
            def == typeof(IReadOnlyList<>) ||
            def == typeof(IReadOnlyCollection<>))
        {
            itemType = t.GetGenericArguments()[0];
            return true;
        }
    }

    if (t.IsArray)
    {
        itemType = t.GetElementType();
        return true;
    }

    return false;
}

static Dictionary<MethodInfo, string> BuildUniqueMethodNames(IEnumerable<MethodInfo> methods)
{
    var groups = methods.GroupBy(m => m.Name);
    var result = new Dictionary<MethodInfo, string>();

    foreach (var group in groups)
    {
        var overloads = group.OrderBy(m => m.GetParameters().Length)
                             .ThenBy(m => m.ToString())
                             .ToList();

        if (overloads.Count == 1)
        {
            result[overloads[0]] = overloads[0].Name;
            continue;
        }

        result[overloads[0]] = overloads[0].Name;

        for (int i = 1; i < overloads.Count; i++)
        {
            var m = overloads[i];
            string suffix = string.Join("_", m.GetParameters().Select(p => GetTypeSuffix(p.ParameterType)));
            if (string.IsNullOrEmpty(suffix))
                suffix = $"Overload{i}";
            result[m] = $"{m.Name}_{suffix}";
        }
    }

    return result;
}

static string GetTypeSuffix(Type t)
{
    t = Unwrap(t);

    if (t == typeof(string)) return "String";
    if (t == typeof(int)) return "Int32";
    if (t == typeof(long)) return "Int64";
    if (t == typeof(bool)) return "Boolean";
    if (t == typeof(double)) return "Double";
    if (t == typeof(float)) return "Single";
    if (t == typeof(decimal)) return "Decimal";
    if (t == typeof(DateTime)) return "DateTime";
    if (t == typeof(Guid)) return "Guid";
    if (t == typeof(object)) return "Object";
    if (t == typeof(byte)) return "Byte";
    if (t == typeof(short)) return "Int16";
    if (t.IsEnum) return t.Name;

    if (IsCollectionType(t, out _)) return "List";
    if (t.IsArray) return GetTypeSuffix(t.GetElementType()!) + "Array";

    return t.Name;
}

static bool GenerateWrapper(
    StringBuilder sb,
    Type type,
    string ns,
    Dictionary<string, string> wrapperMap,
    Dictionary<string, Type> allPublic)
{
    string className = type.Name;
    string interfaceName = "I" + className;
    string eventsInterfaceName = "I" + className + "Events";
    string fullOriginalName = type.FullName!;

    List<ConstructorInfo> constructors;
    try
    {
        constructors = type.GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .Where(c =>
            {
                try { return c.GetParameters().All(p => IsSupported(p.ParameterType, wrapperMap)); }
                catch { return false; }
            })
            .ToList();
    }
    catch
    {
        constructors = new List<ConstructorInfo>();
    }

    bool hasParameterless = constructors.Any(c => c.GetParameters().Length == 0);
    var parameterizedCtors = constructors.Where(c => c.GetParameters().Length > 0).ToList();

    var instanceProps = GetInstanceProperties(type, wrapperMap);
    var instanceMethods = GetInstanceMethods(type, wrapperMap);
    var staticProps = GetStaticProperties(type, wrapperMap);
    var staticMethods = GetStaticMethods(type, wrapperMap);
    var indexers = GetIndexers(type, wrapperMap);
    var indexerNames = BuildIndexerNames(indexers);

    bool hasStaticMembers = staticProps.Count > 0 || staticMethods.Count > 0;
    bool hasIndexers = indexers.Count > 0;
    bool hasInstanceApi = instanceProps.Count > 0 || instanceMethods.Count > 0 || hasIndexers;
    // Always generate wrapper if type was selected: needed as return/arg of other types.
    // Even without public constructors (factory-only / internal DTOs).
    if (!hasParameterless && parameterizedCtors.Count == 0 && !hasStaticMembers && !hasInstanceApi)
    {
        // Truly empty public surface — still emit a thin shell so interface name exists
        Console.WriteLine($"  [OK]   {className} (opaque shell)");
    }

    var events = type.GetEvents(BindingFlags.Public | BindingFlags.Instance)
        .Where(e => e.EventHandlerType != null)
        .ToList();
    bool hasEvents = events.Count > 0;

    var allMethods = instanceMethods.Concat(staticMethods).ToList();
    var uniqueNames = BuildUniqueMethodNames(allMethods);

    Guid interfaceGuid = Guid.NewGuid();
    Guid classGuid = Guid.NewGuid();
    Guid eventsGuid = Guid.NewGuid();

    sb.AppendLine();
    sb.AppendLine($"    // =================================================================");
    sb.AppendLine($"    // {className}");
    sb.AppendLine($"    // =================================================================");
    sb.AppendLine();

    if (hasEvents)
    {
        sb.AppendLine($"    [ComVisible(true)]");
        sb.AppendLine($"    [Guid(\"{eventsGuid}\")]");
        sb.AppendLine($"    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]");
        sb.AppendLine($"    public interface {eventsInterfaceName}");
        sb.AppendLine("    {");

        int dispId = 1;
        foreach (var ev in events)
        {
            var parameters = GetEventParameters(ev);
            var comParams = FormatEventParameters(parameters, wrapperMap);
            sb.AppendLine($"        [DispId({dispId++})]");
            sb.AppendLine($"        void {ev.Name}({string.Join(", ", comParams)});");
        }

        sb.AppendLine("    }");
        sb.AppendLine();
    }

    sb.AppendLine($"    [ComVisible(true)]");
    sb.AppendLine($"    [Guid(\"{interfaceGuid}\")]");
    sb.AppendLine($"    [InterfaceType(ComInterfaceType.InterfaceIsDual)]");
    sb.AppendLine($"    public interface {interfaceName}");
    sb.AppendLine("    {");

    foreach (var ctor in parameterizedCtors)
    {
        string paramList = string.Join(", ", ctor.GetParameters()
            .Select(p => FormatParameter(p, wrapperMap, forCom: true)));
        sb.AppendLine($"        {interfaceName} Create({paramList});");
    }

    var usedMemberNames = new HashSet<string>(StringComparer.Ordinal);

    foreach (var prop in instanceProps)
    {
        if (!usedMemberNames.Add(prop.Name)) continue;
        WritePropertyToInterface(sb, prop, wrapperMap);
    }

    foreach (var method in instanceMethods)
    {
        string methodName = uniqueNames[method];
        if (!usedMemberNames.Add(methodName)) continue;
        WriteMethodToInterface(sb, method, methodName, wrapperMap);
    }

    foreach (var idx in indexers)
    {
        var (getName, setName) = indexerNames[idx];
        if (usedMemberNames.Add(getName) || true)
            WriteIndexerToInterface(sb, idx, getName, setName, wrapperMap);
        usedMemberNames.Add(getName);
        usedMemberNames.Add(setName);
    }

    foreach (var prop in staticProps)
    {
        if (!usedMemberNames.Add(prop.Name)) continue;
        WritePropertyToInterface(sb, prop, wrapperMap);
    }

    foreach (var method in staticMethods)
    {
        string methodName = uniqueNames[method];
        if (!usedMemberNames.Add(methodName)) continue;
        WriteMethodToInterface(sb, method, methodName, wrapperMap);
    }

    sb.AppendLine("    }");
    sb.AppendLine();

    sb.AppendLine($"    [ComVisible(true)]");
    sb.AppendLine($"    [Guid(\"{classGuid}\")]");
    sb.AppendLine($"    [ClassInterface(ClassInterfaceType.None)]");
    if (hasEvents)
        sb.AppendLine($"    [ComSourceInterfaces(typeof({eventsInterfaceName}))]");
    sb.AppendLine($"    [ProgId(\"{ns}.{className}\")]");
    sb.AppendLine($"    public class {className} : {interfaceName}");
    sb.AppendLine("    {");
    sb.AppendLine($"        private {fullOriginalName}? _inner;");
    sb.AppendLine();
    sb.AppendLine($"        internal {fullOriginalName}? Inner => _inner;");
    sb.AppendLine();

    sb.AppendLine($"        public {className}()");
    sb.AppendLine("        {");
    if (hasParameterless)
        sb.AppendLine($"            _inner = new {fullOriginalName}();");
    else
        sb.AppendLine("            _inner = null;");
    sb.AppendLine("        }");
    sb.AppendLine();

    sb.AppendLine($"        internal {className}({fullOriginalName} inner)");
    sb.AppendLine("        {");
    sb.AppendLine("            _inner = inner;");
    if (hasEvents)
        sb.AppendLine("            SubscribeEvents();");
    sb.AppendLine("        }");
    sb.AppendLine();

    foreach (var ctor in parameterizedCtors)
    {
        var parameters = ctor.GetParameters();
        string paramList = string.Join(", ", parameters.Select(p => FormatParameter(p, wrapperMap, forCom: true)));
        string argList = string.Join(", ", parameters.Select(p => ConvertToOriginal(p.Name!, p.ParameterType, wrapperMap)));

        sb.AppendLine($"        public {interfaceName} Create({paramList})");
        sb.AppendLine("        {");
        sb.AppendLine($"            var original = new {fullOriginalName}({argList});");
        sb.AppendLine($"            return new {className}(original);");
        sb.AppendLine("        }");
        sb.AppendLine();
    }

    sb.AppendLine("        private void EnsureInitialized()");
    sb.AppendLine("        {");
    sb.AppendLine("            if (_inner == null)");
    sb.AppendLine($"                throw new InvalidOperationException(\"Object '{className}' is not created yet. Call Create(...) first.\");");
    sb.AppendLine("        }");
    sb.AppendLine();

    if (hasEvents)
    {
        sb.AppendLine("        private void SubscribeEvents()");
        sb.AppendLine("        {");
        sb.AppendLine("            if (_inner == null) return;");
        sb.AppendLine();

        foreach (var ev in events)
        {
            var parameters = GetEventParameters(ev);
            var paramNames = parameters.Select(p => p.Name!).ToList();
            string lambdaParams = string.Join(", ", paramNames);

            var raiseArgs = parameters.Select(p =>
            {
                if (p.Name == "sender" && p.ParameterType == typeof(object))
                    return "sender";
                return ConvertToWrapper(p.Name!, p.ParameterType, wrapperMap);
            });

            sb.AppendLine($"            _inner.{ev.Name} += ({lambdaParams}) =>");
            sb.AppendLine("            {");
            sb.AppendLine($"                {ev.Name}?.Invoke({string.Join(", ", raiseArgs)});");
            sb.AppendLine("            };");
            sb.AppendLine();
        }

        sb.AppendLine("        }");
        sb.AppendLine();

        foreach (var ev in events)
        {
            var parameters = GetEventParameters(ev);
            var comParams = FormatEventParameters(parameters, wrapperMap);
            string paramList = string.Join(", ", comParams);

            // Unique name avoids TYPE_E_NAMECONFLICT in TLB (dscom flattens nested types)
            string handlerTypeName = $"{className}_{ev.Name}EventHandler";
            sb.AppendLine($"        public delegate void {handlerTypeName}({paramList});");
            sb.AppendLine($"        public event {handlerTypeName}? {ev.Name};");
        }
        sb.AppendLine();
    }

    var usedImplNames = new HashSet<string>(StringComparer.Ordinal);

    foreach (var prop in instanceProps)
    {
        if (!usedImplNames.Add(prop.Name)) continue;
        WriteInstancePropertyImpl(sb, prop, wrapperMap);
    }

    foreach (var method in instanceMethods)
    {
        string methodName = uniqueNames[method];
        if (!usedImplNames.Add(methodName)) continue;
        WriteInstanceMethodImpl(sb, method, methodName, wrapperMap);
    }

    foreach (var idx in indexers)
    {
        var (getName, setName) = indexerNames[idx];
        if (!usedImplNames.Add(getName)) continue;
        usedImplNames.Add(setName);
        WriteIndexerImpl(sb, idx, getName, setName, wrapperMap);
    }

    foreach (var prop in staticProps)
    {
        if (!usedImplNames.Add(prop.Name)) continue;
        WriteStaticPropertyImpl(sb, prop, fullOriginalName, wrapperMap);
    }

    foreach (var method in staticMethods)
    {
        string methodName = uniqueNames[method];
        if (!usedImplNames.Add(methodName)) continue;
        WriteStaticMethodImpl(sb, method, methodName, fullOriginalName, wrapperMap);
    }

    sb.AppendLine("    }");

    Console.WriteLine($"  [OK]   {className}" +
                      (parameterizedCtors.Count > 0 ? " (factory)" : "") +
                      (hasEvents ? " (events)" : "") +
                      (hasStaticMembers ? " (static)" : "") +
                      (hasIndexers ? " (indexers)" : ""));
    return true;
}

static List<PropertyInfo> GetInstanceProperties(Type type, Dictionary<string, string> wrapperMap)
{
    // DeclaredOnly avoids duplicate properties from base types; then distinct by name
    return type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
        .Where(p => p.GetIndexParameters().Length == 0)
        .Where(p => IsSupported(p.PropertyType, wrapperMap))
        .GroupBy(p => p.Name)
        .Select(g => g.First())
        .ToList();
}

static List<MethodInfo> GetInstanceMethods(Type type, Dictionary<string, string> wrapperMap)
{
    return type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
        .Where(m => !m.IsSpecialName)
        .Where(m => m.Name is not ("ToString" or "GetHashCode" or "Equals" or "GetType"))
        .Where(m =>
        {
            bool ok = m.ReturnType == typeof(void) || IsSupported(m.ReturnType, wrapperMap);
            ok &= m.GetParameters().All(p => IsSupported(p.ParameterType, wrapperMap));
            return ok;
        })
        .ToList();
}

static List<PropertyInfo> GetStaticProperties(Type type, Dictionary<string, string> wrapperMap)
{
    return type.GetProperties(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
        .Where(p => p.GetIndexParameters().Length == 0)
        .Where(p => IsSupported(p.PropertyType, wrapperMap))
        .GroupBy(p => p.Name)
        .Select(g => g.First())
        .ToList();
}

static List<MethodInfo> GetStaticMethods(Type type, Dictionary<string, string> wrapperMap)
{
    return type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
        .Where(m => !m.IsSpecialName)
        .Where(m => m.Name is not ("ToString" or "GetHashCode" or "Equals" or "GetType"))
        .Where(m =>
        {
            bool ok = m.ReturnType == typeof(void) || IsSupported(m.ReturnType, wrapperMap);
            ok &= m.GetParameters().All(p => IsSupported(p.ParameterType, wrapperMap));
            return ok;
        })
        .ToList();
}

static List<PropertyInfo> GetIndexers(Type type, Dictionary<string, string> wrapperMap)
{
    return type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
        .Where(p => p.GetIndexParameters().Length > 0)
        .Where(p => IsSupported(p.PropertyType, wrapperMap))
        .Where(p => p.GetIndexParameters().All(ip => IsSupported(ip.ParameterType, wrapperMap)))
        .ToList();
}

/// <summary>
/// Build unique GetItem/SetItem names for indexers (overload renaming by index parameter types).
/// </summary>
static Dictionary<PropertyInfo, (string getName, string setName)> BuildIndexerNames(IEnumerable<PropertyInfo> indexers)
{
    var list = indexers.OrderBy(p => p.GetIndexParameters().Length)
                       .ThenBy(p => string.Join(",", p.GetIndexParameters().Select(x => x.ParameterType.FullName)))
                       .ToList();
    var result = new Dictionary<PropertyInfo, (string getName, string setName)>();

    if (list.Count == 0) return result;

    if (list.Count == 1)
    {
        var p = list[0];
        result[p] = ("GetItem", "SetItem");
        return result;
    }

    // First keeps plain GetItem/SetItem; rest get type suffix
    result[list[0]] = ("GetItem", "SetItem");
    for (int i = 1; i < list.Count; i++)
    {
        var p = list[i];
        string suffix = string.Join("_", p.GetIndexParameters().Select(ip => GetTypeSuffix(ip.ParameterType)));
        if (string.IsNullOrEmpty(suffix)) suffix = $"Overload{i}";
        result[p] = ($"GetItem_{suffix}", $"SetItem_{suffix}");
    }
    return result;
}

static void WritePropertyToInterface(StringBuilder sb, PropertyInfo prop, Dictionary<string, string> wrapperMap)
{
    string typeName = GetTypeName(prop.PropertyType, wrapperMap, forCom: true);
    if (prop.CanRead && prop.CanWrite)
        sb.AppendLine($"        {typeName} {prop.Name} {{ get; set; }}");
    else if (prop.CanRead)
        sb.AppendLine($"        {typeName} {prop.Name} {{ get; }}");
    else if (prop.CanWrite)
        sb.AppendLine($"        {typeName} {prop.Name} {{ set; }}");
}

static void WriteMethodToInterface(StringBuilder sb, MethodInfo method, string methodName, Dictionary<string, string> wrapperMap)
{
    string returnType = method.ReturnType == typeof(void)
        ? "void"
        : GetTypeName(method.ReturnType, wrapperMap, forCom: true);

    string parameters = string.Join(", ", method.GetParameters()
        .Select(p => FormatParameter(p, wrapperMap, forCom: true)));

    sb.AppendLine($"        {returnType} {methodName}({parameters});");
}

static void WriteIndexerToInterface(StringBuilder sb, PropertyInfo indexer, string getName, string setName, Dictionary<string, string> wrapperMap)
{
    string valueType = GetTypeName(indexer.PropertyType, wrapperMap, forCom: true);
    string indexParams = string.Join(", ", indexer.GetIndexParameters()
        .Select(p => $"{GetTypeName(p.ParameterType, wrapperMap, forCom: true)} {p.Name}"));

    if (indexer.CanRead)
        sb.AppendLine($"        {valueType} {getName}({indexParams});");
    if (indexer.CanWrite)
        sb.AppendLine($"        void {setName}({indexParams}, {valueType} value);");
}

static void WriteIndexerImpl(StringBuilder sb, PropertyInfo indexer, string getName, string setName, Dictionary<string, string> wrapperMap)
{
    string valueType = GetTypeName(indexer.PropertyType, wrapperMap, forCom: true);
    var indexParamsInfos = indexer.GetIndexParameters();
    string indexParams = string.Join(", ", indexParamsInfos
        .Select(p => $"{GetTypeName(p.ParameterType, wrapperMap, forCom: true)} {p.Name}"));
    string indexArgs = string.Join(", ", indexParamsInfos.Select(p => p.Name));

    if (indexer.CanRead)
    {
        sb.AppendLine($"        public {valueType} {getName}({indexParams})");
        sb.AppendLine("        {");
        sb.AppendLine("            EnsureInitialized();");
        sb.AppendLine($"            return {ConvertToWrapper($"_inner![{indexArgs}]", indexer.PropertyType, wrapperMap)};");
        sb.AppendLine("        }");
        sb.AppendLine();
    }

    if (indexer.CanWrite)
    {
        sb.AppendLine($"        public void {setName}({indexParams}, {valueType} value)");
        sb.AppendLine("        {");
        sb.AppendLine("            EnsureInitialized();");
        sb.AppendLine($"            _inner![{indexArgs}] = {ConvertToOriginal("value", indexer.PropertyType, wrapperMap)};");
        sb.AppendLine("        }");
        sb.AppendLine();
    }
}

static void WriteInstancePropertyImpl(StringBuilder sb, PropertyInfo prop, Dictionary<string, string> wrapperMap)
{
    string comType = GetTypeName(prop.PropertyType, wrapperMap, forCom: true);

    sb.AppendLine($"        public {comType} {prop.Name}");
    sb.AppendLine("        {");

    if (prop.CanRead)
    {
        sb.AppendLine("            get");
        sb.AppendLine("            {");
        sb.AppendLine("                EnsureInitialized();");
        sb.AppendLine($"                return {ConvertToWrapper($"_inner!.{prop.Name}", prop.PropertyType, wrapperMap)};");
        sb.AppendLine("            }");
    }

    if (prop.CanWrite)
    {
        sb.AppendLine("            set");
        sb.AppendLine("            {");
        sb.AppendLine("                EnsureInitialized();");
        sb.AppendLine($"                _inner!.{prop.Name} = {ConvertToOriginal("value", prop.PropertyType, wrapperMap)};");
        sb.AppendLine("            }");
    }

    sb.AppendLine("        }");
    sb.AppendLine();
}

static void WriteInstanceMethodImpl(StringBuilder sb, MethodInfo method, string methodName, Dictionary<string, string> wrapperMap)
{
    string returnType = method.ReturnType == typeof(void)
        ? "void"
        : GetTypeName(method.ReturnType, wrapperMap, forCom: true);

    string parameters = string.Join(", ", method.GetParameters()
        .Select(p => FormatParameter(p, wrapperMap, forCom: true)));

    var callArgs = method.GetParameters().Select(p =>
    {
        if (p.IsOut) return $"out {p.Name}";
        if (p.ParameterType.IsByRef) return $"ref {p.Name}";
        return ConvertToOriginal(p.Name!, p.ParameterType, wrapperMap);
    });

    sb.AppendLine($"        public {returnType} {methodName}({parameters})");
    sb.AppendLine("        {");
    sb.AppendLine("            EnsureInitialized();");

    if (method.ReturnType == typeof(void))
        sb.AppendLine($"            _inner!.{method.Name}({string.Join(", ", callArgs)});");
    else
    {
        sb.AppendLine($"            var result = _inner!.{method.Name}({string.Join(", ", callArgs)});");
        sb.AppendLine($"            return {ConvertToWrapper("result", method.ReturnType, wrapperMap)};");
    }

    sb.AppendLine("        }");
    sb.AppendLine();
}

static void WriteStaticPropertyImpl(StringBuilder sb, PropertyInfo prop, string fullOriginalName, Dictionary<string, string> wrapperMap)
{
    string comType = GetTypeName(prop.PropertyType, wrapperMap, forCom: true);

    sb.AppendLine($"        public {comType} {prop.Name}");
    sb.AppendLine("        {");

    if (prop.CanRead)
    {
        sb.AppendLine("            get");
        sb.AppendLine("            {");
        sb.AppendLine($"                return {ConvertToWrapper($"{fullOriginalName}.{prop.Name}", prop.PropertyType, wrapperMap)};");
        sb.AppendLine("            }");
    }

    if (prop.CanWrite)
    {
        sb.AppendLine("            set");
        sb.AppendLine("            {");
        sb.AppendLine($"                {fullOriginalName}.{prop.Name} = {ConvertToOriginal("value", prop.PropertyType, wrapperMap)};");
        sb.AppendLine("            }");
    }

    sb.AppendLine("        }");
    sb.AppendLine();
}

static void WriteStaticMethodImpl(StringBuilder sb, MethodInfo method, string methodName, string fullOriginalName, Dictionary<string, string> wrapperMap)
{
    string returnType = method.ReturnType == typeof(void)
        ? "void"
        : GetTypeName(method.ReturnType, wrapperMap, forCom: true);

    string parameters = string.Join(", ", method.GetParameters()
        .Select(p => FormatParameter(p, wrapperMap, forCom: true)));

    var callArgs = method.GetParameters().Select(p =>
    {
        if (p.IsOut) return $"out {p.Name}";
        if (p.ParameterType.IsByRef) return $"ref {p.Name}";
        return ConvertToOriginal(p.Name!, p.ParameterType, wrapperMap);
    });

    sb.AppendLine($"        public {returnType} {methodName}({parameters})");
    sb.AppendLine("        {");

    if (method.ReturnType == typeof(void))
        sb.AppendLine($"            {fullOriginalName}.{method.Name}({string.Join(", ", callArgs)});");
    else
    {
        sb.AppendLine($"            var result = {fullOriginalName}.{method.Name}({string.Join(", ", callArgs)});");
        sb.AppendLine($"            return {ConvertToWrapper("result", method.ReturnType, wrapperMap)};");
    }

    sb.AppendLine("        }");
    sb.AppendLine();
}

static bool IsSupported(Type t, Dictionary<string, string> wrapperMap)
{
    t = Unwrap(t);

    if (t == typeof(void) || t.IsPrimitive) return true;
    if (t == typeof(string) || t == typeof(decimal) || t == typeof(DateTime) ||
        t == typeof(Guid) || t == typeof(object)) return true;
    if (t.IsEnum) return true;

    if (IsCollectionType(t, out var itemType))
        return itemType != null && IsSupported(itemType, wrapperMap);

    if (wrapperMap.ContainsKey(t.FullName ?? ""))
        return true;

    return false;
}

static string GetTypeName(Type t, Dictionary<string, string> wrapperMap, bool forCom)
{
    t = Unwrap(t);

    if (t == typeof(void)) return "void";
    if (t == typeof(string)) return "string";
    if (t == typeof(int)) return "int";
    if (t == typeof(long)) return "long";
    if (t == typeof(bool)) return "bool";
    if (t == typeof(double)) return "double";
    if (t == typeof(float)) return "float";
    if (t == typeof(decimal)) return "decimal";
    if (t == typeof(DateTime)) return "DateTime";
    if (t == typeof(Guid)) return "Guid";
    if (t == typeof(object)) return "object";
    if (t == typeof(byte)) return "byte";
    if (t == typeof(short)) return "short";
    if (t == typeof(uint)) return "uint";
    if (t == typeof(ulong)) return "ulong";
    if (t.IsEnum) return forCom ? "int" : (t.FullName ?? t.Name);

    if (IsCollectionType(t, out _))
    {
        if (forCom) return "object[]";
        return t.FullName ?? t.Name;
    }

    if (wrapperMap.TryGetValue(t.FullName ?? "", out var iface))
        return iface;

    return t.FullName ?? t.Name;
}

static string FormatParameter(ParameterInfo p, Dictionary<string, string> wrapperMap, bool forCom)
{
    Type pt = Unwrap(p.ParameterType);
    string typeName = GetTypeName(pt, wrapperMap, forCom);

    if (p.IsOut) return $"out {typeName} {p.Name}";
    if (p.ParameterType.IsByRef) return $"ref {typeName} {p.Name}";
    return $"{typeName} {p.Name}";
}

static string ConvertToOriginal(string varName, Type type, Dictionary<string, string> wrapperMap)
{
    type = Unwrap(type);

    // Enums are exposed as int on COM side
    if (type.IsEnum)
        return $"({type.FullName}){varName}";

    if (IsCollectionType(type, out var itemType) && itemType != null)
    {
        string elemCast;
        if (wrapperMap.ContainsKey(itemType.FullName ?? ""))
        {
            elemCast = $"{varName}.Cast<{wrapperMap[itemType.FullName!]}>()" +
                       $".Select(x => (({itemType.Name})x).Inner!)";
        }
        else if (itemType.IsEnum)
        {
            elemCast = $"{varName}.Cast<int>().Select(x => ({itemType.FullName})x)";
        }
        else
        {
            elemCast = $"{varName}.Cast<{GetTypeName(itemType, wrapperMap, false)}>()";
        }

        // Preserve array vs List vs IEnumerable
        if (type.IsArray)
            return $"{elemCast}.ToArray()";
        if (type.IsGenericType)
        {
            var def = type.GetGenericTypeDefinition();
            if (def == typeof(List<>))
                return $"{elemCast}.ToList()";
            // IList/IEnumerable/ICollection — List is assignable
            return $"{elemCast}.ToList()";
        }
        return $"{elemCast}.ToList()";
    }

    if (wrapperMap.ContainsKey(type.FullName ?? ""))
        return $"(({type.Name}){varName}).Inner!";

    return varName;
}

static string ConvertToWrapper(string expression, Type type, Dictionary<string, string> wrapperMap)
{
    type = Unwrap(type);

    if (type.IsEnum)
        return $"(int){expression}";

    if (IsCollectionType(type, out var itemType) && itemType != null)
    {
        if (wrapperMap.ContainsKey(itemType.FullName ?? ""))
        {
            return $"{expression}.Select(x => ({wrapperMap[itemType.FullName!]})new {itemType.Name}(x))" +
                   $".Cast<object>().ToArray()";
        }
        if (itemType.IsEnum)
            return $"{expression}.Select(x => (object)(int)x).ToArray()";
        return $"{expression}.Cast<object>().ToArray()";
    }

    if (wrapperMap.ContainsKey(type.FullName ?? ""))
        return $"new {type.Name}({expression})";

    return expression;
}

static ParameterInfo[] GetEventParameters(EventInfo ev)
{
    var invoke = ev.EventHandlerType?.GetMethod("Invoke");
    return invoke?.GetParameters() ?? Array.Empty<ParameterInfo>();
}

static List<string> FormatEventParameters(ParameterInfo[] parameters, Dictionary<string, string> wrapperMap)
{
    var result = new List<string>();
    foreach (var p in parameters)
    {
        if (p.Name == "sender" && p.ParameterType == typeof(object))
        {
            result.Add("object sender");
            continue;
        }
        string typeName = GetTypeName(p.ParameterType, wrapperMap, forCom: true);
        result.Add($"{typeName} {p.Name}");
    }
    return result;
}
