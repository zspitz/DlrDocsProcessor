using System;
using static System.IO.Path;
using static System.IO.Directory;
using static System.Reflection.Assembly;
using System.IO;
using static System.Console;
using System.Linq;
using static ZSpitz.Util.Functions;
using System.Diagnostics;

const string pandocPath = @"c:\Program Files\Pandoc\pandoc.exe";

var rootPath = GetFullPath($"{GetParent(GetExecutingAssembly().Location)}");
var sourcePath = Combine(rootPath, "source");
var outputPath = Combine(rootPath, "output");
if (Directory.Exists(outputPath)) {
    WriteLine("Delete output folder?");
    if (ReadKey().Key != ConsoleKey.Y) { return; }
    Directory.Delete(outputPath, true);
}
CreateDirectory(outputPath);

foreach (var doc in EnumerateFiles(sourcePath).Where(x => !x.Contains("~$"))) {
    var name = GetFileNameWithoutExtension(doc);
    var docRoot = Combine(outputPath, name);
    if (!Directory.Exists(docRoot)) { CreateDirectory(docRoot); }

    Process process = new();
    process.StartInfo = new() {
        FileName = "cmd",
        Arguments = @"/C """"C:\Program Files\Pandoc\pandoc.exe"" -v""",
        RedirectStandardOutput=true,
        RedirectStandardError = true
    };
    process.EnableRaisingEvents = true;
    
    var result = RunProcess(process);
    WriteLine(result.StdOut);
    WriteLine(result.StdErr);
}