using System;
using static System.IO.Path;
using static System.IO.Directory;
using static System.Reflection.Assembly;
using System.IO;
using static System.Console;
using System.Linq;
using static ZSpitz.Util.Functions;
using System.Diagnostics;
using System.Text.Json;
using System.Collections.Generic;
using ZSpitz.Util;

const string pandocPath = @"""C:\Program Files\Pandoc\pandoc.exe""";

var rootPath = GetFullPath($"{GetParent(GetExecutingAssembly().Location)}");
var sourcePath = Combine(rootPath, "source");
var outputPath = Combine(rootPath, "output");
if (Directory.Exists(outputPath)) {
    //WriteLine("Delete output folder?");
    //if (ReadKey().Key != ConsoleKey.Y) { return; }
    Directory.Delete(outputPath, true);
}
CreateDirectory(outputPath);

var filterPath = Combine(rootPath, "Filter.exe");
foreach (var doc in EnumerateFiles(sourcePath).Where(x => !x.Contains("~$"))) {
    var name = GetFileNameWithoutExtension(doc).ToLower();
    var docRoot = Combine(outputPath, name);
    if (!Directory.Exists(docRoot)) { CreateDirectory(docRoot); }

    Process process = new();
    process.StartInfo = new() {
        FileName = "cmd",
        Arguments = @$"/C ""{pandocPath} -s {doc} -t json | {filterPath}""",
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        WorkingDirectory = docRoot
    };
    process.EnableRaisingEvents = true;

    var result = RunProcess(process);
    WriteLine(result.StdOut);
    WriteLine(result.StdErr);

    var json = File.ReadAllText(Combine(docRoot, "headers.json"));
    JsonSerializer.Deserialize<Dictionary<int, string>>(json)!.ForEachKVP((pass, id) => {
        process = new();
        process.StartInfo = new() {
            FileName = "cmd",
            Arguments = @$"/C ""{pandocPath} -s {doc} -t json | {filterPath} --pass={pass} | {pandocPath} -s -f json -t gfm --wrap=preserve --extract-media=. -o {id}.md""",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = docRoot
        };
        process.EnableRaisingEvents = true;

        result = RunProcess(process);
        WriteLine(result.StdOut);
        WriteLine(result.StdErr);
    });

    // while testing, we want to break early
    break;
}
