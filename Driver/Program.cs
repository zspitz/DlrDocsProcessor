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
using Kurukuru;

const string pandocPath = @"""C:\Program Files\Pandoc\pandoc.exe""";

var rootPath = GetFullPath($"{GetParent(GetExecutingAssembly().Location)}");
var sourcePath = Combine(rootPath, "source");
var outputPath = Combine(rootPath, "output");
if (Directory.Exists(outputPath)) {
    Directory.Delete(outputPath, true);
}
CreateDirectory(outputPath);

var filterPath = Combine(rootPath, "Filter.exe");

foreach (var doc in EnumerateFiles(sourcePath).Where(x => !x.Contains("~$"))) {
    WriteLine($"Beginning {doc}");

    Spinner.Start($"Word processing", () => {
        Process process = new();
        process.StartInfo = new() {
            FileName = "WordProcessing.exe",
            Arguments = $@"""{doc}""",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = rootPath
        };
        process.EnableRaisingEvents = true;

        var result = RunProcess(process);
        if (!result.StdOut.IsNullOrWhitespace()) { WriteLine(result.StdOut); }
        if (!result.StdErr.IsNullOrWhitespace()) { WriteLine(result.StdErr); }
    });

    var name = GetFileNameWithoutExtension(doc).ToLower();
    var docRoot = Combine(outputPath, name);
    if (!Directory.Exists(docRoot)) { CreateDirectory(docRoot); }

    Spinner.Start($"First pass; sidebar generation", () => {
        RunCmd(docRoot, $"{pandocPath} -s {doc} -t gfm --extract-media=. -F {filterPath} --wrap=preserve -o index.md");
    });

    var headersPath = Combine(docRoot, "headers.json");
    var json = File.ReadAllText(headersPath);
    JsonSerializer.Deserialize<Dictionary<int, string>>(json)!.ForEachKVP((pass, id) => {
        Spinner.Start($"Pass: {pass}, id: {id}", () => {
            RunCmd(docRoot, $"{pandocPath} -s {doc} -t json | {filterPath} --pass={pass} | {pandocPath} -s -f json -t gfm+gfm_auto_identifiers --wrap=preserve -o {id}.md");

            // Pandoc emits bullet lists with 3 spaces after the bullet
            // reduce to one space
            var currentHeadingPath = Combine(docRoot, $"{id}.md");
            var lines = File.ReadLines(currentHeadingPath).Select(line => {
                if (line.StartsWith("-   ")) {
                    line = $"- {line[4..]}";
                }
                return line;
            }).ToList();
            File.WriteAllLines(
                currentHeadingPath,
                lines
            );
        });
    });

    File.Delete(headersPath);
}

// copy SVG files from source to corresponding location of EMF
// delete EMF
var correspondingSvg =
    Directory.EnumerateFiles(outputPath!, "*.emf", new EnumerationOptions {
        MatchCasing = MatchCasing.CaseInsensitive,
        RecurseSubdirectories = true
    })
    .Select(emfPath => (
        emfPath,
        svgPath: ChangeExtension(emfPath.Replace(outputPath, sourcePath), ".svg")
    ));

var missingSvg =
    correspondingSvg
       .WhereT((_, svgPath) => !File.Exists(svgPath))
       .ToList();

if (missingSvg.Any()) {
    WriteLine("The following .emf files need to be converted to svg. Press any key to conitnue.");
    WriteLine(missingSvg.JoinedT("\n\n", (emfPath, svgPath) => $"{emfPath}\n{svgPath}"));
    return;
}

correspondingSvg.ForEachT((emfPath, svgSourcePath) => {
    var targetSvgPath = ChangeExtension(emfPath, ".svg");
    File.Copy(svgSourcePath, targetSvgPath, true);
    File.Delete(emfPath);
});


static void RunCmd(string workingDirectory, string arguments) {
    Process process = new();
    process.StartInfo = new() {
        FileName = "cmd",
        Arguments = @$"/C ""{arguments}""",
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        WorkingDirectory = workingDirectory
    };
    process.EnableRaisingEvents = true;

    var result = RunProcess(process);
    if (!result.StdOut.IsNullOrWhitespace()) { WriteLine(result.StdOut); }
    if (!result.StdErr.IsNullOrWhitespace()) { WriteLine(result.StdErr); }
}
