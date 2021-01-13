﻿using System;
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

    var name = GetFileNameWithoutExtension(doc).ToLower();
    var docRoot = Combine(outputPath, name);
    if (!Directory.Exists(docRoot)) { CreateDirectory(docRoot); }

    Spinner.Start($"First pass; sidebar generation", () => {
        Process process = new();
        process.StartInfo = new() {
            FileName = "cmd",
            Arguments = @$"/C ""{pandocPath} -s {doc} -t gfm --extract-media=. -F {filterPath} --wrap=preserve -o _sidebar.md""",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = docRoot
        };
        process.EnableRaisingEvents = true;

        var result = RunProcess(process);
        if (!result.StdOut.IsNullOrWhitespace()) { WriteLine(result.StdOut); }
        if (!result.StdErr.IsNullOrWhitespace()) { WriteLine(result.StdErr); }
    });

    var json = File.ReadAllText(Combine(docRoot, "headers.json"));
    JsonSerializer.Deserialize<Dictionary<int, string>>(json)!.ForEachKVP((pass, id) => {
        Spinner.Start($"Starting pass: {pass}, id: {id}", () => {
            Process process = new();
            process.StartInfo = new() {
                FileName = "cmd",
                Arguments = @$"/C ""{pandocPath} -s {doc} -t json | {filterPath} --pass={pass} | {pandocPath} -s -f json -t gfm+gfm_auto_identifiers --wrap=preserve -o {id}.md""",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = docRoot
            };
            process.EnableRaisingEvents = true;

            var result = RunProcess(process);
            if (!result.StdOut.IsNullOrWhitespace()) { WriteLine(result.StdOut); }
            if (!result.StdErr.IsNullOrWhitespace()) { WriteLine(result.StdErr); }
        });
    });

    // while testing, we want to break early
    break;
}
