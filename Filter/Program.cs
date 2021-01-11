using System.Collections.Immutable;
using PandocFilters;
using PandocFilters.Ast;
using ZSpitz.Util;
using System.Linq;
using System.CommandLine;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;
using System.Diagnostics;
using System;
using static ZSpitz.Util.Functions;

var command = new RootCommand() {
    new Option<int>("--pass", () => -1)
};
var result = command.Parse(args);
var pass = result.ValueForOption<int>("--pass");

var headerCount = 0;

if (pass == -1) {
    var firstPass = new DelegateVisitor();

    // generate headers.json
    var headers = new Dictionary<int, string>();
    firstPass.Add((Pandoc pandoc) => {
        foreach (var block in pandoc.Blocks) {
            if (block.TryPickT9(out var header, out var _) && header.Level == 1) {
                headerCount += 1;
                headers.Add(headerCount, header.Attr.Identifier);
            } else if (headerCount == 0 && !headers.ContainsKey(0)) {
                headers.Add(0, "frontmatter");
            }
        }
        return pandoc;
    });
    Filter.Run(firstPass);

    var jsonString = JsonSerializer.Serialize(headers);
    File.WriteAllText("headers.json", jsonString);
    return;
}

var numberingGenerator = new DelegateVisitor();
// insert hierarchal numbering into headers
numberingGenerator.Add((Header header) =>
    header with
    {
        Text = new Inline[] {
                new Str(HeaderNumber.Next(header)),
                new Space()
        }.Concat(header.Text).ToImmutableList()
    }
);

var splitter = new DelegateVisitor();
splitter.Add((Pandoc pandoc) => 
    pandoc with
    {
        Blocks = pandoc.Blocks.Select(block => {
            if (block.IsT9 && block.AsT9.Level == 1) {
                headerCount += 1;
            }
            return (block, headerCount);
        })
        .WhereT((_, headerCount) => headerCount == pass)
        .SelectT((block, _) => block)
        .ToImmutableList()
    }
);


var visitor = new DelegateVisitor();

// remove height and width from images, so they'll output as standard markdown images, instead of HTML img elements
visitor.Add((Image img) => 
    img with
    {
        Attr = img.Attr with
        {
            KeyValuePairs = img.Attr.KeyValuePairs.WhereT((key, value) => key.NotIn("height", "width")).ToImmutableList()
        }
    }
);

// replace underline with emphasis for markdown bold
visitor.Add((Inline inline) => {
    if (inline.TryPickT2(out var underline, out var _)) {
        return new Strong(underline.Inlines);
    }
    return inline;
});

Filter.Run(numberingGenerator, splitter, visitor);

static class HeaderNumber {
    private static (int, int, int, int) current { get; set; }
    public static string Next(Header header) {
        switch (header.Level) {
            case 1:
                current = (current.Item1 + 1, 0, 0, 0);
                break;
            case 2:
                current = (current.Item1, current.Item2 + 1, 0, 0);
                break;
            case 3:
                current = (current.Item1, current.Item2, current.Item3 + 1, 0);
                break;
            case 4:
                current = (current.Item1, current.Item2, current.Item3, current.Item4 + 1);
                break;
            default:
                throw new NotImplementedException();
        }
        return TupleValues(current).Cast<int>().Where(x => x>0).Joined(".");
    }
}