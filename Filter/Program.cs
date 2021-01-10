using System.Collections.Immutable;
using PandocFilters;
using PandocFilters.Ast;
using ZSpitz.Util;
using System.Linq;
using System.CommandLine;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;

var command = new RootCommand() {
    new Option<int>("--pass", () => -1)
};
var result = command.Parse(args);
var pass = result.ValueForOption<int>("--pass");

var headerCount = 0;

if (pass == -1) {
    var headersJsonGenerator = new DelegateVisitor();
    var headers = new Dictionary<int, string>();
    headersJsonGenerator.Add((Pandoc pandoc) => {
        foreach (var block in pandoc.Blocks) {
            if (block.IsT9) {
                var header = block.AsT9;
                if (header.Level==1) {
                    headerCount += 1;
                    headers.Add(headerCount, header.Attr.Identifier);
                }
            }
        }
        return pandoc;
    });
    Filter.Run(headersJsonGenerator);

    var jsonString = JsonSerializer.Serialize(headers);
    File.WriteAllText("headers.json", jsonString);
    return;
}

var splitter = new DelegateVisitor();
splitter.Add((Pandoc pandoc) => {
    return pandoc with
    {
        Blocks = pandoc.Blocks.Select(block => {
            if (block.IsT9 && block.AsT9.Level == 1) {
                headerCount += 1;
            }
            return (block, headerCount);
        })
        .WhereT((_, headerCount) => headerCount==pass)
        .SelectT((block, _) => block)
        .ToImmutableList()
    };
});


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

Filter.Run(splitter, visitor);