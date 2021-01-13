using System.Collections.Immutable;
using PandocFilters;
using PandocFilters.Ast;
using ZSpitz.Util;
using System.Linq;
using System.CommandLine;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;
using System;
using static ZSpitz.Util.Functions;
using Endless;
using static System.Linq.Enumerable;
using System;

var command = new RootCommand() {
    new Option<int>("--pass", () => -1)
};
var result = command.Parse(args);
// there are multiple passes on each document; enables splitting the document based on its' first-level headers
var pass = result.ValueForOption<int>("--pass");

var headerCount = 0;

if (pass == -1) {
    var firstPass = new DelegateVisitor();

    // generate headers.json
    var toplevelHeaders = new Dictionary<int, string>();
    var tocEntries = new List<(int level, ImmutableList<Inline> text, string url)>();

    firstPass.Add((Pandoc pandoc) => {
        foreach (var block in pandoc.Blocks) {
            if (block.TryPickT9(out var header, out var _) && header.Level == 1) {
                headerCount += 1;
                toplevelHeaders.Add(headerCount, header.Attr.Identifier);
            } else if (headerCount == 0 && !toplevelHeaders.ContainsKey(0)) {
                toplevelHeaders.Add(0, "frontmatter");
                tocEntries.Add(
                    0,
                    ((Inline)new Str("Frontmatter")).Yield().ToImmutableList(),
                    "forntmatter.md"
                );
            }

            if (header is { }) {
                tocEntries.Add(header.Level, header.Text, $"{toplevelHeaders[headerCount]}.md#{header.Attr.Identifier}");
            }
        }

        // generate the TOC for the sidebar
        return pandoc with
        {
            Blocks = ((Block)new LineBlock(
                tocEntries.SelectT((level, text, url) => {
                    var repetitions = (level == 0 ? 0 : level - 1) * 2;
                    var inlines = new List<Inline>();
                    Repeat<Inline>(new RawInline("markdown", "&nbsp;"), repetitions).AddRangeTo(inlines);
                    inlines.Add(new Link(
                        Attr.Empty,
                        text,
                        (url, "")
                    ));
                    return inlines.ToImmutableList();
                }).ToImmutableList()
            )).Yield().ToImmutableList()
        };
    });
    Filter.Run(new HierarchyNumberGenerator(), firstPass);

    var jsonString = JsonSerializer.Serialize(toplevelHeaders);
    File.WriteAllText("headers.json", jsonString);
    return;
}

// excludes parts of the document from the output based on the current pass
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

// rewrite headers to embedded HTML with explicitly defined ids
visitor.Add((Pandoc pandoc) =>
    pandoc with
    {
        Blocks = pandoc.Blocks.Select<Block, Block>(block => {
            if (block.IsT9) {
                var header = block.AsT9;
                var id = header.Attr.Identifier;
                return new Para(
                    new Inline[] { new RawInline("markdown", @$"<h{header.Level} id=""{header.Attr.Identifier}"">") }
                        .Concat(header.Text)
                        .ConcatOne(new RawInline("markdown", $"</h{header.Level}>"))
                        .ToImmutableList()
                );
            }
            return block;
        }).ToImmutableList()
    });

// remove height and width from images, so they'll output as standard markdown images, instead of HTML img elements
visitor.Add((Image img) =>
    img with
    {
        Attr = img.Attr with
        {
            KeyValuePairs = img.Attr.KeyValuePairs.WhereT((key, _) => key.NotIn("height", "width")).ToImmutableList()
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

Filter.Run(new HierarchyNumberGenerator(), splitter, visitor);

public class HierarchyNumberGenerator : VisitorBase {
    private (int, int, int, int) current { get; set; }
    private string Next(Header header) {
        current = header.Level switch {
            1 => (current.Item1 + 1, 0, 0, 0),
            2 => (current.Item1, current.Item2 + 1, 0, 0),
            3 => (current.Item1, current.Item2, current.Item3 + 1, 0),
            4 => (current.Item1, current.Item2, current.Item3, current.Item4 + 1),
            _ => throw new NotImplementedException(),
        };
        return TupleValues(current).Cast<int>().Where(x => x > 0).Joined(".");
    }

    public override Header VisitHeader(Header header) {
        header = header with
        {
            Text = new Inline[] {
                    new Str(Next(header)),
                    new Space()
            }.Concat(header.Text).ToImmutableList()
        };
        return base.VisitHeader(header);
    }
}
