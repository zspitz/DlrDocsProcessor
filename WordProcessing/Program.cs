using System;
using System.CommandLine;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using ZSpitz.Util;

var arg = new Argument<string>();
var command = new RootCommand() { arg };
var filename = command.Parse(args).ValueForArgument(arg);

var wdApp = new Application();
wdApp.Visible = true;

var wdDoc = wdApp.Documents.Open(filename);
wdDoc.Styles.Add("Source Code", WdStyleType.wdStyleTypeParagraph);
wdDoc.Styles.Add("Verbatim Char", WdStyleType.wdStyleTypeCharacter);

void executeFind(Action<Find> findBuilder, Action<Microsoft.Office.Interop.Word.Range> onFind) {
    var rng = wdDoc!.Content;
    rng.Collapse(WdCollapseDirection.wdCollapseStart);
    rng.Select();
    var find = rng.Find;
    find.ClearFormatting();
    find.Text = "";
    find.Forward = true;
    find.Wrap = WdFindWrap.wdFindStop;
    find.Format = true;
    findBuilder(find);
    while (find.Execute()) {
        rng.Select();
        onFind(rng);
        rng.Collapse(WdCollapseDirection.wdCollapseEnd);
    }
}

executeFind(
    find => find.Font.Name = "Consolas",
    rng => {
        var isCompleteParagraph = IsCompleteParagraph(rng);
        switch (isCompleteParagraph) {
            case true:
                rng.set_Style("Source Code");
                break;
            case false:
                rng.set_Style("Verbatim Char");
                break;
        }
    }
);

executeFind(
    find => find.set_Style("Code"),
    rng => {
        rng.set_Style("Source Code");
    }
);

wdDoc.Save();
wdDoc.Close();
wdDoc = null;
wdApp.Quit(true);
wdApp = null;

static bool? IsCompleteParagraph(Microsoft.Office.Interop.Word.Range rng) {
    var rngExpanded = rng.Duplicate;
    rngExpanded.Expand(WdUnits.wdParagraph);
    var (length, expandedLength) = (rng.Length(), rngExpanded.Length());
    if (length == 1) { return null; } // empty paragraph
    return (expandedLength - length).In(0, 1); // code paragraph, may or may not include the end-paragraph character
}

public static class Extensions {
    public static int Length(this Microsoft.Office.Interop.Word.Range rng) => rng.End - rng.Start;
}