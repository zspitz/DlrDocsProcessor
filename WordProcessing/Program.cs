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

wdDoc.Fields.Unlink(); // Convert all Visio diagrams to images

void executeFind(Action<Find> findBuilder, Action<Microsoft.Office.Interop.Word.Range> onFind) {
    var rng = wdDoc!.Content;
    rng.Collapse(WdCollapseDirection.wdCollapseStart);
    var find = rng.Find;
    find.ClearFormatting();
    find.Text = "";
    find.Forward = true;
    find.Wrap = WdFindWrap.wdFindStop;
    find.Format = true;
    findBuilder(find);
    while (find.Execute()) {
        if (Debugger.IsAttached) { rng.Select(); }
        onFind(rng);
        rng.Collapse(WdCollapseDirection.wdCollapseEnd);
    }
}

void applyCodeStyle(Microsoft.Office.Interop.Word.Range rng) {
    var isCompleteParagraph = IsCompleteParagraph(rng);
    switch (isCompleteParagraph) {
        case true:
            rng.set_Style("Source Code");
            break;
        case false:
            rng.set_Style("Verbatim Char");
            break;
    }
};

executeFind(
    find => find.Font.Name = "Consolas",
    applyCodeStyle
);

executeFind(
    find => find.Font.Name = "Courier New",
    applyCodeStyle
);

if (wdDoc.HasStyle("Code")) {
    executeFind(
        find => find.set_Style("Code"),
        applyCodeStyle
    );
}

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
    public static bool HasStyle(this Document doc, string styleName) {
        try {
            var s = doc.Styles[styleName];
            return true;
        } catch (Exception) {
            return false;
        }
    }
}