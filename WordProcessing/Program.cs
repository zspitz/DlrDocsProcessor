using System.CommandLine;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using ZSpitz.Util;

Debugger.Launch();

var arg = new Argument<string>();
var command = new RootCommand() { arg };
var filename = command.Parse(args).ValueForArgument(arg);

var wdApp = new Application();
wdApp.Visible = true;

var wdDoc = wdApp.Documents.Open(filename);
wdDoc.Styles.Add("Source Code", WdStyleType.wdStyleTypeParagraph);

var rng = wdDoc.Content;
rng.Collapse(WdCollapseDirection.wdCollapseStart);
rng.Select();
var find = rng.Find;
find.ClearFormatting();
find.Text = "";
find.Forward = true;
find.Wrap = WdFindWrap.wdFindStop;
find.Format = true;
find.Font.Name = "Consolas";
while (find.Execute()) {
    var rngExpanded = rng.Duplicate;
    rngExpanded.Expand(WdUnits.wdParagraph);
    var (length, expandedLength) = (rng.Length(), rngExpanded.Length());
    var (text, expandedText) = (rng.Text, rngExpanded.Text);
    if (length == 1) { continue; } // empty paragraph
    if ((expandedLength - length).In(0,1) ) { // code paragraph, may or may not include the end-paragraph character
        rng.set_Style("Source Code");
    }

    // if entire paragraph has the same font
    //      set the style to Source Code (done)
    // otherwise
    //      insert $_$_$_$ before each word
    //      rely on visitor to convert a given set of words to code
    rng.Collapse(WdCollapseDirection.wdCollapseEnd);
}
wdDoc.Save();
wdDoc.Close();
wdDoc = null;
wdApp.Quit(true);
wdApp = null;


// TODO not implemented
/*
visitor:
get the minimun number of spaces at the beginning of a paragraph
remove that number of spaces from the beginning of each paragraph
 */

public static class Extensions {
    public static int Length(this Range rng) => rng.End - rng.Start;
}