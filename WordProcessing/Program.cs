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
    if (length == 1) { continue; } // empty paragraph
    if ((expandedLength - length).In(0,1) ) { // code paragraph, may or may not include the end-paragraph character
        rng.set_Style("Source Code");
    } else {
        rng.set_Style("Verbatim Char");
    }
    rng.Collapse(WdCollapseDirection.wdCollapseEnd);
}
wdDoc.Save();
wdDoc.Close();
wdDoc = null;
wdApp.Quit(true);
wdApp = null;

public static class Extensions {
    public static int Length(this Range rng) => rng.End - rng.Start;
}