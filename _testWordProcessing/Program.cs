using Microsoft.Office.Interop.Word;
using static System.IO.Path;
using static System.IO.Directory;
using static System.Reflection.Assembly;
using static System.Console;
using ZSpitz.Util;

var filename = Combine(GetFullPath($"{GetParent(GetExecutingAssembly().Location)}"), "dlr-overview.docx");

var wdApp = new Application();
wdApp.Visible = true;

var wdDoc = wdApp.Documents.Open(filename, ReadOnly: true);
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
    // if entire paragraph has the same font
    //      set the style to Source Code
    // otherwise
    //      insert 
    //rng.set_Style("Source Code");
    var rng2 = rng.Duplicate;
    rng2.Expand(WdUnits.wdParagraph);
    WriteLine(rng.Text);
    WriteLine((rng.Start, rng.End, rng.Length()));
    WriteLine(rng2.Text);
    WriteLine((rng2.Start, rng2.End,rng2.Length()));
    if (rng2.Length() == 1) {
        WriteLine("ignore - empty paragraph");
    } else if ((rng2.Length() - rng.Length()).In(0,1)) {
        WriteLine("paragraph style");
    } else {
        WriteLine("character style");
    }
    WriteLine();
}
wdDoc = null;
wdApp.Quit(false);
wdApp = null;

public static class Extensions {
    public static int Length(this Range rng) => rng.End - rng.Start;
}