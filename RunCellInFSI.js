/// RunCellInFSI.js
///
/// Visual Studio macro for selecting the current "cell" (a textually
/// delimited document fragment), and sending it to F# interactive.
///
/// To use, install "Macros for Visual Studio" extension. Open the
/// macro editor and add this macro. Right click, assign shortcut,
/// assign custom keyboard shortcut. Pick `Tools.MacroCommand1`.
/// Assign `Ctrl+Alt+/`.
///
/// For latest versions, check https://github.com/t0yv0/fsharp-vs-macros

(function () {
    var se = dte.ActiveDocument.Selection
    var editPoint = se.ActivePoint.CreateEditPoint()
    function line(n) {
        return editPoint.GetLines(n, n + 1)
    }
    function lineIsSeparator(n) {
        return line(n).substr(0, 3) == '//-'
    }
    var here = se.ActivePoint.Line
    var top = here
    while (!lineIsSeparator(top)) { top-- }
    var bottom = here
    while (!lineIsSeparator(bottom)) { bottom++ }
    var originalLine = se.CurrentLine
    var originalColumn = se.CurrentColumn
    se.MoveTo(top + 1, 1)
    se.MoveTo(bottom, 1, true)
    dte.ExecuteCommand("EditorContextMenus.CodeWindow.ExecuteInInteractive")
    se.MoveToLineAndOffset(originalLine, originalColumn)
})();
