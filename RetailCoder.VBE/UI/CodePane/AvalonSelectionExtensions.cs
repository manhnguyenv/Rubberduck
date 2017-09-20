using Rubberduck.VBEditor;

namespace Rubberduck.UI.CodePane
{
    public static class AvalonSelectionExtensions
    {
        public static Selection ToRubberduckSelection(this ICSharpCode.AvalonEdit.Editing.Selection selection)
        {
            return new Selection(selection.StartPosition.Line + 1, selection.StartPosition.Column + 1, selection.EndPosition.Line + 1, selection.EndPosition.Column + 1);
        }
    }
}