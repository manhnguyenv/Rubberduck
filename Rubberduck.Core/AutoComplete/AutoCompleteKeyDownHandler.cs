﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteKeyDownHandler
    {
        private readonly Func<AutoCompleteSettings> _getSettings;
        private readonly Func<List<SelfClosingPair>> _getClosingPairs;
        private readonly Func<SelfClosingPairCompletionService> _getClosingPairCompletion;
        private readonly Func<ICodeStringPrettifier> _getPrettifier;

        public AutoCompleteKeyDownHandler(Func<AutoCompleteSettings> getSettings, Func<List<SelfClosingPair>> getClosingPairs, Func<SelfClosingPairCompletionService> getClosingPairCompletion, Func<ICodeStringPrettifier> getPrettifier)
        {
            _getSettings = getSettings;
            _getClosingPairs = getClosingPairs;
            _getClosingPairCompletion = getClosingPairCompletion;
            _getPrettifier = getPrettifier;
        }

        public void Run(ICodeModule module, Selection pSelection, AutoCompleteEventArgs e)
        {
            var currentContent = module.GetLines(pSelection);
            HandleSmartConcat(e, pSelection, currentContent, module);
            if (e.Handled) { return; }

            HandleSelfClosingPairs(e, module, pSelection);
            if (e.Handled) { return; }

            //HandleSomethingElse(?)
            //if (e.Handled) { return; }
        }

        /// <summary>
        /// Adds a line continuation when {ENTER} is pressed inside a string literal.
        /// </summary>
        private void HandleSmartConcat(AutoCompleteEventArgs e, Selection pSelection, string currentContent, ICodeModule module)
        {
            var shouldHandle = _getSettings().EnableSmartConcat &&
                               e.Character == '\r' &&
                               IsInsideStringLiteral(pSelection, ref currentContent);

            var lastIndexLeftOfCaret = currentContent.Length > 2 ? currentContent.Substring(0, pSelection.StartColumn - 1).LastIndexOf('"') : 0;
            if (shouldHandle && lastIndexLeftOfCaret > 0)
            {
                var indent = currentContent.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);
                var code = $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";

                if (e.ControlDown)
                {
                    code = $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & vbNewLine & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";
                }

                module.ReplaceLine(pSelection.StartLine, code);
                using (var pane = module.CodePane)
                {
                    pane.Selection = new Selection(pSelection.StartLine + 1, indent + currentContent.Substring(pSelection.StartColumn - 2).Length);
                    e.Handled = true;
                }
            }
        }

        private void HandleSelfClosingPairs(AutoCompleteEventArgs e, ICodeModule module, Selection pSelection)
        {
            if (!pSelection.IsSingleCharacter)
            {
                return;
            }

            var currentCode = e.CurrentLine;
            var currentSelection = e.CurrentSelection;

            var original = new CodeString(currentCode, new Selection(0, currentSelection.EndColumn - 1), new Selection(pSelection.StartLine, 1));

            var prettifier = _getPrettifier();
            var scp = _getClosingPairCompletion();

            foreach (var selfClosingPair in _getClosingPairs())
            {
                CodeString result;
                if (e.Character == '\b' && pSelection.StartColumn > 1)
                {
                    result = scp.Execute(selfClosingPair, original, Keys.Back);
                }
                else
                {
                    result = scp.Execute(selfClosingPair, original, e.Character);
                }

                if (!result?.Equals(default) ?? false)
                {
                    using (var pane = module.CodePane)
                    {
                        var prettified = prettifier.Prettify(module, original);
                        if (e.Character == '\b' && pSelection.StartColumn > 1)
                        {
                            result = scp.Execute(selfClosingPair, prettified, Keys.Back);
                        }
                        else
                        {
                            result = scp.Execute(selfClosingPair, prettified, e.Character);
                        }

                        module.DeleteLines(result.SnippetPosition);
                        module.InsertLines(result.SnippetPosition.StartLine, result.Code);

                        var reprettified = module.GetLines(result.SnippetPosition);
                        var offByOne = result.Code != reprettified;
                        Debug.Assert(!offByOne || reprettified.Length - result.Code.Length == 1, "Prettified code is off by more than one character.");

                        var finalSelection = new Selection(result.SnippetPosition.StartLine, result.CaretPosition.StartColumn + 1)
                            .ShiftRight(offByOne ? 1 : 0);
                        pane.Selection = finalSelection;
                        e.Handled = true;
                        return;
                    }
                }
            }
        }

        private bool IsInsideStringLiteral(Selection pSelection, ref string currentContent)
        {
            if (!currentContent.Substring(pSelection.StartColumn - 1).Contains("\"") ||
                currentContent.StripStringLiterals().HasComment(out _))
            {
                return false;
            }

            var zSelection = pSelection.ToZeroBased();
            var leftOfCaret = currentContent.Substring(0, zSelection.StartColumn);
            var rightOfCaret = currentContent.Substring(Math.Min(zSelection.StartColumn + 1, currentContent.Length - 1));
            if (!rightOfCaret.Contains("\""))
            {
                // the string isn't terminated, but VBE would terminate it here.
                currentContent += "\"";
                rightOfCaret += "\"";
            }

            // odd number of double quotes on either side of the caret means we're inside a string literal, right?
            return (leftOfCaret.Count(c => c.Equals('"')) % 2) != 0 &&
                   (rightOfCaret.Count(c => c.Equals('"')) % 2) != 0;
        }
    }
}