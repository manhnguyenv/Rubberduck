﻿using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class MainParseExceptionErrorListener : ParsePassExceptionErrorListener
    {
        public MainParseExceptionErrorListener(string moduleName, CodeKind codeKind)
        :base(moduleName, codeKind)
        { }

        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            // adding 1 to line, because line is 0-based, but it's 1-based in the VBE
            throw new MainParseSyntaxErrorException(msg, e, offendingSymbol, line, charPositionInLine + 1, ModuleName, CodeKind);
        }
    }
}
