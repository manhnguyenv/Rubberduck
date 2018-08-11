﻿using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class ExceptionErrorListener : RubberduckParseErrorListenerBase
    {
        public ExceptionErrorListener(CodeKind codeKind)
        :base(codeKind)
        { }

        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            // adding 1 to line, because line is 0-based, but it's 1-based in the VBE
            throw new SyntaxErrorException(msg, e, offendingSymbol, line, charPositionInLine + 1, CodeKind);
        }
    }
}
