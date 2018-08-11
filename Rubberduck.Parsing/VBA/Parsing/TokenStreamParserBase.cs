﻿using System;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Symbols.ParsingExceptions;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public abstract class TokenStreamParserBase : ITokenStreamParser
    {
        protected static ILogger Logger = LogManager.GetCurrentClassLogger();

        private readonly IParsePassErrorListenerFactory _sllErrorListenerFactory;
        private readonly IParsePassErrorListenerFactory _llErrorListenerFactory;

        public TokenStreamParserBase(IParsePassErrorListenerFactory sllErrorListenerFactory,
            IParsePassErrorListenerFactory llErrorListenerFactory)
        {
            _sllErrorListenerFactory = sllErrorListenerFactory;
            _llErrorListenerFactory = llErrorListenerFactory;
        }

        protected abstract IParseTree Parse(ITokenStream tokenStream, PredictionMode predictionMode, IParserErrorListener errorListener);

        public IParseTree Parse(string moduleName, CommonTokenStream tokenStream, CodeKind codeKind = CodeKind.SnippetCode,
            ParserMode parserMode = ParserMode.FallBackSllToLl)
        {
            switch (parserMode)
            {
                case ParserMode.FallBackSllToLl:
                    return ParseWithFallBack(moduleName, tokenStream, codeKind);
                case ParserMode.LlOnly:
                    return ParseLl(moduleName, tokenStream, codeKind);
                case ParserMode.SllOnly:
                    return ParseSll(moduleName, tokenStream, codeKind);
                default:
                    throw new ArgumentException(nameof(parserMode));
            }

        }

        private IParseTree ParseWithFallBack(string moduleName, CommonTokenStream tokenStream, CodeKind codeKind)
        {
            try
            {
                return ParseLl(moduleName, tokenStream, codeKind);
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                var message = $"SLL mode failed while parsing the {codeKind} version of module {moduleName} at symbol {syntaxErrorException.OffendingSymbol.Text} at L{syntaxErrorException.LineNumber}C{syntaxErrorException.Position}. Retrying using LL.";
                LogAndReset(tokenStream, message, syntaxErrorException);
                return ParseSll(moduleName, tokenStream, codeKind);
            }
            catch (Exception exception)
            {
                var message = $"SLL mode failed while parsing the {codeKind} version of module {moduleName}. Retrying using LL.";
                LogAndReset(tokenStream, message, exception);
                return ParseSll(moduleName, tokenStream, codeKind);
            }
        }

        private void LogAndReset(CommonTokenStream tokenStream, string logWarnMessage, Exception exception)
        {
            Logger.Warn(logWarnMessage);
            Logger.Debug(exception);
            tokenStream.Reset();
        }

        private IParseTree ParseLl(string moduleName, ITokenStream tokenStream, CodeKind codeKind)
        {
            var errorListener = _llErrorListenerFactory.Create(moduleName, codeKind);
            var tree = Parse(tokenStream, PredictionMode.Ll, errorListener);
            if (errorListener.HasPostponedException(out var exception))
            {
                throw exception;
            }
            return tree;
        }

        private IParseTree ParseSll(string moduleName, ITokenStream tokenStream, CodeKind codeKind)
        {
            var errorListener = _sllErrorListenerFactory.Create(moduleName, codeKind);
            var tree = Parse(tokenStream, PredictionMode.Sll, errorListener);
            if (errorListener.HasPostponedException(out var exception))
            {
                throw exception;
            }
            return tree;
        }
    }
}
