using System;
using System.Linq;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.AddRemoveReferences;

namespace Rubberduck.UI.Command
{
    public class AddRemoveReferencesCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public AddRemoveReferencesCommand(RubberduckParserState state) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Projects.Any();
        }

        protected override void OnExecute(object parameter)
        {
            var presenter = new AddRemoveReferencesPresenter();
            using (var project = _state.ActiveProject)
            {
                presenter.Show(project, Environment.Is64BitProcess);
            }
        }
    }
}