using System;
using System.Linq;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.AddRemoveReferences;
using AddRemoveReferencesViewModel = Rubberduck.UI.AddRemoveReferences.AddRemoveReferencesViewModel;

namespace Rubberduck.UI.Command
{
    public class AddRemoveReferencesCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public AddRemoveReferencesCommand(RubberduckParserState state, IMessageBox messageBox) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _messageBox = messageBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Projects.Any();
        }

        protected override void OnExecute(object parameter)
        {
            using (var project = _state.ActiveProject)
            {
                var service = new ProjectReferencesService(project);
                var finder = new RegisteredLibraryFinderService(Environment.Is64BitProcess);
                var fileDialog = new OpenFileDialog();

                var vm = new AddRemoveReferencesViewModel(finder, service, fileDialog, _messageBox);

                var presenter = new AddRemoveReferencesPresenter();
                presenter.Show(vm);
            }
        }
    }
}