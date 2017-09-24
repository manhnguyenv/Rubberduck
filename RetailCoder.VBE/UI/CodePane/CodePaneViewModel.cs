using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using NLog;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodePane
{
    public class CodePaneViewModel : INotifyPropertyChanged
    {
        private readonly ICodeModule _module;

        public CodePaneViewModel(ICodeModule module, ReparseCommand reparseCommand)
        {
            _module = module;
            RefreshCommand = reparseCommand;
            // todo: implement some module-scoped inspection results pane
            ShowInspectionResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), o => {}, o => false);
        }

        public ICommand RefreshCommand { get; }
        public ICommand ShowInspectionResultsCommand { get; }

        // todo: figure out a way to get filtered IInspector results in here
        public ObservableCollection<IInspectionResult> InspectionResults { get; }

        // todo: populate for LHS combobox
        public ObservableCollection<Declaration> MemberProviders { get; }
        public Declaration CurrentMemberProvider { get; }

        // todo: populate for RHS combobox
        public ObservableCollection<Declaration> Members { get; }
        public Declaration CurrentMember { get; }

        private string _content;
        public string Content
        {
            get => _content;
            private set
            {
                _content = value;
                OnPropertyChanged();
            }
        }

        private string _inspectionStatus;
        public string InspectionStatus
        {
            get => _inspectionStatus;
            private set
            {
                _inspectionStatus = value;
                OnPropertyChanged();
            }
        }

        private string _statusBarText;
        public string StatusBarText
        {
            get => _statusBarText;
            set
            {
                _statusBarText = value;
                OnPropertyChanged();
            }
        }

        private string _lineCountText;
        public string LineCountText
        {
            get => _lineCountText;
            set
            {
                _lineCountText = value;
                OnPropertyChanged();
            }
        }

        private bool _isDirty;
        public bool IsDirty
        {
            get => _isDirty;
            private set
            {
                _isDirty = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Clears the module's entire content.
        /// </summary>
        /// <param name="setDirty">Indicates if this action should mark the code pane as modified.</param>
        public void Clear(bool setDirty = true)
        {
            _module.Clear();
            Content = string.Empty;
            if (setDirty)
            {
                IsDirty = true;
            }
            OnLineCountChanged();
        }

        /// <summary>
        /// Clears the module's entire content.
        /// </summary>
        /// <param name="content">The new/replacement content.</param>
        /// <param name="setDirty">Indicates if this action should mark the code pane as modified.</param>
        public void Clear(string content, bool setDirty = true)
        {
            Clear(setDirty);
            _module.AddFromString(content);
            Content = content;
            OnLineCountChanged();
        }

        /// <summary>
        /// Writes the specified content at the specified 1-based codepane location.
        /// </summary>
        /// <param name="content">The content to be written.</param>
        /// <param name="selection">The codepane location.</param>
        public void Write(string content, Selection selection)
        {
            if (selection.LineCount > 1)
            {
                _module.DeleteLines(selection);
                _module.InsertLines(selection.StartLine, content);
            }
            else
            {
                _module.ReplaceLine(selection.StartLine, content);
            }

            IsDirty = true;
            OnLineCountChanged();
            Content = _module.Content();
        }

        private void OnLineCountChanged()
        {
            LineCountText = string.Format(RubberduckUI.Editor_LineCount, _module.CountOfLines); 
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
