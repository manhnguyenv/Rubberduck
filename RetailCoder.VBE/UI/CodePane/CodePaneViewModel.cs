using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Properties;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodePane
{
    public class CodePaneViewModel : INotifyPropertyChanged
    {
        private readonly ICodeModule _module;

        public CodePaneViewModel(ICodeModule module)
        {
            _module = module;
            _inspectionStatus = "OK";
            _showLineNumbers = true;
        }

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

        private bool _showLineNumbers;
        public bool ShowLineNumbers
        {
            get => _showLineNumbers;
            set
            {
                _showLineNumbers = value;
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
            Content = _module.Content();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
