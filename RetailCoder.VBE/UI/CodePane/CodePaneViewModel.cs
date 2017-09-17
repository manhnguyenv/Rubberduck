using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Properties;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodePane
{
    public class CodePaneViewModel : INotifyPropertyChanged
    {
        private readonly ICodeModule _module;

        public CodePaneViewModel(ICodeModule module)
        {
            _module = module;
        }

        public string Content
        {
            get => _module.Content();
            set
            {
                _module.Clear();
                _module.InsertLines(1, value);
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
