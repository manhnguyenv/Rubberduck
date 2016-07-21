using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.UI.ReferenceBrowser;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Reference Browser window.
    /// </summary>
    [ComVisible(false)]
    public class ReferenceBrowserCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly IRegisteredCOMLibraryService _service;
        private readonly IMessageBox _messageBox;
        private readonly IOpenFileDialogFactory _pickerFactory;

        public ReferenceBrowserCommand(VBE vbe, IRegisteredCOMLibraryService service, IMessageBox messageBox, IOpenFileDialogFactory pickerFactory) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _service = service;
            _messageBox = messageBox;
            _pickerFactory = pickerFactory;
        }

        protected override void ExecuteImpl(object parameter)
        {
            var viewModel = new ReferenceBrowserViewModel(_vbe, _service, _pickerFactory, _messageBox);
            using (var window = new ReferenceBrowserWindow(viewModel))
            {
                if (window.ShowDialog() == DialogResult.OK)
                {
                    // todo: sync viewmodel with active vbproject's references
                }
            }
        }
    }
}
