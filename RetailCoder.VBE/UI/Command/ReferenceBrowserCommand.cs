using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.UI.ReferenceBrowser;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Reference Browser window.
    /// </summary>
    [ComVisible(false)]
    public class ReferenceBrowserCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IReferenceBrowserPresenter _presenter;

        public ReferenceBrowserCommand(IVBE vbe, IReferenceBrowserPresenter presenter) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _presenter = presenter;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _presenter.Show(_vbe.ActiveVBProject);
        }
    }

    public interface IReferenceBrowserPresenter
    {
        ReferenceBrowserViewModel Show(IVBProject project);
    }

    public class ReferenceBrowserPresenter : IReferenceBrowserPresenter
    {
        private readonly IRegisteredCOMLibraryService _service;

        public ReferenceBrowserPresenter(IRegisteredCOMLibraryService service)
        {
            _service = service;
       }

        public ReferenceBrowserViewModel Show(IVBProject project)
        {
            if (project.IsWrappingNullReference)
            {
                return null;
            }

            var viewModel = new ReferenceBrowserViewModel(project, _service);
            using (var window = new ReferenceBrowserWindow(viewModel))
            {
                if (window.ShowDialog() == DialogResult.OK)
                {
                    // todo: sync viewmodel with project's references
                }
            }

            return viewModel;
        }
    }
}
