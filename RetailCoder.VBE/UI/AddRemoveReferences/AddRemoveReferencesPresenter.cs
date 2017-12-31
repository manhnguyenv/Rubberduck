using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.AddRemoveReferences;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.AddRemoveReferences
{
    public class AddRemoveReferencesPresenter
    {
        public void Show(IVBProject project, bool is64BitHost)
        {
            var service = new ProjectReferencesService(project);
            var finder = new RegisteredLibraryFinderService(is64BitHost);
            var fileDialog = new OpenFileDialog();
            var vm = new AddRemoveReferencesViewModel(finder, service, fileDialog);
            using (var dialog = new AddRemoveReferencesDialog(vm))
            {
                dialog.ShowDialog();
            }
        }
    }
}
