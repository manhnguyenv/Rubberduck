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
        public void Show(AddRemoveReferencesViewModel viewModel)
        {
            using (var dialog = new AddRemoveReferencesDialog(viewModel))
            {
                dialog.ShowDialog();
            }
        }
    }
}
