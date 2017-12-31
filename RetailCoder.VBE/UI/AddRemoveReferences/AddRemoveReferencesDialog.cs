using System.Windows.Forms;

namespace Rubberduck.UI.AddRemoveReferences
{
    public partial class AddRemoveReferencesDialog : Form
    {
        public AddRemoveReferencesDialog()
        {
            InitializeComponent();
        }

        public AddRemoveReferencesDialog(AddRemoveReferencesViewModel vm)
            : this()
        {
            addRemoveReferencesWindow1.DataContext = vm;
            RegisterViewModelEvents(vm);
        }

        private void RegisterViewModelEvents(AddRemoveReferencesViewModel vm)
        {
            vm.Accepted += ViewModelAccepted;
            vm.Cancelled += ViewModelCancelled;
        }

        private void UnregisterViewModelEvents(AddRemoveReferencesViewModel vm)
        {
            vm.Accepted -= ViewModelAccepted;
            vm.Cancelled -= ViewModelCancelled;
        }

        private void ViewModelAccepted(object sender, System.EventArgs e)
        {
            UnregisterViewModelEvents((AddRemoveReferencesViewModel)sender);
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ViewModelCancelled(object sender, System.EventArgs e)
        {
            UnregisterViewModelEvents((AddRemoveReferencesViewModel)sender);
            OnCancel();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                OnCancel();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void OnCancel()
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
