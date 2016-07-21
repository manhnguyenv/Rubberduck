using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ReferenceBrowserViewModel : ViewModelBase
    {
        private readonly VBE _vbe;
        private readonly IRegisteredCOMLibraryService _service;
        private readonly IMessageBox _messageBox;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _vbaProjectReferences;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _registeredComReferences;

        private string _filter;

        public ReferenceBrowserViewModel(VBE vbe, IRegisteredCOMLibraryService service, IOpenFileDialogFactory pickerFactory, IMessageBox messageBox)
        {
            _vbe = vbe;
            _service = service;
            _messageBox = messageBox;

            _registeredComReferences = new ObservableCollection<RegisteredLibraryViewModel>();

            ComReferences = new CollectionViewSource {Source = _registeredComReferences}.View;
            ComReferences.SortDescriptions.Add(new SortDescription("CanRemoveReference", ListSortDirection.Ascending));
            ComReferences.SortDescriptions.Add(new SortDescription("IsActiveProjectReference", ListSortDirection.Descending));
            ComReferences.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));

            _vbaProjectReferences = new ObservableCollection<RegisteredLibraryViewModel>();
            VbaProjectReferences = new CollectionViewSource {Source = _vbaProjectReferences }.View;

            BuildTypeLibraryReferenceViewModels();
            BuildVbaProjectReferenceViewModels();

            _addVBReferenceCommand = new AddReferenceFromFileCommand(vbe);
        }

        public ICollectionView ComReferences { get; private set; }

        public ICollectionView VbaProjectReferences { get; private set; }

        public string ComReferencesFilter
        {
            get { return _filter; }
            set
            {
                if (value == _filter)
                {
                    return;
                }
                _filter = value;
                FilterComReferences();
                OnPropertyChanged();
            }
        }

        private readonly ICommand _addVBReferenceCommand;
        public ICommand AddVbaProjectReferenceCommand { get {return _addVBReferenceCommand; } }

        private void AddVbaReference(string path)
        {
            try
            {
                var reference = _vbe.ActiveVBProject.References.AddFromFile(path);
                CreateViewModelForVbaProjectReference(reference);
            }
            catch (COMException)
            {
                // todo: localize
                _messageBox.Show(string.Format("Could not add reference to file or library '{0}'.", path));
            }
        }

        private void FilterComReferences()
        {
            if (string.IsNullOrWhiteSpace(_filter))
            {
                ComReferences.Filter = null;
            }
            else
            {
                ComReferences.Filter = o => 
                    ((RegisteredLibraryViewModel) o).Name.ToLowerInvariant()
                    .Contains(_filter.ToLowerInvariant());
            }
        }

        private void BuildTypeLibraryReferenceViewModels()
        {
            var viewModels = _service.GetAll()
                .Select(item => new RegisteredLibraryViewModel(item, _vbe.ActiveVBProject));

            foreach (var vm in viewModels)
            {
                _registeredComReferences.Add(vm);
            }
        }

        private void BuildVbaProjectReferenceViewModels()
        {
            var vbaReferences = _vbe.ActiveVBProject.References
                .OfType<Reference>()
                .Where(r => r.Type == vbext_RefKind.vbext_rk_Project);

            foreach (var reference in vbaReferences)
            {
                CreateViewModelForVbaProjectReference(reference);
            }
        }

        private void CreateViewModelForVbaProjectReference(Reference reference)
        {
            var model = new VbaProjectReferenceModel(reference);
            var vm = new RegisteredLibraryViewModel(model, _vbe.ActiveVBProject);
            _vbaProjectReferences.Add(vm);
        }
    }
}
