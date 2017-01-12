using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ReferenceBrowserViewModel : ViewModelBase
    {
        private readonly IVBProject _project;
        private readonly IRegisteredCOMLibraryService _service;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _vbaProjectReferences;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _registeredComReferences;

        private string _filter;

        public ReferenceBrowserViewModel(IVBProject project, IRegisteredCOMLibraryService service)
        {
            _project = project;
            _service = service;

            _registeredComReferences = new ObservableCollection<RegisteredLibraryViewModel>();

            ComReferences = new CollectionViewSource {Source = _registeredComReferences}.View;
            ComReferences.SortDescriptions.Add(new SortDescription("CanRemoveReference", ListSortDirection.Ascending));
            ComReferences.SortDescriptions.Add(new SortDescription("IsActiveProjectReference", ListSortDirection.Descending));
            ComReferences.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));

            _vbaProjectReferences = new ObservableCollection<RegisteredLibraryViewModel>();
            VbaProjectReferences = new CollectionViewSource {Source = _vbaProjectReferences }.View;

            BuildTypeLibraryReferenceViewModels();
            BuildVbaProjectReferenceViewModels();

            _addVBReferenceCommand = new AddReferenceFromFileCommand(_project.VBE);
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
        public ICommand AddVbaProjectReferenceCommand { get { return _addVBReferenceCommand; } }

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
            var references = _service
                .GetAll()
                .Select(model => new RegisteredLibraryViewModel(model, _project));

            foreach (var reference in references)
            {
                _registeredComReferences.Add(reference);
            }
        }

        private void BuildVbaProjectReferenceViewModels()
        {
            var references = _project
                .References
                .Where(r => r.Type == ReferenceKind.Project)
                .Select(reference => new RegisteredLibraryViewModel(new RegisteredLibraryModel(reference), _project));

            foreach (var reference in references)
            {
                _vbaProjectReferences.Add(reference);
            }
        }
    }
}
