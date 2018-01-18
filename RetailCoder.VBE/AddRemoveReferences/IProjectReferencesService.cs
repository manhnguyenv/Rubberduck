using System.Collections.Generic;

namespace Rubberduck.AddRemoveReferences
{
    public interface IProjectReferencesService
    {
        ReferenceModel GetLibraryInfo(string path);

        IEnumerable<ReferenceModel> References { get; }
        IEnumerable<ReferenceModel> AvailableProjects { get; }
        void Apply(IEnumerable<ReferenceModel> model);
    }
}