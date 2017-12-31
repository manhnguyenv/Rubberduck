using System.Collections.Generic;

namespace Rubberduck.AddRemoveReferences
{
    public interface IProjectReferencesService
    {
        IEnumerable<ReferenceModel> References { get; }
        void Apply(IEnumerable<ReferenceModel> model);
    }
}