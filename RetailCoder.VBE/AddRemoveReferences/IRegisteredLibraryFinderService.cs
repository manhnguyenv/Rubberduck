using System.Collections.Generic;

namespace Rubberduck.AddRemoveReferences
{
    public interface IRegisteredLibraryFinderService
    {
        IEnumerable<ReferenceModel> FindRegisteredLibraries();
    }
}