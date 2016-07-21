using System.Collections.Generic;
using System.Linq;
using Kavod.ComReflection;

namespace Rubberduck.UI.ReferenceBrowser
{
    public interface IRegisteredCOMLibraryService
    {
        IEnumerable<RegisteredLibraryModel> GetAll();
    }

    public class RegisteredCOMLibraryService : IRegisteredCOMLibraryService
    {
        public IEnumerable<RegisteredLibraryModel> GetAll()
        {
            return LibraryRegistration.GetRegisteredTypeLibraryEntries()
                                      .Select(library => new RegisteredLibraryModel(library))
                                      .ToList();
        }
    }
}