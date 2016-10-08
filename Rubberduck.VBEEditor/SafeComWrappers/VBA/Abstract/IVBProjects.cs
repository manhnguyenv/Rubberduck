using System;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract
{
    public interface IVBProjects : ISafeComWrapper, IComCollection<IVBProject>, IEquatable<IVBProjects>
    {
        IVBE VBE { get; }
        IVBE Parent { get; }
        IVBProject Add(ProjectType type);
        IVBProject Open(string path);
        void Remove(IVBProject project);
    }
}