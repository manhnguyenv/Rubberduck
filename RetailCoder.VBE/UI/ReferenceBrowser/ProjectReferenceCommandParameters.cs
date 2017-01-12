using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ProjectReferenceCommandParameters
    {
        private readonly string _path;

        public ProjectReferenceCommandParameters(string path)
        {
            _path = path;
        }

        public string Path { get { return _path; } }
        public IReference Reference { get; set; }
    }
}