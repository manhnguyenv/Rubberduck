using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AddRemoveReferences
{
    public sealed class ProjectReferencesService : IProjectReferencesService
    {
        private readonly IVBProject _project;
        private static readonly ILogger Logger = LogManager.GetCurrentClassLogger();

        public ProjectReferencesService(IVBProject project)
        {
            _project = project;
        }

        public IEnumerable<ReferenceModel> References
        {
            get
            {
                using (var references = _project.References)
                {
                    var priority = 0;
                    foreach (var reference in references)
                    {
                        yield return new ReferenceModel(reference, priority);
                        priority++;
                    }
                }
            }
        }

        public void Apply(IEnumerable<ReferenceModel> model)
        {
            ClearReferences();
            using (var references = _project.References)
            {
                references.EnableEvents = false;
                foreach (var reference in model.Where(m => !m.IsBuiltIn).OrderBy(m => m.Priority))
                {
                    try
                    {
                        Logger.Trace($"Adding reference to {reference.Name}...");
                        references.AddFromFile(reference.FullPath);
                    }
                    catch (Exception e)
                    {
                        Logger.Error(e);
                    }
                }
                references.EnableEvents = true;
            }
        }

        private void ClearReferences()
        {
            using (var references = _project.References)
            {
                references.EnableEvents = false;
                foreach (var reference in references)
                {
                    if (!reference.IsBuiltIn)
                    {
                        references.Remove(reference);
                    }
                }
                references.EnableEvents = true;
            }
            Logger.Trace("Project references cleared.");
        }
    }
}