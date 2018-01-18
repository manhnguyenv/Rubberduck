using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.ComReflection;
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
            AvailableProjects = GetAvailableProjects(project);
        }

        private IEnumerable<ReferenceModel> GetAvailableProjects(IVBProject project)
        {
            using (var projects = project.Collection)
            {
                foreach (var vbProject in projects)
                {
                    if (vbProject.ProjectId != project.ProjectId)
                    {
                        yield return new ReferenceModel(vbProject);
                    }
                    vbProject.Dispose();
                }
            }
        }

        public IEnumerable<ReferenceModel> AvailableProjects { get; }

        public ReferenceModel GetLibraryInfo(string path)
        {
            var info = ReferencedDeclarationsCollector.LoadTypeLibInfo(path);
            if (info == null)
            {
                return null;
            }

            return new ReferenceModel(info);
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
                        reference.Dispose();
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
                foreach (var reference in model.Where(m => !m.IsBuiltIn && m.IsSelected).OrderBy(m => m.Priority))
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
                    reference.Dispose();
                }
                references.EnableEvents = true;
            }
            Logger.Trace("Project references cleared.");
        }
    }
}