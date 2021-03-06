﻿// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NuGet.Configuration;
using NuGet.ProjectManagement;
using EnvDTEProject = EnvDTE.Project;

namespace NuGet.PackageManagement.VisualStudio
{
    [PartCreationPolicy(CreationPolicy.Shared)]
    [Export(typeof(ISolutionManager))]
    public class VSSolutionManager : ISolutionManager, IVsSelectionEvents
    {
        private readonly DTE _dte;
        private readonly SolutionEvents _solutionEvents;
        private readonly IVsMonitorSelection _vsMonitorSelection;
        private readonly uint _solutionLoadedUICookie;
        private readonly IVsSolution _vsSolution;
        private readonly NuGetAndEnvDTEProjectCache _nuGetAndEnvDTEProjectCache = new NuGetAndEnvDTEProjectCache();

        private bool _initialized;

        public INuGetProjectContext NuGetProjectContext { get; set; }

        public NuGetProject DefaultNuGetProject
        {
            get
            {
                Init();

                if (String.IsNullOrEmpty(DefaultNuGetProjectName))
                {
                    return null;
                }

                NuGetProject defaultNuGetProject;
                _nuGetAndEnvDTEProjectCache.TryGetNuGetProject(DefaultNuGetProjectName, out defaultNuGetProject);
                return defaultNuGetProject;
            }
        }

        public string DefaultNuGetProjectName { get; set; }

        public event EventHandler<NuGetProjectEventArgs> NuGetProjectAdded;
        public event EventHandler<NuGetProjectEventArgs> NuGetProjectRemoved;
        public event EventHandler<NuGetProjectEventArgs> NuGetProjectRenamed;

        public event EventHandler SolutionClosed;
        public event EventHandler SolutionClosing;
        public event EventHandler SolutionOpened;
        public event EventHandler SolutionOpening;

        public VSSolutionManager()
        {
            _dte = ServiceLocator.GetInstance<DTE>();
            _vsSolution = ServiceLocator.GetGlobalService<SVsSolution, IVsSolution>();
            _vsMonitorSelection = ServiceLocator.GetGlobalService<SVsShellMonitorSelection, IVsMonitorSelection>();

            // Keep a reference to SolutionEvents so that it doesn't get GC'ed. Otherwise, we won't receive events.
            _solutionEvents = _dte.Events.SolutionEvents;

            // can be null in unit tests
            if (_vsMonitorSelection != null)
            {
                Guid solutionLoadedGuid = VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_guid;
                _vsMonitorSelection.GetCmdUIContextCookie(ref solutionLoadedGuid, out _solutionLoadedUICookie);

                uint cookie;
                int hr = _vsMonitorSelection.AdviseSelectionEvents(this, out cookie);
                ErrorHandler.ThrowOnFailure(hr);
            }

            _solutionEvents.BeforeClosing += OnBeforeClosing;
            _solutionEvents.AfterClosing += OnAfterClosing;
            _solutionEvents.ProjectAdded += OnEnvDTEProjectAdded;
            _solutionEvents.ProjectRemoved += OnEnvDTEProjectRemoved;
            _solutionEvents.ProjectRenamed += OnEnvDTEProjectRenamed;
        }

        public NuGetProject GetNuGetProject(string nuGetProjectSafeName)
        {
            if (string.IsNullOrEmpty(nuGetProjectSafeName))
            {
                throw new ArgumentException(ProjectManagement.Strings.Argument_Cannot_Be_Null_Or_Empty, "nuGetProjectSafeName");
            }

            Init();

            NuGetProject nuGetProject = null;
            // NuGetAndEnvDTEProjectCache could be null when solution is not open.
            if (_nuGetAndEnvDTEProjectCache != null)
            {
                _nuGetAndEnvDTEProjectCache.TryGetNuGetProject(nuGetProjectSafeName, out nuGetProject);
            }
            return nuGetProject;
        }

        // Return short name if it's non-ambiguous.
        // Return CustomUniqueName for projects that have ambigous names (such as same project name under different solution folder)
        // Example: return Folder1/ProjectA if there are both ProjectA under Folder1 and Folder2
        public string GetNuGetProjectSafeName(NuGetProject nuGetProject)
        {
            if (nuGetProject == null)
            {
                throw new ArgumentNullException("nuGetProject");
            }

            Init();

            // Try searching for simple names first
            string name = nuGetProject.GetMetadata<string>(NuGetProjectMetadataKeys.Name);
            if (GetNuGetProject(name) == nuGetProject)
            {
                return name;
            }

            return NuGetProject.GetUniqueNameOrName(nuGetProject);
        }

        public Project GetDTEProject(string nuGetProjectSafeName)
        {
            if (string.IsNullOrEmpty(nuGetProjectSafeName))
            {
                throw new ArgumentException(ProjectManagement.Strings.Argument_Cannot_Be_Null_Or_Empty, "nuGetProjectSafeName");
            }

            Init();

            Project dteProject;
            _nuGetAndEnvDTEProjectCache.TryGetDTEProject(nuGetProjectSafeName, out dteProject);
            return dteProject;
        }

        public IEnumerable<NuGetProject> GetNuGetProjects()
        {
            Init();

            return _nuGetAndEnvDTEProjectCache.GetNuGetProjects();
        }

        private IEnumerable<EnvDTEProject> GetEnvDTEProjects()
        {
            Debug.Assert(ThreadHelper.CheckAccess());

            Init();

            return _nuGetAndEnvDTEProjectCache.GetEnvDTEProjects();
        }

        public bool IsSolutionOpen
        {
            get
            {
                return ThreadHelper.JoinableTaskFactory.Run(async delegate
                    {
                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                        return _dte != null &&
                               _dte.Solution != null &&
                               _dte.Solution.IsOpen &&
                               !IsSolutionSavedAsRequired();
                    });
            }
        }

        public string SolutionDirectory
        {
            get
            {
                if (!IsSolutionOpen)
                {
                    return null;
                }

                return ThreadHelper.JoinableTaskFactory.Run(async delegate
                    {
                        string solutionFilePath = await GetSolutionFilePathAsync();

                        if (String.IsNullOrEmpty(solutionFilePath))
                        {
                            return null;
                        }
                        return Path.GetDirectoryName(solutionFilePath);
                    });
            }
        }

        private async Task<string> GetSolutionFilePathAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            // Use .Properties.Item("Path") instead of .FullName because .FullName might not be
            // available if the solution is just being created
            string solutionFilePath = null;

            Property property = _dte.Solution.Properties.Item("Path");
            if (property == null)
            {
                return null;
            }
            try
            {
                // When using a temporary solution, (such as by saying File -> New File), querying this value throws.
                // Since we wouldn't be able to do manage any packages at this point, we return null. Consumers of this property typically 
                // use a String.IsNullOrEmpty check either way, so it's alright.
                solutionFilePath = (string)property.Value;
            }
            catch (COMException)
            {
                return null;
            }

            return solutionFilePath;
        }

        /// <summary>
        /// Checks whether the current solution is saved to disk, as opposed to be in memory.
        /// </summary>
        private bool IsSolutionSavedAsRequired()
        {
            Debug.Assert(ThreadHelper.CheckAccess());

            // Check if user is doing File - New File without saving the solution.
            object value;
            _vsSolution.GetProperty((int)(__VSPROPID.VSPROPID_IsSolutionSaveAsRequired), out value);
            if ((bool)value)
            {
                return true;
            }

            // Check if user unchecks the "Tools - Options - Project & Soltuions - Save new projects when created" option
            _vsSolution.GetProperty((int)(__VSPROPID2.VSPROPID_DeferredSaveSolution), out value);
            return (bool)value;
        }

        private void OnSolutionOpened()
        {
            Debug.Assert(ThreadHelper.CheckAccess());

            if (SolutionOpening != null)
            {
                SolutionOpening(this, EventArgs.Empty);
            }

            // although the SolutionOpened event fires, the solution may be only in memory (e.g. when
            // doing File - New File). In that case, we don't want to act on the event.
            if (!IsSolutionOpen)
            {
                return;
            }

            EnsureNuGetAndEnvDTEProjectCache();

            if (SolutionOpened != null)
            {
                SolutionOpened(this, EventArgs.Empty);
            }
        }

        private void OnAfterClosing()
        {
            if (SolutionClosed != null)
            {
                SolutionClosed(this, EventArgs.Empty);
            }
        }

        private void OnBeforeClosing()
        {
            DefaultNuGetProjectName = null;
            _nuGetAndEnvDTEProjectCache.Clear();

            if (SolutionClosing != null)
            {
                SolutionClosing(this, EventArgs.Empty);
            }
        }

        private void OnEnvDTEProjectRenamed(EnvDTEProject envDTEProject, string oldName)
        {
            // This is a solution event. Should be on the UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!String.IsNullOrEmpty(oldName) && IsSolutionOpen)
            {
                EnsureNuGetAndEnvDTEProjectCache();

                if (EnvDTEProjectUtility.IsSupported(envDTEProject))
                {
                    RemoveEnvDTEProjectFromCache(oldName);
                    AddEnvDTEProjectToCache(envDTEProject);
                    NuGetProject nuGetProject;
                    _nuGetAndEnvDTEProjectCache.TryGetNuGetProject(envDTEProject.Name, out nuGetProject);

                    if (NuGetProjectRenamed != null)
                    {
                        NuGetProjectRenamed(this, new NuGetProjectEventArgs(nuGetProject));
                    }
                }
                else if (EnvDTEProjectUtility.IsSolutionFolder(envDTEProject))
                {
                    // In the case where a solution directory was changed, project FullNames are unchanged. 
                    // We only need to invalidate the projects under the current tree so as to sync the CustomUniqueNames.
                    foreach (var item in EnvDTEProjectUtility.GetSupportedChildProjects(envDTEProject))
                    {
                        RemoveEnvDTEProjectFromCache(item.FullName);
                        AddEnvDTEProjectToCache(item);
                    }
                }
            }
        }

        private void OnEnvDTEProjectRemoved(EnvDTEProject envDTEProject)
        {
            // This is a solution event. Should be on the UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            RemoveEnvDTEProjectFromCache(envDTEProject.FullName);
            NuGetProject nuGetProject;
            _nuGetAndEnvDTEProjectCache.TryGetNuGetProject(envDTEProject.Name, out nuGetProject);

            if (NuGetProjectRemoved != null)
            {
                NuGetProjectRemoved(this, new NuGetProjectEventArgs(nuGetProject));
            }
        }

        private void OnEnvDTEProjectAdded(EnvDTEProject envDTEProject)
        {
            // This is a solution event. Should be on the UI thread
            ThreadHelper.ThrowIfNotOnUIThread();

            if (IsSolutionOpen
                && EnvDTEProjectUtility.IsSupported(envDTEProject)
                && !EnvDTEProjectUtility.IsParentProjectExplicitlyUnsupported(envDTEProject))
            {
                EnsureNuGetAndEnvDTEProjectCache();
                AddEnvDTEProjectToCache(envDTEProject);
                NuGetProject nuGetProject;
                _nuGetAndEnvDTEProjectCache.TryGetNuGetProject(envDTEProject.Name, out nuGetProject);

                if (NuGetProjectAdded != null)
                {
                    NuGetProjectAdded(this, new NuGetProjectEventArgs(nuGetProject));
                }
            }
        }

        private void SetDefaultProjectName()
        {
            // when a new solution opens, we set its startup project as the default project in NuGet Console
            var solutionBuild = (SolutionBuild2)_dte.Solution.SolutionBuild;
            if (solutionBuild.StartupProjects != null)
            {
                IEnumerable<object> startupProjects = (IEnumerable<object>)solutionBuild.StartupProjects;
                string startupProjectName = startupProjects.Cast<string>().FirstOrDefault();
                if (!String.IsNullOrEmpty(startupProjectName))
                {
                    EnvDTEProjectName envDTEProjectName;
                    if (_nuGetAndEnvDTEProjectCache.TryGetNuGetProjectName(startupProjectName, out envDTEProjectName))
                    {
                        DefaultNuGetProjectName = _nuGetAndEnvDTEProjectCache.IsAmbiguous(envDTEProjectName.ShortName) ?
                            envDTEProjectName.CustomUniqueName :
                            envDTEProjectName.ShortName;
                    }
                }
            }
        }

        private void EnsureNuGetAndEnvDTEProjectCache()
        {
            Debug.Assert(ThreadHelper.CheckAccess());

            if (!_nuGetAndEnvDTEProjectCache.IsInitialized && IsSolutionOpen)
            {
                var factory = GetProjectFactory();

                try
                {
                    var supportedProjects = EnvDTESolutionUtility.GetAllEnvDTEProjects(_dte)
                        .Where(project => EnvDTEProjectUtility.IsSupported(project));

                    _nuGetAndEnvDTEProjectCache.Initialize(supportedProjects, factory);
                    SetDefaultProjectName();
                }
                catch
                {
                    _nuGetAndEnvDTEProjectCache.Clear();
                    DefaultNuGetProjectName = null;

                    throw;
                }
            }
        }

        private void AddEnvDTEProjectToCache(EnvDTEProject envDTEProject)
        {
            Debug.Assert(ThreadHelper.CheckAccess());

            if (!EnvDTEProjectUtility.IsSupported(envDTEProject))
            {
                return;
            }

            EnvDTEProjectName oldEnvDTEProjectName;
            _nuGetAndEnvDTEProjectCache.TryGetProjectNameByShortName(EnvDTEProjectUtility.GetName(envDTEProject), out oldEnvDTEProjectName);

            EnvDTEProjectName newEnvDTEProjectName = _nuGetAndEnvDTEProjectCache.AddProject(envDTEProject, GetProjectFactory());

            if (String.IsNullOrEmpty(DefaultNuGetProjectName)
                ||
                newEnvDTEProjectName.ShortName.Equals(DefaultNuGetProjectName, StringComparison.OrdinalIgnoreCase))
            {
                DefaultNuGetProjectName = oldEnvDTEProjectName != null ?
                    oldEnvDTEProjectName.CustomUniqueName :
                    newEnvDTEProjectName.ShortName;
            }
        }

        private void RemoveEnvDTEProjectFromCache(string name)
        {
            // Do nothing if the cache hasn't been set up
            if (_nuGetAndEnvDTEProjectCache == null)
            {
                return;
            }

            EnvDTEProjectName envDTEProjectName;
            _nuGetAndEnvDTEProjectCache.TryGetNuGetProjectName(name, out envDTEProjectName);

            // Remove the project from the cache
            _nuGetAndEnvDTEProjectCache.RemoveProject(name);

            if (!_nuGetAndEnvDTEProjectCache.Contains(DefaultNuGetProjectName))
            {
                DefaultNuGetProjectName = null;
            }

            // for LightSwitch project, the main project is not added to _projectCache, but it is called on removal. 
            // in that case, projectName is null.
            if (envDTEProjectName != null
                && envDTEProjectName.CustomUniqueName.Equals(DefaultNuGetProjectName, StringComparison.OrdinalIgnoreCase)
                && !_nuGetAndEnvDTEProjectCache.IsAmbiguous(envDTEProjectName.ShortName))
            {
                DefaultNuGetProjectName = envDTEProjectName.ShortName;
            }
        }

        private void Init()
        {
            try
            {
                // If already initialized, need not be on the UI thread
                if (!_initialized)
                {
                    _initialized = true;

                    ThreadHelper.JoinableTaskFactory.Run(async delegate
                        {
                            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                            if (_dte.Solution.IsOpen)
                            {
                                OnSolutionOpened();
                            }
                        });
                }
            }
            catch (Exception ex)
            {
                // ignore errors
                Debug.Fail(ex.ToString());
                Trace.WriteLine(ex.ToString());
            }
        }

        private VSNuGetProjectFactory GetProjectFactory()
        {
            var settings = ServiceLocator.GetInstance<Configuration.ISettings>();

            // We are doing this to avoid a loop at initialization. We probably want to remove this dependency alltogether.
            var factory = new VSNuGetProjectFactory(() => PackagesFolderPathUtility.GetPackagesFolderPath(this, settings));

            return factory;
        }

        // REVIEW: This might be inefficient, see what we can do with caching projects until references change
        internal static IEnumerable<EnvDTEProject> GetDependentEnvDTEProjects(IDictionary<string, List<EnvDTEProject>> dependentEnvDTEProjectsDictionary, EnvDTEProject envDTEProject)
        {
            Debug.Assert(ThreadHelper.CheckAccess());

            if (envDTEProject == null)
            {
                throw new ArgumentNullException(nameof(envDTEProject));
            }

            List<Project> dependents;
            if (dependentEnvDTEProjectsDictionary.TryGetValue(EnvDTEProjectUtility.GetUniqueName(envDTEProject), out dependents))
            {
                return dependents;
            }

            return Enumerable.Empty<EnvDTEProject>();
        }

        internal async Task<IDictionary<string, List<EnvDTEProject>>> GetDependentEnvDTEProjectsDictionaryAsync()
        {
            // Get all of the projects in the solution and build the reverse graph. i.e.
            // if A has a project reference to B (A -> B) the this will return B -> A
            // We need to run this on the ui thread so that it doesn't freeze for websites. Since there might be a 
            // large number of references.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            Init();

            var dependentEnvDTEProjectsDictionary = new Dictionary<string, List<Project>>();
            var envDTEProjects = GetEnvDTEProjects();

            foreach (EnvDTEProject envDTEProj in envDTEProjects)
            {
                if (EnvDTEProjectUtility.SupportsReferences(envDTEProj))
                {
                    foreach (var referencedProject in EnvDTEProjectUtility.GetReferencedProjects(envDTEProj))
                    {
                        AddDependentProject(dependentEnvDTEProjectsDictionary, referencedProject, envDTEProj);
                    }
                }
            }

            return dependentEnvDTEProjectsDictionary;
        }

        private static void AddDependentProject(IDictionary<string, List<EnvDTEProject>> dependentEnvDTEProjectsDictionary,
            EnvDTEProject envDTEProject, EnvDTEProject dependentEnvDTEProject)
        {
            Debug.Assert(ThreadHelper.CheckAccess());

            string uniqueName = EnvDTEProjectUtility.GetUniqueName(envDTEProject);

            List<EnvDTEProject> dependentEnvDTEProjects;
            if (!dependentEnvDTEProjectsDictionary.TryGetValue(uniqueName, out dependentEnvDTEProjects))
            {
                dependentEnvDTEProjects = new List<EnvDTEProject>();
                dependentEnvDTEProjectsDictionary[uniqueName] = dependentEnvDTEProjects;
            }
            dependentEnvDTEProjects.Add(dependentEnvDTEProject);
        }

        #region IVsSelectionEvents implementation

        public int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive)
        {
            if (dwCmdUICookie == _solutionLoadedUICookie
                && fActive == 1)
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                OnSolutionOpened();
                // We must call DeleteMarkedPackageDirectories outside of OnSolutionOpened, because OnSolutionOpened might be called in the constructor
                // and DeleteOnRestartManager requires VsFileSystemProvider and RepositorySetings which both have dependencies on SolutionManager.
                // In practice, this code gets executed even when a solution is opened directly during Visual Studio startup.
                //DeleteOnRestartManager.Value.DeleteMarkedPackageDirectories();
            }

            return VSConstants.S_OK;
        }

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            return VSConstants.S_OK;
        }

        public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld, IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew)
        {
            return VSConstants.S_OK;
        }

        #endregion
    }
}
