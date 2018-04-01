using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TabCleaner
{

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CloseExternalCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("914e0122-bbe2-4687-bf51-8f4997603c42");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Paths fetched from the project files
        /// </summary>
        private static List<string> m_localPaths;

        /// <summary>
        /// Initializes a new instance of the <see cref="CloseExternalCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CloseExternalCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CloseExternalCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new CloseExternalCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Get the DTE service for obtaining open documents
            var dteService = ServiceProvider.GetService(typeof(DTE)) as DTE;

            // Get the package service
            var tabCleanerPkg = ServiceProvider.GetService(typeof(TabCleanerPackage)) as TabCleanerPackage;


            // Rescan all projects to find paths that we'll consider local.
            RescanProjects();

            // If the document is not local to one of the projects in solution, close it.
            foreach (Document doc in dteService.Documents)
            {
                if (doc.Path.Length > 0 && !IsFileLocal(doc.Path.ToLower()))
                {
                    // Close only if the document has been saved or the user has selected to close modified.
                    if (tabCleanerPkg.CloseModified || doc.Saved)
                        doc.Close();
                }
            }
        }

        /// <summary>
        /// Scans all solution projects and gets their base locations.
        /// </summary>
        private void RescanProjects()
        {
            // Reset the local paths list
            m_localPaths = new List<string>();

            // The following bit is taken from: https://stackoverflow.com/a/41052537
            // Find all project files
            IVsSolution sol = ServiceProvider.GetService(typeof(SVsSolution)) as IVsSolution;
            uint numProjects;
            ErrorHandler.ThrowOnFailure(sol.GetProjectFilesInSolution((uint)__VSGETPROJFILESFLAGS.GPFF_SKIPUNLOADEDPROJECTS, 0, null, out numProjects));
            string[] projects = new string[numProjects];
            ErrorHandler.ThrowOnFailure(sol.GetProjectFilesInSolution((uint)__VSGETPROJFILESFLAGS.GPFF_SKIPUNLOADEDPROJECTS, numProjects, projects, out numProjects));
            // End of public wisdom

            // Get just the base dirs
            for (int i = 0; i < projects.Length; i++)
            {
                m_localPaths.Add(new FileInfo(projects[i]).Directory.ToString().ToLower());
            }
        }

        /// <summary>
        /// Checks whether a file is located in one of the solution projects base folders.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private bool IsFileLocal(string filePath)
        {
            // Go through each local path.
            foreach (string projectPath in m_localPaths)
            {
                // And check if it is somewhere in the base of the filepath that's being checked.
                if (filePath.Contains(projectPath))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
