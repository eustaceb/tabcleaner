using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace TabCleaner
{

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TabCleaner
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
        /// Initializes a new instance of the <see cref="TabCleaner"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TabCleaner(Package package)
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
        public static TabCleaner Instance
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
            Instance = new TabCleaner(package);
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
            // Get the list of open documents
            var dteService = ServiceProvider.GetService(typeof(DTE)) as DTE;
            var openDocs = dteService.Documents;

            // Rescan all projects to find paths that we'll consider local.
            RescanProjects();

            // If the document is not local to one of the projects in solution, close it.
            foreach (Document doc in openDocs)
            {
                if (!IsFileLocal(doc.Path))
                {
                    // Close only if the document has been saved or the user has selected to close modified.
                    var tabCleanerPkg = ServiceProvider.GetService(typeof(TabCleanerPackage)) as TabCleanerPackage;
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

            var dteService = ServiceProvider.GetService(typeof(DTE)) as DTE;

            // Go through each project
            foreach (Project pro in dteService.Solution.Projects)
            {
                // Go through each property of the project
                foreach (Property prop in pro.Properties)
                {
                    try
                    {
                        // And find the "ProjectDirectory" property.
                        if (prop.Name != null && prop.Name.ToString().Equals("ProjectDirectory"))
                        {
                            if (prop.Value != null)
                            {
                                // If it's set to something sensible
                                String strValue = prop.Value.ToString();
                                if (!strValue.Equals(""))
                                {
                                    // Add it to list of local paths
                                    m_localPaths.Add((new FileInfo(strValue).DirectoryName).ToString());
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine("Caught exception: " + e.Message);
                    }
                }
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
