using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TabCleaner
{

    /// <summary>
    /// Option page for the extension.
    /// </summary>
    public class OptionPageGrid : DialogPage
    {
        private bool closeModified = false;
        private bool closeLRU = false;
        private uint noOfLRUDocs = 2;
        
        [Category("Common settings")]
        [DisplayName("Attempt to close modified documents")]
        [Description("If this is set to true, modified documents will be ignored. Otherwise, you will be prompted to save when closing.")]
        public bool CloseModified
        {
            get { return closeModified; }
            set { closeModified = value; }
        }

        //[Category("Least Recently Used options")]
        //[DisplayName("Enable LRU closing")]
        //[Description("If this is set to true, an extra option to close LRU documents will be added to the menu. REQUIRES RESTART")]
        //public bool CloseLRU
        //{
        //    get { return closeLRU; }
        //    set { closeLRU = value; }
        //}

        //[Category("Least Recently Used options")]
        //[DisplayName("Number of documents to close")]
        //[Description("Determines how many least recently used documents will be closed.")]
        //public uint NoOfLRUDocuments
        //{
        //    get { return noOfLRUDocs; }
        //    set { noOfLRUDocs = value; }
        //}

    }

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    /// 
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(TabCleanerPackage.PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [ProvideOptionPage(typeof(OptionPageGrid), "TabCleaner", "Basic settings", 0, 0, true)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class TabCleanerPackage : Package
    {

        /// <summary>
        /// CloseExternalPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "c4a00f92-744c-40b4-a889-ef3baafd10b0";

        /// <summary>
        /// Initializes a new instance of the <see cref="CloseExternalCommand"/> class.
        /// </summary>
        public TabCleanerPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <summary>
        /// Returns a bool suggesting whether to close modified documents.
        /// </summary>
        public bool CloseModified
        {
            get
            {
                OptionPageGrid page = (OptionPageGrid)GetDialogPage(typeof(OptionPageGrid));
                return page.CloseModified;
            }
        }

        ///// <summary>
        ///// Returns a bool suggesting whether to provide the ability to close LRU documents.
        ///// </summary>
        //public bool CloseLRU
        //{
        //    get
        //    {
        //        OptionPageGrid page = (OptionPageGrid)GetDialogPage(typeof(OptionPageGrid));
        //        return page.CloseLRU;
        //    }
        //}

        ///// <summary>
        ///// Returns the number of least recently used documents to be closed.
        ///// </summary>
        //public uint NoOfLRUDocuments
        //{
        //    get
        //    {
        //        OptionPageGrid page = (OptionPageGrid)GetDialogPage(typeof(OptionPageGrid));
        //        return page.NoOfLRUDocuments;
        //    }
        //}

        #region Package Members

        DocumentListener m_docListener;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();
            CloseExternalCommand.Initialize(this);

            //if (CloseLRU)
            //{
            //    CloseLRUCommand.Initialize(this);
            //    var rdt = (IVsRunningDocumentTable)Package.GetGlobalService(typeof(SVsRunningDocumentTable));
            //    m_docListener = new DocumentListener(rdt);

            //}
        }

        public List<string> GetLRUDocs()
        {
            //if (m_docListener != null)
            //    return m_docListener.GetLRUDocs(NoOfLRUDocuments);
            return new List<string>();
        }
        #endregion
    }
}
