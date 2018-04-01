using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TabCleaner
{
    public class DocumentListener : IVsRunningDocTableEvents, IDisposable
    {
        uint m_rdtCookie;
        IVsRunningDocumentTable m_rdt;

        private static uint m_currentStamp = 0;
        private static Dictionary<uint, uint> m_docStamps = new Dictionary<uint, uint>();


        public DocumentListener(IVsRunningDocumentTable rdt)
        {
            m_rdt = rdt;
            m_rdt.AdviseRunningDocTableEvents(this, out m_rdtCookie);
            System.Diagnostics.Debug.WriteLine("Registered a document listener");
        }

        public List<string> GetLRUDocs(uint nCount)
        {
            // Store resulting document paths in a list.
            List<string> result = new List<string>();

            // Sort documents by timestamps by putting them in a SortedList.
            SortedList<uint, uint> sortedByTimestamps = new SortedList<uint, uint>();
            foreach (var kvp in m_docStamps.AsEnumerable())
            {
                // Store the timestamp in the key, and the docId in the value.
                sortedByTimestamps.Add(kvp.Value, kvp.Key);
            }

            // Variable to capture the document path in.
            string docPath;

            // These variables are used for scrapping unnecessary data.
            uint nil;
            IVsHierarchy nilHierarchy;
            IntPtr nilData;

            // Document counter for knowing when to stop when taking nCount LRU documents.
            int nDocuments = 0;

            // Go through each key-value pair in the timestamp list.
            foreach (var kvp in sortedByTimestamps.ToList())
            {
                // If the document counter is at the requested count, stop.
                if (nDocuments >= nCount)
                    break;

                // Get document information from the running document table.
                m_rdt.GetDocumentInfo(kvp.Value, out nil, out nil, out nil, out docPath, out nilHierarchy, out nil, out nilData);

                // @TODO investigate why sometimes the path ends up being null.
                if (docPath != null)
                {
                    // Add to result list, increment document counter and remove document from timestamp table.
                    result.Add(docPath);
                    nDocuments++;
                    m_docStamps.Remove(kvp.Value);
                }
            }

            return result;
        }

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            m_docStamps[docCookie] = m_currentStamp++;
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        #region IDisposable Members
        void IDisposable.Dispose()
        {
            try
            {
                if (m_rdtCookie != 0) m_rdt.UnadviseRunningDocTableEvents(m_rdtCookie);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message.ToString());
            }
        }
        #endregion
    }
}
