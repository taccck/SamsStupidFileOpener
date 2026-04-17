using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace SamsStupidFileOpener
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class SamsStupidFileOpenerPackage : AsyncPackage
    {
        public const string PackageGuidString = "3e353f4d-02ef-4545-8b6a-bcdf4eb1fc71";

        private DocumentEventListener _docListener;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _docListener = new DocumentEventListener();
            _docListener.Subscribe(this);
        }
    }

    public class DocumentEventListener : IVsRunningDocTableEvents
    {
        private uint _rdtEventsCookie;
        private IVsRunningDocumentTable _rdt;
        private DTE2 _dte;

        private string _lastProcessed;

        public void Subscribe(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _dte = serviceProvider.GetService(typeof(DTE)) as DTE2;
            _rdt = serviceProvider.GetService(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            _rdt?.AdviseRunningDocTableEvents(this, out _rdtEventsCookie);
        }

        /*looks for another document with the same file extension, if found close this document -> activate the second documents window -> open this document*/
        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (fFirstShow == 0)
                return VSConstants.S_OK;

            pFrame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out object filePath);
            string currFile = filePath?.ToString();

            /*cache last processed to avoid infinite for loops of closing and opening*/
            if (currFile == null || currFile.Length == 0 || currFile == _lastProcessed)
                return VSConstants.S_OK;
            _lastProcessed = currFile;

            string currentExt = Path.GetExtension(currFile);

            foreach (Window window in _dte.Windows)
            {
                if (window == null || window.Document == null || window.Visible == false)
                    continue;

                string docPathName = window.Document.FullName;

                if (docPathName == currFile || Path.GetExtension(docPathName) != currentExt)
                    continue;

                Document doc = _dte.Documents.Cast<Document>().FirstOrDefault(d =>
                {
                    ThreadHelper.ThrowIfNotOnUIThread();
                    return d.FullName == currFile;
                });
                doc?.Close(vsSaveChanges.vsSaveChangesNo);
                window.Activate();
                _dte.ItemOperations.OpenFile(currFile, EnvDTE.Constants.vsViewKindCode);
                break;
            }

            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame) => VSConstants.S_OK;
        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => VSConstants.S_OK;
        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => VSConstants.S_OK;
        public int OnAfterSave(uint docCookie) => VSConstants.S_OK;
        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs) => VSConstants.S_OK;
    }
}
