using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
        private IVsRunningDocumentTable _rdt;
        private DTE2 _dte;

        private string _lastProcessed;

        public void Subscribe(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _dte = serviceProvider.GetService(typeof(DTE)) as DTE2;
            _rdt = serviceProvider.GetService(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            _rdt?.AdviseRunningDocTableEvents(this, out _);
        }

        /*looks for another document with the same file extension, if found close this document -> activate the second documents window -> open this document*/
        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (fFirstShow == 0)
                return VSConstants.S_OK;

            pFrame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out object filePath);
            string currFile = filePath?.ToString();

            string currentExt = Path.GetExtension(currFile);
            if (currFile == null || currFile.Length == 0 || currentExt == null || currentExt.Length == 0)
                return VSConstants.S_OK;

            /*cache last processed to avoid infinite for loops of closing and opening*/
            if (currFile == _lastProcessed)
            {
                _lastProcessed = null;
                return VSConstants.S_OK;
            }
            _lastProcessed = currFile;

            /*no action needed if the active document alredy has the correct file extension*/
            string activeFile = _dte.ActiveWindow?.Document?.FullName;
            string activeExt = Path.GetExtension(activeFile);
            if (activeExt == currentExt)
                return VSConstants.S_OK;

            Window bestWindow = null;
            foreach (Window window in _dte.Windows)
            {
                if (window == null || window.Document == null || window.Visible == false)
                    continue;

                string docPathName = window.Document.FullName;

                /*if a window is already open for this doc then we don't want to sort it*/
                if (docPathName == currFile)
                    return VSConstants.S_OK;


                if (Path.GetExtension(docPathName) != currentExt)
                    continue;

                //TODO: prioritize active tabs
                if (bestWindow == null)
                {
                    bestWindow = window;
                }
            }

            /*close and repoen in correct document group*/
            if (bestWindow != null)
            {
                ReopenDocumentAsync(currFile, bestWindow);
            }

            return VSConstants.S_OK;
        }

        private async Task ReopenDocumentAsync(string filePathToReopen, Window windowToOpenFrom)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (filePathToReopen == null || windowToOpenFrom == null)
                return;

            /*wait so the text selection has time to load in*/
            await Task.Delay(100);

            Document doc = _dte.Documents.Cast<Document>().FirstOrDefault(d =>
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return d.FullName == filePathToReopen;
            });

            if (doc == null)
                return;

            int line = 0;
            if (doc.Selection is TextSelection startSelection)
            {
                line = startSelection.CurrentLine;
            }

            doc.Close(vsSaveChanges.vsSaveChangesNo);

            windowToOpenFrom.Activate();
            Window newWindow = _dte.ItemOperations.OpenFile(filePathToReopen, EnvDTE.Constants.vsViewKindCode);

            /*removed windowToOpenFrom from navigation history*/
            _dte.ExecuteCommand("View.NavigateBackward");
            _dte.ExecuteCommand("View.NavigateBackward");
            newWindow.Activate();
            if (newWindow.Selection is TextSelection endSelection)
            {
                endSelection?.GotoLine(line);
            }
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame) => VSConstants.S_OK;
        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => VSConstants.S_OK;
        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => VSConstants.S_OK;
        public int OnAfterSave(uint docCookie) => VSConstants.S_OK;
        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs) => VSConstants.S_OK;
    }
}
