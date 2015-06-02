// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;

namespace NuGet.PackageManagement.VisualStudio
{
    /// <summary>
    /// Represents a text file inside a Visual Studio solution
    /// </summary>
    [DebuggerDisplay("{Moniker}")]
    internal class BufferedTextFile : IDisposable
    {
        private string _fullPath;
        private RunningDocumentTable _runningDocumentTable;
        private IVsEditorAdaptersFactoryService _editorAdapterFactory;
        private IVsInvisibleEditorManager _invisibleEditorManager;
        private ITextBuffer _textBuffer;
        private uint _docCookie;
        private bool _isDisposed = false; // To detect redundant calls

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="path">Full path to the file</param>
        /// <param name="runningDocumentTable">The running document table</param>
        /// <param name="editorAdapterFactory">The editor adapter factory</param>
        /// <param name="invisibleEditorManager">The invisible editor manager</param>
        private BufferedTextFile(
            string path,
            RunningDocumentTable runningDocumentTable,
            IVsEditorAdaptersFactoryService editorAdapterFactory,
            IVsInvisibleEditorManager invisibleEditorManager)
        {
            _fullPath = BufferedTextFile.ValidateNotNull(path, "path");
            _runningDocumentTable = BufferedTextFile.ValidateNotNull(runningDocumentTable, "runningDocumentTable");
            _editorAdapterFactory = BufferedTextFile.ValidateNotNull(editorAdapterFactory, "editorAdapterFactory");
            _invisibleEditorManager = BufferedTextFile.ValidateNotNull(invisibleEditorManager, "invisibleEditorManager");

            InjectableReadAllText = File.ReadAllText;
        }

        public static BufferedTextFile CreateBufferedTextFile(IServiceProvider serviceProvider, string path)
        {
            var componentModelHost = serviceProvider.GetService(typeof(SComponentModel)) as IComponentModel;
            var invisibleEditorManager = serviceProvider.GetService(typeof(IVsInvisibleEditorManager)) as IVsInvisibleEditorManager;
            var editorAdapterFactory = componentModelHost.GetService<IVsEditorAdaptersFactoryService>();
            var runningDocumentTable = new RunningDocumentTable(serviceProvider);

            return new BufferedTextFile(path, runningDocumentTable, editorAdapterFactory, invisibleEditorManager);
        }

        /// <summary>
        /// Injectable function for reading all text from a file on disk
        /// </summary>
        private Func<string, string> InjectableReadAllText { get; set; }

        /// <summary>
        /// Returns a unique moniker representing this file
        /// </summary>
        public string Moniker
        {
            get
            {
                return _fullPath;
            }
        }

        /// <summary>
        /// Returns the name of the file
        /// </summary>
        public string FileName
        {
            get
            {
                return Path.GetFileName(_fullPath);
            }
        }

        /// <summary>
        /// Retrieves the current text contents, whether currently being edited or not
        /// </summary>
        /// <returns>Current contents as a string</returns>
        /// <remarks>Can throw I/O exceptions</remarks>
        public string GetCurrentContents()
        {
            // If the file is open in an editor, return the editor's current contents, saved or
            // not. Otherwise, read the file straight from disk.
            var contents = _runningDocumentTable.GetRunningDocumentContents(_fullPath);
            return contents ?? InjectableReadAllText(_fullPath);
        }

        /// <summary>
        /// Gets an ITextBuffer that represents the given file
        /// </summary>
        /// <param name="ensureWritable">If true, make the buffer editable immediately</param>
        /// <returns>An ITextBuffer instance</returns>
        public ITextBuffer GetTextBuffer(bool ensureWritable)
        {
            if (_textBuffer != null)
            {
                return _textBuffer;
            }

            bool wrapException = true;
            try
            {
                // Retrieve an invisible editor
                IVsInvisibleEditor editor;
                ErrorHandler.ThrowOnFailure(
                    _invisibleEditorManager.RegisterInvisibleEditor(
                        _fullPath,
                        null, // Any project
                        0,    // Don't try to keep the docData in the RDT longer than it needs to be
                        null, // Use default doc factory
                        out editor));

                var pDocData = IntPtr.Zero;
                try
                {
                    // Get the DocData for the editor
                    object docData = null;
                    var riid = typeof(IVsPersistDocData).GUID;
                    ErrorHandler.ThrowOnFailure(
                        editor.GetDocData(
                            ensureWritable ? 1 : 0,
                            ref riid,
                            out pDocData));
                    docData = Marshal.GetObjectForIUnknown(pDocData);

                    // Retrieve IVsTextBuffer, from which we can create an ITextBuffer
                    var vsTextBuffer = docData as IVsTextBuffer;
                    if (vsTextBuffer == null)
                    {
                        // We have a docData, but it's not IVsTextBuffer, so we can play in the same sandbox
                        wrapException = false;
                        throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture,
                            Strings.Error_OpenInIncompatibleBuffer_1arg, _fullPath));
                    }

                    // Create the ITextBuffer instance
                    var buffer = _editorAdapterFactory.GetDataBuffer(vsTextBuffer);
                    if (buffer == null)
                    {
                        Debug.Fail("GetDataBuffer failed");
                        wrapException = false;
                        throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture,
                            Strings.Error_OpenInIncompatibleBuffer_1arg, _fullPath));
                    }

                    // Lock the document for editing
                    var info = _runningDocumentTable.GetDocumentInfo(Moniker);
                    var cookie = info.DocCookie;

                    _runningDocumentTable.LockDocument(_VSRDTFLAGS.RDT_EditLock, cookie);
                    _textBuffer = buffer;
                    _docCookie = cookie;
                    return _textBuffer;
                }
                finally
                {
                    if (pDocData != IntPtr.Zero)
                    {
                        Marshal.Release(pDocData);
                    }
                }
            }
            catch (Exception ex)
            {
                if (wrapException)
                {
                    throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture,
                        Strings.Error_GetTextBuffer_2args, _fullPath, ex.Message), ex);
                }
                else
                {
                    throw;
                }
            }
        }

        public void SaveIfDirty()
        {
            _runningDocumentTable.SaveFileIfDirty(_fullPath);
        }

        public bool Equals(BufferedTextFile other)
        {
            if (other == null)
            {
                return false;
            }

            return string.Equals(other.Moniker, Moniker, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as BufferedTextFile);
        }

        public override int GetHashCode()
        {
            return _fullPath.GetHashCode();
        }

        #region IDisposable Support

        protected virtual void DisposeManagedResources()
        {
            if (_textBuffer != null)
            {
                _textBuffer = null;

                var editLocks = _runningDocumentTable.GetDocumentInfo(Moniker).EditLocks;
                if (editLocks == 1)
                {
                    // If we're the last edit lock, then close/save the file
                    var closeResult = _runningDocumentTable.CloseDocument(__FRAMECLOSE.FRAMECLOSE_SaveIfDirty, _docCookie);
                    Debug.Assert(closeResult == CloseResult.Closed);
                }

                // Release edit lock
                var unlockResult = _runningDocumentTable.UnlockDocument(_VSRDTFLAGS.RDT_EditLock, _docCookie);
                Debug.Assert(unlockResult == UnlockResult.Unlocked);
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_isDisposed)
            {
                try
                {
                    if (disposing)
                    {
                        DisposeManagedResources();
                    }
                }
                finally
                {
                    _isDisposed = true;
                }
            }
        }

        public void Dispose()
        {
            // Do not change code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
        }

        #endregion

        private static T ValidateNotNull<T>(T value, string name)
        {
            if (value == null)
            {
                throw new ArgumentNullException(name);
            }

            return value;
        }
    }
}
