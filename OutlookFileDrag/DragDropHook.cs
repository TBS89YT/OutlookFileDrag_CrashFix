using System;
using EasyHook;
using log4net;

namespace OutlookFileDrag
{
    class DragDropHook : IDisposable
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private LocalHook hook;
        private NativeMethods.DragDropDelegate hookDelegate;
        private bool disposed = false;
        private bool isHooked = false;

        public DragDropHook()
        {
            try
            {
                //Hook OLE drag and drop event
                log.Info("Creating hook for DoDragDrop method of ole32.dll");
                this.hookDelegate = new NativeMethods.DragDropDelegate(DragDropHook.DoDragDropHook);
                hook = EasyHook.LocalHook.Create(EasyHook.LocalHook.GetProcAddress("ole32.dll", "DoDragDrop"),
                    this.hookDelegate, null);
            }
            catch (Exception ex)
            {
                log.Error("Error creating hook", ex);
                throw;
            }
        }

        public bool IsHooked
        {
            get
            {
                return isHooked;
            }
        }

        public void Start()
        {
            try
            {
                if (isHooked)
                    return;

                log.Info("Starting hook");
                //Hook current (UI) thread
                hook.ThreadACL.SetInclusiveACL(new Int32[] { 0 });
                isHooked = true;
            }
            catch (Exception ex)
            {
                log.Error("Error starting hook", ex);
                throw;
            }
        }

        public void Stop()
        {
            try
            {
                if (!isHooked)
                    return;

                log.Info("Stopping hook");
                //Stop hooking all threads
                hook.ThreadACL.SetInclusiveACL(new Int32[] { });
                isHooked = false;
                log.Info("Stopped hook");
            }
            catch (Exception ex)
            {
                log.Error("Error stopping hook", ex);
                throw;
            }
        }

        public static int DoDragDropHook(IntPtr pDataObjPtr, IntPtr pDropSource, uint dwOKEffects, IntPtr pdwEffectPtr)
        {
            // OUTER try/catch: If ANYTHING goes wrong (including marshaling the COM pointer),
            // fall back to calling the original DoDragDrop with raw pointers.
            try
            {
                log.Info("Drag started");

                uint pdwEffect = 0;
                if (pdwEffectPtr != IntPtr.Zero)
                {
                    pdwEffect = (uint)System.Runtime.InteropServices.Marshal.ReadInt32(pdwEffectPtr);
                }

                // Manually marshal the raw IntPtr to our managed COM interface.
                // This is the step that used to crash when CLR did it automatically.
                NativeMethods.IDataObject pDataObj = null;
                try
                {
                    pDataObj = (NativeMethods.IDataObject)System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(pDataObjPtr);
                }
                catch (Exception marshalEx)
                {
                    // COM object is incompatible or protected -- pass through to original API
                    log.Warn("Could not marshal IDataObject from IntPtr -- passing through to original DoDragDrop", marshalEx);
                    int failRes = NativeMethods.DoDragDropRaw(pDataObjPtr, pDropSource, dwOKEffects, out pdwEffect);
                    if (pdwEffectPtr != IntPtr.Zero) System.Runtime.InteropServices.Marshal.WriteInt32(pdwEffectPtr, (int)pdwEffect);
                    return failRes;
                }

                // Check if the data object contains virtual files (Outlook attachments/emails)
                if (!DataObjectHelper.GetDataPresent(pDataObj, "FileGroupDescriptorW") && !DataObjectHelper.GetDataPresent(pDataObj, "FileGroupDescriptor"))
                {
                    log.Info("No virtual files found -- continuing original drag");
                    // Use raw call to avoid double-marshaling issues
                    int rawRes = NativeMethods.DoDragDropRaw(pDataObjPtr, pDropSource, dwOKEffects, out pdwEffect);
                    if (pdwEffectPtr != IntPtr.Zero) System.Runtime.InteropServices.Marshal.WriteInt32(pdwEffectPtr, (int)pdwEffect);
                    return rawRes;
                }

                //Start new drag
                log.Info("Virtual files found -- starting new drag adding CF_HDROP format");
                log.InfoFormat("Files: {0}", string.Join(",", DataObjectHelper.GetFilenames(pDataObj)));

                OutlookDataObject newDataObj = new OutlookDataObject(pDataObj);

                // Pin the managed wrapper so the GC doesn't relocate it during the native call.
                // Get a COM-callable IntPtr for our managed IDataObject wrapper.
                IntPtr pNewDataObj = System.Runtime.InteropServices.Marshal.GetComInterfaceForObject(
                    newDataObj, typeof(NativeMethods.IDataObject));
                try
                {
                    int result = NativeMethods.DoDragDropRaw(pNewDataObj, pDropSource, dwOKEffects, out pdwEffect);

                    //If files were dropped and drop effect was "move", then override to "copy" so original item is not deleted
                    if (newDataObj.FilesDropped && pdwEffect == NativeMethods.DROPEFFECT_MOVE)
                        pdwEffect = NativeMethods.DROPEFFECT_COPY;

                    //Get result
                    log.InfoFormat("DoDragDrop effect: {0} result: {1}", pdwEffect, result);
                    if (pdwEffectPtr != IntPtr.Zero) System.Runtime.InteropServices.Marshal.WriteInt32(pdwEffectPtr, (int)pdwEffect);
                    return result;
                }
                finally
                {
                    // Release the COM reference we obtained via GetComInterfaceForObject
                    System.Runtime.InteropServices.Marshal.Release(pNewDataObj);
                }
            }
            catch (Exception ex)
            {
                log.Warn("Dragging error -- attempting raw fallback", ex);

                // Last resort: try to call the original DoDragDrop with raw pointers
                uint effectFallback = 0;
                int fallbackResult = NativeMethods.DoDragDropRaw(pDataObjPtr, pDropSource, dwOKEffects, out effectFallback);
                if (pdwEffectPtr != IntPtr.Zero) System.Runtime.InteropServices.Marshal.WriteInt32(pdwEffectPtr, (int)effectFallback);
                return fallbackResult;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                hook.Dispose();
            }

            disposed = true;
        }
    }

}
