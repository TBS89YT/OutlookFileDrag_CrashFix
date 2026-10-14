using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using log4net;

namespace OutlookFileDrag
{
    static class DataObjectHelper
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        internal static int GetClipboardFormat(string name)
        {
            var format = System.Windows.Forms.DataFormats.GetFormat(name);
            if (format == null)
                return 0;
            else
                return format.Id;
        }

        internal static bool GetDataPresent(NativeMethods.IDataObject data, string formatName)
        {
            //Check if data is present
            log.DebugFormat("Get data present: {0}", formatName);
            FORMATETC format = new FORMATETC();
            format.cfFormat = (short)GetClipboardFormat(formatName);
            format.dwAspect = DVASPECT.DVASPECT_CONTENT;
            format.lindex = -1;
            format.ptd = IntPtr.Zero;
            format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

            return data.QueryGetData(ref format) == NativeMethods.S_OK;
        }

        internal static void SetDropFiles(ref STGMEDIUM medium, string[] filenames)
        {
            //Create DROPFILES structure
            NativeMethods.DROPFILES dropFiles = new NativeMethods.DROPFILES();
            int dropFilesSize = Marshal.SizeOf(typeof(NativeMethods.DROPFILES));
            dropFiles.pFiles = dropFilesSize;
            dropFiles.fWide = true;     //Unicode

            //Get null-separated list of filenames terminated with double null
            string filenameList = string.Join("\0", filenames) + "\0\0";
            byte[] filenameBytes = System.Text.Encoding.Unicode.GetBytes(filenameList);

            // OLE requires a movable global memory handle for TYMED_HGLOBAL.
            int dataLength = dropFilesSize + filenameBytes.Length;
            IntPtr hGlobal = NativeMethods.GlobalAlloc(
                NativeMethods.GMEM_MOVEABLE, new UIntPtr((uint)dataLength));
            if (hGlobal == IntPtr.Zero)
                throw new OutOfMemoryException("GlobalAlloc failed for CF_HDROP");

            IntPtr ptrDropFiles = IntPtr.Zero;
            bool copied = false;
            try
            {
                ptrDropFiles = NativeMethods.GlobalLock(hGlobal);
                if (ptrDropFiles == IntPtr.Zero)
                    throw new InvalidOperationException("GlobalLock failed for CF_HDROP");
                Marshal.StructureToPtr(dropFiles, ptrDropFiles, false);
                Marshal.Copy(filenameBytes, 0,
                    IntPtr.Add(ptrDropFiles, dropFilesSize), filenameBytes.Length);
                copied = true;
            }
            finally
            {
                if (ptrDropFiles != IntPtr.Zero)
                    NativeMethods.GlobalUnlock(hGlobal);
                if (!copied)
                    NativeMethods.GlobalFree(hGlobal);
            }

            // Load structure into medium
            medium.unionmember = hGlobal;
            medium.tymed = TYMED.TYMED_HGLOBAL;
            medium.pUnkForRelease = null;        // Caller releases medium via ReleaseStgMedium
        }

        internal static string[] GetFilenames(NativeMethods.IDataObject data)
        {
            log.Debug("Getfilenames");

            //Try Unicode first
            string[] filenames = GetFilenamesUnicode(data);

            //If Unicode returns null, try ANSI
            if (filenames == null)
                filenames = GetFilenamesAnsi(data);

            return filenames;
        }

        internal static string[] GetFilenamesAnsi(NativeMethods.IDataObject data)
        {
            return GetDescriptorFilenames(data, false);
        }

        internal static string[] GetFilenamesUnicode(NativeMethods.IDataObject data)
        {
            return GetDescriptorFilenames(data, true);
        }

        // Limit descriptor metadata, not file contents. Reject untrusted counts before allocation.
        internal const int MaxDescriptorBytes = 16 * 1024 * 1024;

        private sealed class DescriptorBuffer : MemoryStream
        {
            public override void Write(byte[] buffer, int offset, int count)
            {
                if (count < 0 || Position > MaxDescriptorBytes - (long)count)
                    throw new InvalidDataException("File descriptor metadata exceeds the safety limit.");
                base.Write(buffer, offset, count);
            }
        }

        private static string[] GetDescriptorFilenames(NativeMethods.IDataObject data, bool unicode)
        {
            STGMEDIUM medium = new STGMEDIUM();
            try
            {
                FORMATETC format = new FORMATETC();
                format.cfFormat = (short)GetClipboardFormat(unicode ? "FileGroupDescriptorW" : "FileGroupDescriptor");
                format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                format.lindex = -1;
                format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;
                if (data.QueryGetData(ref format) != NativeMethods.S_OK) return null;
                int result = data.GetData(ref format, out medium);
                if (result != NativeMethods.S_OK)
                    throw new COMException("Could not get file descriptors.", result);
                using (var buffer = new DescriptorBuffer())
                {
                    ReadMediumIntoStream(medium, buffer);
                    return ParseFileDescriptors(buffer.ToArray(), unicode);
                }
            }
            finally
            {
                // ReleaseStgMedium also honors a provider's pUnkForRelease controller.
                if (medium.tymed != TYMED.TYMED_NULL) NativeMethods.ReleaseStgMedium(ref medium);
            }
        }

        internal static string[] ParseFileDescriptors(byte[] bytes, bool unicode)
        {
            if (bytes == null || bytes.Length < sizeof(uint) || bytes.Length > MaxDescriptorBytes)
                throw new InvalidDataException("Invalid file descriptor header size.");
            Type descriptorType = unicode ? typeof(NativeMethods.FILEDESCRIPTORW) : typeof(NativeMethods.FILEDESCRIPTORA);
            int descriptorSize = Marshal.SizeOf(descriptorType);
            int nameOffset = Marshal.OffsetOf(descriptorType, "cFileName").ToInt32();
            int nameBytes = unicode ? 520 : 260;
            if (nameOffset != 72 || descriptorSize != nameOffset + nameBytes)
                throw new InvalidDataException("Unexpected native file descriptor layout.");
            uint count = BitConverter.ToUInt32(bytes, 0);
            // ULONG arithmetic must not wrap; compare against actual bytes before allocating.
            ulong required = sizeof(uint) + (ulong)count * (uint)descriptorSize;
            if (required > (ulong)bytes.Length)
                throw new InvalidDataException("Truncated file descriptor array.");
            var filenames = new string[checked((int)count)];
            for (int i = 0; i < filenames.Length; i++)
            {
                int start = checked(sizeof(uint) + i * descriptorSize + nameOffset);
                int end = start;
                int stride = unicode ? 2 : 1;
                int limit = start + nameBytes;
                while (end < limit && (bytes[end] != 0 || (unicode && bytes[end + 1] != 0))) end += stride;
                if (end == limit) throw new InvalidDataException("Unterminated descriptor filename.");
                filenames[i] = (unicode ? System.Text.Encoding.Unicode : System.Text.Encoding.Default)
                    .GetString(bytes, start, end - start);
            }
            return filenames;
        }
        internal static void ReadFileContents(NativeMethods.IDataObject data, int index, Stream stream)
        {
            STGMEDIUM medium = new STGMEDIUM();            
            try
            {
                //Define FileContents format
                FORMATETC format = new FORMATETC();
                format.cfFormat = (short)GetClipboardFormat("FileContents");
                format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                format.lindex = index;
                format.ptd = IntPtr.Zero;
                format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

                //Get data
                int retVal = data.GetData(ref format, out medium);
                if (retVal != NativeMethods.S_OK)
                    throw new Exception(string.Format("Could not get FileContents format.  Error returned: {0}", retVal));

                //Read medium into stream
                ReadMediumIntoStream(medium, stream);
            }
            finally
            {
                //Release all unmanaged objects
                if (medium.tymed != TYMED.TYMED_NULL)
                    NativeMethods.ReleaseStgMedium(ref medium);
            }
        }

        internal static void ReadMediumIntoStream(STGMEDIUM medium, Stream stream)
        {
            switch (medium.tymed)
            {
                case TYMED.TYMED_ISTREAM:
                    ReadIStreamIntoStream(medium.unionmember, stream);
                    break;
                case TYMED.TYMED_ISTORAGE:
                    ReadIStorageIntoStream(medium.unionmember, stream);
                    break;
                case TYMED.TYMED_HGLOBAL:
                    ReadHGlobalIntoStream(medium.unionmember, stream);
                    break;
                default:
                    throw new NotImplementedException(string.Format("Cannot read medium type {0}", medium.tymed));
            }
        }

        private static void ReadIStorageIntoStream(IntPtr handle, Stream stream)
        {
            //To handle a IStorage it needs to be written into a second unmanaged memory mapped storage 
            //and then the data can be read from memory into a managed byte and returned as a MemoryStream

            NativeMethods.ILockBytes iLockBytes = null;
            NativeMethods.IStorage iStorageNew = null;
            NativeMethods.IStorage iStorage = null;
            try
            {
                //Marshal pointer to an IStorage object
                iStorage = (NativeMethods.IStorage)Marshal.GetObjectForIUnknown(handle);

                //Create an ILockBytes object on a HGlobal, then create a IStorage object on top of the ILockBytes object
                iLockBytes = NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true);
                iStorageNew = NativeMethods.StgCreateDocfileOnILockBytes(iLockBytes, 0x00001012, 0);

                //Copy the IStorage object into the new IStorage object
                iStorage.CopyTo(0, null, IntPtr.Zero, iStorageNew);
                iLockBytes.Flush();
                iStorageNew.Commit(0);

                //Get length of ILockBytes byte array
                System.Runtime.InteropServices.ComTypes.STATSTG stat = new System.Runtime.InteropServices.ComTypes.STATSTG();
                iLockBytes.Stat(out stat, 1);
                CopyLockBytes(iLockBytes, stat.cbSize, stream);
            }
            finally
            {
                //Release all unmanaged objects
                if (iStorageNew != null)
                    Marshal.ReleaseComObject(iStorageNew);
                if (iLockBytes != null)
                    Marshal.ReleaseComObject(iLockBytes);
                if (iStorage != null && Marshal.IsComObject(iStorage))
                    Marshal.ReleaseComObject(iStorage);
            }
        }

        internal static void CopyLockBytes(NativeMethods.ILockBytes source, long length, Stream stream)
        {
            if (length < 0) throw new InvalidDataException("Invalid storage length.");
            IntPtr ptrRead = Marshal.AllocCoTaskMem(sizeof(int));
            try
            {
                byte[] buffer = new byte[4096];
                long offset = 0;
                while (offset < length)
                {
                    int requested = (int)Math.Min(buffer.Length, length - offset);
                    Marshal.WriteInt32(ptrRead, 0);
                    source.ReadAt(offset, buffer, requested, ptrRead);
                    int read = Marshal.ReadInt32(ptrRead);
                    if (read <= 0 || read > requested)
                        throw new InvalidDataException("Invalid or incomplete ILockBytes read.");
                    stream.Write(buffer, 0, read);
                    offset = checked(offset + read);
                }
            }
            finally { Marshal.FreeCoTaskMem(ptrRead); }
        }

        private static void ReadIStreamIntoStream(IntPtr handle, Stream stream)
        {
            IntPtr ptrRead = IntPtr.Zero;
            IStream iStream = null;
            try
            {
                //Marshal pointer to an IStream object
                iStream = (IStream)Marshal.GetObjectForIUnknown(handle);

                //Create pointer to integer that stores # of bytes read
                ptrRead = Marshal.AllocCoTaskMem(sizeof(int));

                //Copy IStream into managed stream in chunks
                byte[] buffer = new byte[4096];     //4 KB buffer
                int bytesRead;
                while (true)
                {
                    Marshal.WriteInt32(ptrRead, 0);
                    iStream.Read(buffer, buffer.Length, ptrRead);
                    bytesRead = Marshal.ReadInt32(ptrRead);
                    if (bytesRead < 0 || bytesRead > buffer.Length)
                        throw new InvalidDataException("Invalid IStream read count.");
                    if (bytesRead == 0)
                        break;
                    else
                        stream.Write(buffer, 0, bytesRead);
                }
            }
            finally
            {
                //Release all unmanaged objects
                Marshal.FreeCoTaskMem(ptrRead);
                if (iStream != null && Marshal.IsComObject(iStream))
                    Marshal.ReleaseComObject(iStream);
            }

        }

        private static void ReadHGlobalIntoStream(IntPtr handle, Stream stream)
        {
            ulong size = NativeMethods.GlobalSize(handle).ToUInt64();
            if (size > long.MaxValue || (IntPtr.Size == 4 && size > int.MaxValue))
                throw new IOException("HGLOBAL exceeds the supported process address range.");
            long length = (long)size;
            if (length == 0)
            {
                // A zero-sized movable allocation is valid but cannot be locked.
                if (handle == IntPtr.Zero || NativeMethods.GlobalFlags(handle) == 0x8000)
                    throw new InvalidDataException("Invalid HGLOBAL.");
                return;
            }
            //Lock HGlobal so it cannot be moved or discarded
            IntPtr source = NativeMethods.GlobalLock(handle);

            if (source == IntPtr.Zero)
                throw new Exception(string.Format("Unable to lock hglobal {0}", handle.ToString()));
            try
            {
                //Copy HGlobal into managed stream in chunks
                byte[] buffer = new byte[4096];     //4 KB buffer
                int bytesToCopy;
                for(long offset = 0; offset < length; )
                {
                    //Copy buffer length or remaining length, whichever is smaller
                    bytesToCopy = (int)Math.Min(buffer.Length, length - offset);
                    Marshal.Copy(AddByteOffset(source, offset, bytesToCopy), buffer, 0, bytesToCopy);
                    stream.Write(buffer, 0, bytesToCopy);
                    offset = checked(offset + bytesToCopy);
                }
            }
            finally
            {
                //Release all unmanaged objects
                NativeMethods.GlobalUnlock(handle);
            }
        }

        internal static IntPtr AddByteOffset(IntPtr pointer, long offset, int count)
        {
            ulong address = IntPtr.Size == 4 ? unchecked((uint)pointer.ToInt32()) : unchecked((ulong)pointer.ToInt64());
            ulong maximum = IntPtr.Size == 4 ? uint.MaxValue : ulong.MaxValue;
            if (offset < 0 || count < 0 || (ulong)offset > maximum - address)
                throw new OverflowException("Native pointer offset exceeds the address range.");
            ulong target = address + (ulong)offset;
            if (count > 0 && (ulong)(count - 1) > maximum - target)
                throw new OverflowException("Native buffer exceeds the address range.");
            return IntPtr.Size == 4 ? new IntPtr(unchecked((int)target)) : new IntPtr(unchecked((long)target));
        }
    }
}
