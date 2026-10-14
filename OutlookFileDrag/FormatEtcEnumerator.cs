using System.Runtime.InteropServices.ComTypes;

using System;
using System.Runtime.InteropServices;

namespace OutlookFileDrag
{
    class FormatEtcEnumerator : IEnumFORMATETC 
    {
        private readonly FORMATETC[] formats;
        private readonly byte[][] targetDevices;
        private int index = 0;

        public FormatEtcEnumerator(FORMATETC[] formats)
        {
            if (formats == null) throw new ArgumentNullException(nameof(formats));
            this.formats = (FORMATETC[])formats.Clone();
            targetDevices = new byte[formats.Length][];
            for (int i = 0; i < formats.Length; i++)
            {
                if (formats[i].ptd != IntPtr.Zero)
                {
                    int size = Marshal.ReadInt32(formats[i].ptd);
                    if (size < 12 || size > 1024 * 1024)
                        throw new ArgumentException("Invalid DVTARGETDEVICE size.", nameof(formats));
                    targetDevices[i] = new byte[size];
                    Marshal.Copy(formats[i].ptd, targetDevices[i], 0, size);
                    this.formats[i].ptd = IntPtr.Zero;
                }
            }
        }

        private FormatEtcEnumerator(FORMATETC[] formats, byte[][] targetDevices, int index)
        {
            this.formats = formats;
            this.targetDevices = targetDevices;
            this.index = index;
        }

        public void Clone(out IEnumFORMATETC newEnum)
        {
            //Create new enumerators
            newEnum = new FormatEtcEnumerator(formats, targetDevices, index);
        }

        public int Next(int celt, FORMATETC[] rgelt, int[] pceltFetched)
        {
            if (pceltFetched != null && pceltFetched.Length > 0)
                pceltFetched[0] = 0;
            if (celt < 0 || rgelt == null || rgelt.Length < celt ||
                (pceltFetched != null && pceltFetched.Length == 0) ||
                (celt > 1 && pceltFetched == null))
                return NativeMethods.E_INVALIDARG;
            int count = Math.Min(celt, formats.Length - index);
            int copied = 0;
            try
            {
                for (; copied < count; copied++)
                {
                    rgelt[copied] = formats[index + copied];
                    byte[] device = targetDevices[index + copied];
                    if (device != null)
                    {
                        rgelt[copied].ptd = Marshal.AllocCoTaskMem(device.Length);
                        Marshal.Copy(device, 0, rgelt[copied].ptd, device.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                for (int i = 0; i <= copied && i < count; i++)
                {
                    if (rgelt[i].ptd != IntPtr.Zero) Marshal.FreeCoTaskMem(rgelt[i].ptd);
                    rgelt[i] = new FORMATETC();
                }
                return ex is OutOfMemoryException ? NativeMethods.E_OUTOFMEMORY : NativeMethods.E_UNEXPECTED;
            }
            index += count;
            if (pceltFetched != null) pceltFetched[0] = count;
            return count == celt ? NativeMethods.S_OK : NativeMethods.S_FALSE;
        }

        public int Reset()
        {
            //Set format index back to 0
            index = 0;
            return NativeMethods.S_OK;
        }

        public int Skip(int celt)
        {
            if (celt < 0) return NativeMethods.E_INVALIDARG;
            int count = Math.Min(celt, formats.Length - index);
            index += count;
            return count == celt ? NativeMethods.S_OK : NativeMethods.S_FALSE;
        }
    }
}
