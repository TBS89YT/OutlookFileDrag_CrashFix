using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace OutlookFileDrag
{
    internal static class RegressionTests
    {
        private static int passed, failed;
        private static string root;
        [DllImport("ole32.dll", PreserveSig=false)]
        private static extern void CreateStreamOnHGlobal(IntPtr handle, bool deleteOnRelease, out IStream stream);
        [DllImport("ole32.dll", PreserveSig=false)]
        private static extern void GetHGlobalFromStream(IStream stream, out IntPtr handle);

        [STAThread]
        private static int Main()
        {
            root = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "test-runs", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            Console.WriteLine("Process: " + (IntPtr.Size * 8) + " bit; .NET " + Environment.Version);
            try
            {
                ExportTests(); DescriptorTests(); PathTests(); EnumeratorTests(); StreamTests();
            }
            finally
            {
                // This directory was created exclusively by this process.
                Directory.Delete(root, true);
            }
            Console.WriteLine("RESULT: " + passed + " passed, " + failed + " failed");
            return failed == 0 ? 0 : 1;
        }

        private static void Test(string name, Action action)
        {
            try { action(); passed++; Console.WriteLine("PASS " + name); }
            catch (Exception ex) { failed++; Console.WriteLine("FAIL " + name + ": " + ex); }
        }
        private static void Check(bool condition, string message = "Assertion failed")
        { if (!condition) throw new Exception(message); }
        private static void Throws<T>(Action action) where T : Exception
        {
            try { action(); }
            catch (T) { return; }
            throw new Exception("Expected " + typeof(T).Name);
        }
        private static string NewDirectory()
        {
            string path = Path.Combine(root, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path); return path;
        }
        private static OutlookDataObject Wrap(FakeData data, string parent, Func<string, Stream> writer = null)
        {
            return new OutlookDataObject(data, () => {
                string dir = Path.Combine(parent, Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(dir); return dir;
            }, writer ?? (p => new FileStream(p, FileMode.CreateNew, FileAccess.Write, FileShare.None)));
        }
        private static int GetDrop(OutlookDataObject data, out string[] files)
        {
            var format = new FORMATETC { cfFormat = NativeMethods.CF_HDROP, lindex = -1,
                dwAspect = DVASPECT.DVASPECT_CONTENT, tymed = TYMED.TYMED_HGLOBAL };
            STGMEDIUM medium;
            int result = data.GetData(ref format, out medium);
            files = null;
            try
            {
                if (result != 0) { Check(medium.tymed == TYMED.TYMED_NULL); return result; }
                IntPtr pointer = NativeMethods.GlobalLock(medium.unionmember);
                Check(pointer != IntPtr.Zero);
                try
                {
                    int size = checked((int)NativeMethods.GlobalSize(medium.unionmember).ToUInt64());
                    byte[] bytes = new byte[size]; Marshal.Copy(pointer, bytes, 0, size);
                    int offset = BitConverter.ToInt32(bytes, 0);
                    Check(BitConverter.ToInt32(bytes, 16) == 1, "CF_HDROP must be Unicode");
                    files = Encoding.Unicode.GetString(bytes, offset, size - offset).Split(new[] {'\0'}, StringSplitOptions.RemoveEmptyEntries);
                }
                finally { NativeMethods.GlobalUnlock(medium.unionmember); }
                return result;
            }
            finally { if (medium.tymed != TYMED.TYMED_NULL) NativeMethods.ReleaseStgMedium(ref medium); }
        }
        private static void AssertExport(string[] names)
        {
            var source = new FakeData(names); var wrapper = Wrap(source, NewDirectory());
            string[] files; Check(GetDrop(wrapper, out files) == 0);
            Check(files.Length == names.Length && wrapper.FilesDropped);
            for (int i = 0; i < files.Length; i++)
                Check(File.ReadAllBytes(files[i]).SequenceEqual(FakeData.Content(i)));
        }
        private static void ExportTests()
        {
            Test("Export.One", () => AssertExport(new[] {"mail.msg"}));
            Test("Export.Multiple", () => AssertExport(new[] {"mail.msg", "attachment.pdf", "合同.txt"}));
            foreach (int failIndex in new[] {0, 1})
            {
                int index = failIndex;
                Test("Export.ReadFailureAt" + index + "AndRetry", () => {
                    string parent = NewDirectory(); var source = new FakeData(new[] {"a.msg", "b.msg"}) {FailIndex = index};
                    var wrapper = Wrap(source, parent); string[] files;
                    Check(GetDrop(wrapper, out files) < 0 && !wrapper.FilesDropped);
                    Check(!Directory.EnumerateFileSystemEntries(parent).Any(), "Partial batch not cleaned");
                    Check(GetDrop(wrapper, out files) < 0, "Failed batch incorrectly cached");
                    source.FailIndex = -1; Check(GetDrop(wrapper, out files) == 0);
                    Check(files.All(File.Exists) && files.Length == 2);
                });
            }
            Test("Export.WriteFailureAndRetry", () => {
                string parent = NewDirectory(); var source = new FakeData(new[] {"a.msg", "b.msg"}); bool failWrite = true;
                var wrapper = Wrap(source, parent, p => failWrite ? (Stream)new FailingWriteStream(p) : new FileStream(p, FileMode.CreateNew));
                string[] files; Check(GetDrop(wrapper, out files) < 0 && !wrapper.FilesDropped);
                Check(!Directory.EnumerateFileSystemEntries(parent).Any());
                failWrite = false; Check(GetDrop(wrapper, out files) == 0 && files.Length == 2);
            });
            Test("Export.SuccessCached", () => {
                var source = new FakeData(new[] {"a.msg", "b.msg"}); var wrapper = Wrap(source, NewDirectory());
                string[] first, second; Check(GetDrop(wrapper, out first) == 0); int calls = source.ContentCalls;
                Check(GetDrop(wrapper, out second) == 0 && first.SequenceEqual(second) && source.ContentCalls == calls);
            });
            Test("Export.ParallelCalls", () => {
                var source = new FakeData(new[] {"a.msg", "b.msg"}); var wrapper = Wrap(source, NewDirectory());
                var results = new string[8][];
                Parallel.For(0, 8, i => Check(GetDrop(wrapper, out results[i]) == 0));
                Check(source.ContentCalls == 2 && results.All(x => x.SequenceEqual(results[0])));
            });
            Test("Export.ReentrantCall", () => {
                var source = new FakeData(new[] {"a.msg"}); var wrapper = Wrap(source, NewDirectory());
                source.OnRead = () => { string[] nested; Check(GetDrop(wrapper, out nested) < 0); };
                string[] files; Check(GetDrop(wrapper, out files) == 0 && source.ContentCalls == 1);
            });
            Test("Export.InvalidLaterNameCleansBatch", () => {
                string parent = NewDirectory(); var source = new FakeData(new[] {"valid.msg", "..\\outside.msg"});
                string[] files; Check(GetDrop(Wrap(source, parent), out files) < 0);
                Check(!Directory.EnumerateFileSystemEntries(parent).Any());
            });
            Test("Export.DuplicateNames", () => {
                var source = new FakeData(new[] {"same.msg", "same.msg"}); string[] files;
                Check(GetDrop(Wrap(source, NewDirectory()), out files) == 0);
                Check(files[0] != files[1] && File.ReadAllBytes(files[0]).SequenceEqual(FakeData.Content(0)) &&
                    File.ReadAllBytes(files[1]).SequenceEqual(FakeData.Content(1)));
            });
            Test("Export.EmptyBatchRejected", () => {
                string[] files; Check(GetDrop(Wrap(new FakeData(new string[0]), NewDirectory()), out files) < 0);
            });
            Test("Export.StreamReadFailureCleansBatch", () => {
                string parent = NewDirectory(); var source = new FakeData(new[] {"a.msg"}) {FailStream = true};
                string[] files; Check(GetDrop(Wrap(source, parent), out files) < 0);
                Check(!Directory.EnumerateFileSystemEntries(parent).Any());
            });
        }

        internal static byte[] Descriptors(bool unicode, params string[] names)
        {
            int size = unicode ? 592 : 332;
            byte[] bytes = new byte[4 + names.Length * size];
            Array.Copy(BitConverter.GetBytes((uint)names.Length), bytes, 4);
            for (int i = 0; i < names.Length; i++)
            {
                byte[] name = (unicode ? Encoding.Unicode : Encoding.Default).GetBytes(names[i]);
                Array.Copy(name, 0, bytes, 4 + i * size + 72, name.Length);
            }
            return bytes;
        }
        private static void DescriptorTests()
        {
            foreach (bool unicode in new[] {false, true})
            {
                bool wide = unicode; string prefix = "Descriptor." + (wide ? "Unicode." : "ANSI.");
                Test(prefix + "Valid", () => {
                    var source = new FakeData(new[] {wide ? "合同.msg" : "bericht.msg", "b.txt"}) {Unicode = wide};
                    Check(DataObjectHelper.GetFilenames(source).SequenceEqual(source.Names));
                    Check(source.Handles.All(h => NativeMethods.GlobalFlags(h) == 0x8000), "Medium not freed");
                });
                foreach (int length in new[] {0, 1, 3})
                {
                    int len = length;
                    Test(prefix + "Short" + len, () => Throws<InvalidDataException>(() => DataObjectHelper.ParseFileDescriptors(new byte[len], wide)));
                }
                Test(prefix + "ZeroCount", () => Check(DataObjectHelper.ParseFileDescriptors(new byte[4], wide).Length == 0));
                Test(prefix + "ExcessCount", () => {
                    byte[] bytes = Descriptors(wide, "a.msg"); Array.Copy(BitConverter.GetBytes(2U), bytes, 4);
                    Throws<InvalidDataException>(() => DataObjectHelper.ParseFileDescriptors(bytes, wide));
                });
                Test(prefix + "MaxCount", () => Throws<InvalidDataException>(() => DataObjectHelper.ParseFileDescriptors(BitConverter.GetBytes(uint.MaxValue), wide)));
                Test(prefix + "TruncatedLast", () => {
                    byte[] bytes = Descriptors(wide, "a.msg", "b.msg"); Array.Resize(ref bytes, bytes.Length - 1);
                    Throws<InvalidDataException>(() => DataObjectHelper.ParseFileDescriptors(bytes, wide));
                });
                Test(prefix + "WrongLayout", () => Throws<InvalidDataException>(() => DataObjectHelper.ParseFileDescriptors(Descriptors(false, "a.msg"), true)));
                Test(prefix + "Unterminated", () => {
                    byte[] bytes = Descriptors(wide, "a.msg");
                    for (int i = 76; i < bytes.Length; i++) bytes[i] = 65;
                    Throws<InvalidDataException>(() => DataObjectHelper.ParseFileDescriptors(bytes, wide));
                });
                Test(prefix + "FailureReleasesMedium", () => {
                    var source = new FakeData(new[] {"a.msg"}) {Unicode = wide, DescriptorOverride = BitConverter.GetBytes(uint.MaxValue)};
                    Throws<InvalidDataException>(() => DataObjectHelper.GetFilenames(source));
                    Check(source.Handles.All(h => NativeMethods.GlobalFlags(h) == 0x8000));
                });
            }
            Test("Descriptor.MetadataLimit", () => Throws<InvalidDataException>(() => DataObjectHelper.ParseFileDescriptors(new byte[DataObjectHelper.MaxDescriptorBytes + 1], true)));
            Test("Descriptor.MetadataStreamLimitAndRelease", () => {
                var source = new FakeData(new[] {"a.msg"}) {DescriptorOverride = new byte[DataObjectHelper.MaxDescriptorBytes + 4096]};
                Throws<InvalidDataException>(() => DataObjectHelper.GetFilenames(source));
                Check(source.Handles.All(h => NativeMethods.GlobalFlags(h) == 0x8000));
            });
            Test("Descriptor.NativeLayouts", () => {
                Check(Marshal.SizeOf(typeof(NativeMethods.FILEDESCRIPTORA)) == 332);
                Check(Marshal.SizeOf(typeof(NativeMethods.FILEDESCRIPTORW)) == 592);
            });
            Test("Descriptor.ReleaseControllerOnFailure", () => {
                var source=new FakeData(new[]{"a.msg"}) {UseReleaseController=true,DescriptorOverride=BitConverter.GetBytes(uint.MaxValue)};
                try {
                    Throws<InvalidDataException>(()=>DataObjectHelper.GetFilenames(source));
                    Check(NativeMethods.GlobalFlags(source.Handles[0])!=0x8000,"Controller-owned handle freed directly");
                } finally {if(source.ReleaseController!=null) Marshal.ReleaseComObject(source.ReleaseController);}
                Check(NativeMethods.GlobalFlags(source.Handles[0])==0x8000,"Controller-owned HGLOBAL leaked");
            });
        }
        private static void PathTests()
        {
            foreach (string name in new[] {"normal.msg", "合同–Grüße.msg", "report.v1.pdf", "😀.txt"})
            {
                string value = name;
                Test("Path.Valid." + value, () => {
                    string dir = NewDirectory(); string path = FileUtility.GetSafeExportPath(dir, value);
                    Check(Path.GetDirectoryName(path) == dir && Path.GetFileName(path) == value);
                });
            }
            foreach (string name in new[] {"..\\evil.msg", "../evil.msg", "C:\\evil.msg", "C:evil.msg", "\\\\server\\share\\evil.msg",
                "/evil.msg", "folder/file.msg", "folder\\file.msg", "a:b.msg", "a?b.msg", "a|b.msg", "a\0b.msg",
                "CON", "con.txt", "NUL.msg", "COM1.txt", "LPT9", "COM¹.txt", "CONIN$", "AUX .txt", "", " ", ".", "..", "trailing.", "trailing ", new string('a',256)})
            {
                string value = name;
                Test("Path.Reject." + value.Replace('\0', '?'), () => Throws<InvalidDataException>(() => FileUtility.GetSafeExportPath(NewDirectory(), value)));
            }
            Test("Path.CollisionPreservesExisting", () => {
                string dir = NewDirectory(); string old = Path.Combine(dir, "mail.msg"); File.WriteAllText(old, "keep");
                string path = FileUtility.GetSafeExportPath(dir, "mail.msg");
                Check(path != old && File.ReadAllText(old) == "keep" && Path.GetDirectoryName(path) == dir);
            });
            Test("Path.DirectoryCollision", () => {
                string dir = NewDirectory(); Directory.CreateDirectory(Path.Combine(dir, "mail.msg"));
                Check(Path.GetFileName(FileUtility.GetSafeExportPath(dir, "mail.msg")) == "mail (1).msg");
            });
            Test("Path.LongNameCollision", () => {
                string dir = NewDirectory(); string name = new string('a',240) + ".msg";
                string first = FileUtility.GetSafeExportPath(dir, name); File.WriteAllText(first, "keep");
                string second = FileUtility.GetSafeExportPath(dir, name);
                Check(first.Length < 260 && second.Length < 260 && first != second && Path.GetDirectoryName(second) == dir);
            });
            Test("Path.ExcessiveExtension", () => Throws<PathTooLongException>(() => FileUtility.GetSafeExportPath(NewDirectory(), "a." + new string('x',240))));
        }

        private static FormatEtcEnumerator Enumerator()
        { return new FormatEtcEnumerator(new[] {new FORMATETC {cfFormat=10},new FORMATETC {cfFormat=11},new FORMATETC {cfFormat=12}}); }
        private static short Next(IEnumFORMATETC e)
        { var item = new FORMATETC[1]; Check(e.Next(1,item,null)==0); return item[0].cfFormat; }
        private static void EnumeratorTests()
        {
            foreach (int position in new[] {0,1,2})
            {
                int pos = position;
                Test("Enum.CloneAt" + pos, () => {
                    var e=Enumerator(); for(int i=0;i<pos;i++) Next(e);
                    IEnumFORMATETC clone; e.Clone(out clone);
                    Check(Next(e)==10+pos && Next(clone)==10+pos);
                    clone.Reset(); Check(Next(clone)==10); Check(e.Skip(3)==1);
                });
            }
            Test("Enum.SkipWithin", () => { var e=Enumerator(); Check(e.Skip(1)==0 && Next(e)==11); });
            Test("Enum.SkipExactEnd", () => {var e=Enumerator(); Check(e.Skip(3)==0 && e.Next(1,new FORMATETC[1],null)==1);});
            Test("Enum.SkipBeyondEnd", () => {var e=Enumerator(); Check(e.Skip(int.MaxValue)==1 && e.Next(1,new FORMATETC[1],null)==1);});
            Test("Enum.ResetAfterEnd", () => {var e=Enumerator(); e.Skip(4); Check(e.Reset()==0 && Next(e)==10);});
            Test("Enum.NextAtEndFetchedZero", () => {var e=Enumerator(); e.Skip(3); var count=new[]{99}; Check(e.Next(1,new FORMATETC[1],count)==1 && count[0]==0);});
            Test("Enum.PartialNext", () => {var e=Enumerator(); var count=new[]{99}; Check(e.Next(4,new FORMATETC[4],count)==1 && count[0]==3);});
            Test("Enum.InvalidArgumentsDoNotAdvance", () => {
                var e=Enumerator(); Check(e.Skip(-1)<0 && e.Next(-1,new FORMATETC[1],null)<0 && e.Next(1,null,null)<0);
                Check(e.Next(2,new FORMATETC[2],null)<0 && e.Next(2,new FORMATETC[1],new int[1])<0);
                Check(e.Next(1,new FORMATETC[1],new int[0])<0 && Next(e)==10);
            });
            Test("Enum.ZeroCounts", () => {var e=Enumerator(); Check(e.Next(0,new FORMATETC[0],null)==0 && e.Skip(0)==0 && Next(e)==10);});
            Test("Enum.Empty", () => {var e=new FormatEtcEnumerator(new FORMATETC[0]); Check(e.Skip(1)==1 && e.Skip(0)==0 && e.Next(1,new FORMATETC[1],null)==1);});
            Test("Enum.TargetDeviceCopies", () => {
                IntPtr original=Marshal.AllocCoTaskMem(12); FormatEtcEnumerator e;
                try {Marshal.Copy(new byte[12],0,original,12); Marshal.WriteInt32(original,12);
                    e=new FormatEtcEnumerator(new[]{new FORMATETC {cfFormat=10,ptd=original}});
                } finally {Marshal.FreeCoTaskMem(original);}
                IEnumFORMATETC clone; e.Clone(out clone); var a=new FORMATETC[1]; var b=new FORMATETC[1];
                try {Check(e.Next(1,a,null)==0 && clone.Next(1,b,null)==0);
                    Check(a[0].ptd!=IntPtr.Zero && b[0].ptd!=a[0].ptd && Marshal.ReadInt32(a[0].ptd)==12);
                } finally {Marshal.FreeCoTaskMem(a[0].ptd);Marshal.FreeCoTaskMem(b[0].ptd);}
            });
            Test("Enum.WrapperFormats", () => {
                var source=new FakeData(new[]{"a.msg"}); source.Formats=new[] {
                    new FORMATETC {cfFormat=NativeMethods.CF_TEXT}, new FORMATETC {cfFormat=NativeMethods.CF_UNICODETEXT},
                    new FORMATETC {cfFormat=(short)DataObjectHelper.GetClipboardFormat("FileContents")}};
                IEnumFORMATETC e; Check(Wrap(source,NewDirectory()).EnumFormatEtc(DATADIR.DATADIR_GET,out e)==0);
                var items=new FORMATETC[4]; var count=new int[1]; Check(e.Next(4,items,count)==1 && count[0]==2);
                Check(items[0].cfFormat==source.Formats[2].cfFormat && items[1].cfFormat==NativeMethods.CF_HDROP);
            });
        }

        private static void StreamTests()
        {
            foreach (int size in new[] {0,16,4096,4112,8192})
            {
                int length=size;
                Test("HGlobal.Copy"+size, () => {
                    byte[] bytes=Enumerable.Range(0,length).Select(i=>(byte)i).ToArray(); IntPtr handle=FakeData.Alloc(bytes);
                    try { using(var dest=new MemoryStream()) {
                        DataObjectHelper.ReadMediumIntoStream(new STGMEDIUM {tymed=TYMED.TYMED_HGLOBAL,unionmember=handle},dest);
                        Check(dest.ToArray().SequenceEqual(bytes));
                    } Check((NativeMethods.GlobalFlags(handle)&0xFF)==0,"HGLOBAL left locked"); }
                    finally {NativeMethods.GlobalFree(handle);}
                });
                Test("LockBytes.Copy"+size, () => {
                    var source=new FakeLockBytes(length); var dest=new CountingStream();
                    DataObjectHelper.CopyLockBytes(source,length,dest); Check(dest.Length==length && source.Offset==length);
                });
            }
            Test("HGlobal.WriteFailureUnlocks", () => {
                IntPtr handle=FakeData.Alloc(new byte[8192]);
                try {
                    Throws<IOException>(()=>DataObjectHelper.ReadMediumIntoStream(new STGMEDIUM {tymed=TYMED.TYMED_HGLOBAL,unionmember=handle},new CountingStream {Fail=true}));
                    Check((NativeMethods.GlobalFlags(handle)&0xFF)==0);
                } finally {NativeMethods.GlobalFree(handle);}
            });
            Test("HGlobal.InvalidHandle", () => Throws<InvalidDataException>(()=>DataObjectHelper.ReadMediumIntoStream(new STGMEDIUM {tymed=TYMED.TYMED_HGLOBAL},new CountingStream())));
            Test("Pointer.Over2GB", () => {
                IntPtr pointer=new IntPtr(0x10000); long offset=(long)int.MaxValue+4097;
                IntPtr result=DataObjectHelper.AddByteOffset(pointer,offset,4096);
                ulong actual=IntPtr.Size==4?unchecked((uint)result.ToInt32()):unchecked((ulong)result.ToInt64());
                Check(actual==0x10000UL+(ulong)offset);
            });
            Test("Pointer.AddressOverflow", () => Throws<OverflowException>(()=>DataObjectHelper.AddByteOffset(new IntPtr(-1),1,1)));
            Test("Pointer.BufferOverflow", () => Throws<OverflowException>(()=>DataObjectHelper.AddByteOffset(new IntPtr(-2),0,3)));
            Test("Pointer.NegativeOffset", () => Throws<OverflowException>(()=>DataObjectHelper.AddByteOffset(new IntPtr(1),-1,1)));
            Test("LockBytes.Over2GBWithoutLargeAllocation", () => {
                long length=(long)int.MaxValue+8193; var source=new FakeLockBytes(length); var dest=new CountingStream();
                DataObjectHelper.CopyLockBytes(source,length,dest);
                Check(source.Crossed2GB && dest.Length==length && source.Offset==length);
            });
            Test("LockBytes.PartialReads", () => {
                var source=new FakeLockBytes(9000) {MaxRead=17}; var dest=new CountingStream();
                DataObjectHelper.CopyLockBytes(source,9000,dest); Check(dest.Length==9000);
            });
            Test("LockBytes.PrematureEnd", () => Throws<InvalidDataException>(()=>DataObjectHelper.CopyLockBytes(new FakeLockBytes(10),20,new CountingStream())));
            Test("LockBytes.InvalidCount", () => Throws<InvalidDataException>(()=>DataObjectHelper.CopyLockBytes(new FakeLockBytes(10) {InvalidCount=true},10,new CountingStream())));
            Test("LockBytes.ReadFailure", () => Throws<IOException>(()=>DataObjectHelper.CopyLockBytes(new FakeLockBytes(8192) {Fail=true},8192,new CountingStream())));
            Test("LockBytes.WriteFailure", () => Throws<IOException>(()=>DataObjectHelper.CopyLockBytes(new FakeLockBytes(8192),8192,new CountingStream {Fail=true})));
            Test("LockBytes.NegativeLength", () => Throws<InvalidDataException>(()=>DataObjectHelper.CopyLockBytes(new FakeLockBytes(0),-1,new CountingStream())));
            Test("LockBytes.NativeInterop", () => {
                NativeMethods.ILockBytes source=NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero,true);
                IntPtr data=Marshal.AllocCoTaskMem(4112);
                try {
                    byte[] expected=Enumerable.Range(0,4112).Select(i=>(byte)i).ToArray(); Marshal.Copy(expected,0,data,expected.Length);
                    int[] written=new int[1]; source.WriteAt(0,data,expected.Length,written); Check(written[0]==expected.Length);
                    using(var dest=new MemoryStream()) {DataObjectHelper.CopyLockBytes(source,expected.Length,dest); Check(expected.SequenceEqual(dest.ToArray()));}
                } finally {Marshal.FreeCoTaskMem(data); Marshal.ReleaseComObject(source);}
            });
            Test("IStorage.NativeCopyAndFailureRecovery", () => {
                NativeMethods.ILockBytes source=NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero,true);
                NativeMethods.IStorage storage=null; IntPtr pointer=IntPtr.Zero;
                try {
                    storage=NativeMethods.StgCreateDocfileOnILockBytes(source,0x1012,0);
                    IStream content=storage.CreateStream("content",0x1012,0,0);
                    try {content.Write(new byte[8192],8192,IntPtr.Zero);content.Commit(0);}
                    finally {Marshal.ReleaseComObject(content);}
                    storage.Commit(0); pointer=Marshal.GetIUnknownForObject(storage);
                    var medium=new STGMEDIUM {tymed=TYMED.TYMED_ISTORAGE,unionmember=pointer};
                    Throws<IOException>(()=>DataObjectHelper.ReadMediumIntoStream(medium,new CountingStream {Fail=true}));
                    using(var dest=new MemoryStream()) {
                        DataObjectHelper.ReadMediumIntoStream(medium,dest);
                        byte[] bytes=dest.ToArray(); Check(bytes.Length>=8192 && bytes[0]==0xD0 && bytes[1]==0xCF);
                    }
                    STATSTG stat; storage.Stat(out stat,1); Check(stat.type==1,"Source RCW disconnected");
                } finally {
                    if(pointer!=IntPtr.Zero) Marshal.Release(pointer);
                    if(storage!=null) Marshal.ReleaseComObject(storage);
                    Marshal.ReleaseComObject(source);
                }
            });
            foreach(bool fail in new[]{false,true}) {
                bool failWrite=fail;
                Test("IStream.NativeCopyAndRelease."+fail, () => {
                    IStream source; CreateStreamOnHGlobal(IntPtr.Zero,true,out source);
                    IntPtr pointer=IntPtr.Zero,handle=IntPtr.Zero;
                    try {
                        byte[] expected=Enumerable.Range(0,4112).Select(i=>(byte)i).ToArray(); source.Write(expected,expected.Length,IntPtr.Zero);
                        source.Seek(0,0,IntPtr.Zero); GetHGlobalFromStream(source,out handle);
                        pointer=Marshal.GetIUnknownForObject(source);
                        var medium=new STGMEDIUM {tymed=TYMED.TYMED_ISTREAM,unionmember=pointer};
                        if(failWrite) Throws<IOException>(()=>DataObjectHelper.ReadMediumIntoStream(medium,new CountingStream {Fail=true}));
                        else using(var dest=new MemoryStream()) {DataObjectHelper.ReadMediumIntoStream(medium,dest);Check(expected.SequenceEqual(dest.ToArray()));}
                        STATSTG stat;source.Stat(out stat,1);Check(stat.cbSize==expected.Length);
                    } finally {if(pointer!=IntPtr.Zero) Marshal.Release(pointer);Marshal.ReleaseComObject(source);}
                    Check(NativeMethods.GlobalFlags(handle)==0x8000,"IStream reference leaked");
                });
            }
        }
    }

    internal sealed class FakeData : NativeMethods.IDataObject
    {
        internal readonly string[] Names;
        internal bool Unicode=true, FailStream, UseReleaseController;
        internal NativeMethods.ILockBytes ReleaseController;
        internal int FailIndex=-1, ContentCalls;
        internal byte[] DescriptorOverride;
        internal Action OnRead;
        internal FORMATETC[] Formats;
        internal readonly List<IntPtr> Handles=new List<IntPtr>();
        internal FakeData(string[] names) {Names=names;}
        internal static byte[] Content(int index) {return Enumerable.Repeat((byte)(index+1),16).ToArray();}
        internal static IntPtr Alloc(byte[] bytes)
        {
            IntPtr handle=NativeMethods.GlobalAlloc(NativeMethods.GMEM_MOVEABLE,new UIntPtr((uint)bytes.Length));
            if(handle==IntPtr.Zero) throw new OutOfMemoryException();
            if(bytes.Length>0) {IntPtr p=NativeMethods.GlobalLock(handle); try {Marshal.Copy(bytes,0,p,bytes.Length);} finally {NativeMethods.GlobalUnlock(handle);}}
            return handle;
        }
        public int QueryGetData(ref FORMATETC format)
        {return format.cfFormat==(short)DataObjectHelper.GetClipboardFormat(Unicode?"FileGroupDescriptorW":"FileGroupDescriptor")?0:NativeMethods.DV_E_FORMATETC;}
        public int GetData(ref FORMATETC format,out STGMEDIUM medium)
        {
            medium=new STGMEDIUM();
            if(QueryGetData(ref format)==0)
            {
                medium=new STGMEDIUM {tymed=TYMED.TYMED_HGLOBAL,unionmember=Alloc(DescriptorOverride??RegressionTests.Descriptors(Unicode,Names))};
                if(UseReleaseController) {
                    ReleaseController=NativeMethods.CreateILockBytesOnHGlobal(medium.unionmember,true);
                    medium.pUnkForRelease=ReleaseController;
                }
            }
            else if(format.cfFormat==(short)DataObjectHelper.GetClipboardFormat("FileContents"))
            {
                Interlocked.Increment(ref ContentCalls); if(OnRead!=null) OnRead();
                if(format.lindex==FailIndex) return NativeMethods.E_FAIL;
                if(FailStream) {medium=new STGMEDIUM {tymed=TYMED.TYMED_ISTREAM,unionmember=Marshal.GetComInterfaceForObject(new BrokenReadStream(),typeof(IStream))}; return 0;}
                medium=new STGMEDIUM {tymed=TYMED.TYMED_HGLOBAL,unionmember=Alloc(Content(format.lindex))};
            }
            else return NativeMethods.DV_E_FORMATETC;
            lock(Handles) Handles.Add(medium.unionmember); return 0;
        }
        public int GetDataHere(ref FORMATETC f,ref STGMEDIUM m) {return NativeMethods.E_NOTIMPL;}
        public int GetCanonicalFormatEtc(ref FORMATETC f,out FORMATETC o) {o=f;return NativeMethods.DATA_S_SAMEFORMATETC;}
        public int SetData(ref FORMATETC f,ref STGMEDIUM m,bool release) {return NativeMethods.E_NOTIMPL;}
        public int EnumFormatEtc(DATADIR d,out IEnumFORMATETC e) {e=Formats==null?null:new FormatEtcEnumerator(Formats);return Formats==null?NativeMethods.E_NOTIMPL:0;}
        public int DAdvise(ref FORMATETC f,ADVF a,IAdviseSink s,out int c) {c=0;return NativeMethods.E_NOTIMPL;}
        public int DUnadvise(int c) {return NativeMethods.E_NOTIMPL;}
        public int EnumDAdvise(out IEnumSTATDATA e) {e=null;return NativeMethods.E_NOTIMPL;}
    }
    internal sealed class FailingWriteStream : Stream
    {
        private readonly FileStream file;
        internal FailingWriteStream(string path) {file=new FileStream(path,FileMode.CreateNew);}
        public override void Write(byte[] buffer,int offset,int count) {file.Write(buffer,offset,Math.Min(1,count));throw new IOException("Simulated disk failure");}
        protected override void Dispose(bool disposing) {if(disposing) file.Dispose();base.Dispose(disposing);}
        public override bool CanRead {get{return false;}} public override bool CanSeek {get{return false;}} public override bool CanWrite {get{return true;}}
        public override long Length {get{return file.Length;}} public override long Position {get{return file.Position;}set{throw new NotSupportedException();}}
        public override void Flush() {file.Flush();} public override int Read(byte[] b,int o,int c) {throw new NotSupportedException();}
        public override long Seek(long o,SeekOrigin s) {throw new NotSupportedException();} public override void SetLength(long v) {throw new NotSupportedException();}
    }
    internal sealed class CountingStream : Stream
    {
        private long length; internal bool Fail;
        public override void Write(byte[] buffer,int offset,int count) {if(Fail) throw new IOException("Simulated write failure");length=checked(length+count);}
        public override bool CanRead {get{return false;}} public override bool CanSeek {get{return false;}} public override bool CanWrite {get{return true;}}
        public override long Length {get{return length;}} public override long Position {get{return length;}set{throw new NotSupportedException();}}
        public override void Flush() {} public override int Read(byte[] b,int o,int c) {throw new NotSupportedException();}
        public override long Seek(long o,SeekOrigin s) {throw new NotSupportedException();} public override void SetLength(long v) {throw new NotSupportedException();}
    }
    internal sealed class FakeLockBytes : NativeMethods.ILockBytes
    {
        private readonly long length; internal long Offset; internal bool Crossed2GB,Fail,InvalidCount; internal int MaxRead=int.MaxValue;
        internal FakeLockBytes(long length) {this.length=length;}
        public void ReadAt(long offset,byte[] bytes,int count,IntPtr read)
        {
            if(Fail) throw new IOException("Simulated native read failure");
            if(offset!=Offset) throw new Exception("Incorrect 64-bit offset: "+offset+" expected "+Offset);
            Crossed2GB |= offset>int.MaxValue;
            int actual=(int)Math.Min(Math.Min(count,MaxRead),Math.Max(0,length-offset));
            Offset+=actual; Marshal.WriteInt32(read,InvalidCount?count+1:actual);
        }
        public void WriteAt(long o,IntPtr b,int c,int[] w) {throw new NotSupportedException();} public void Flush() {}
        public void SetSize(long s) {throw new NotSupportedException();} public void LockRegion(long o,long c,int t) {throw new NotSupportedException();}
        public void UnlockRegion(long o,long c,int t) {throw new NotSupportedException();}
        public void Stat(out STATSTG stat,int flag) {stat=new STATSTG {cbSize=length};}
    }
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class BrokenReadStream : IStream
    {
        private bool first=true;
        public void Read(byte[] pv,int cb,IntPtr read) {if(!first) throw new IOException("Simulated IStream read failure");first=false;Marshal.WriteInt32(read,cb);}
        public void Write(byte[] b,int c,IntPtr w) {throw new NotSupportedException();}
        public void Seek(long o,int origin,IntPtr n) {throw new NotSupportedException();} public void SetSize(long size) {throw new NotSupportedException();}
        public void CopyTo(IStream s,long cb,IntPtr r,IntPtr w) {throw new NotSupportedException();} public void Commit(int flags) {}
        public void Revert() {throw new NotSupportedException();} public void LockRegion(long o,long c,int t) {throw new NotSupportedException();}
        public void UnlockRegion(long o,long c,int t) {throw new NotSupportedException();} public void Stat(out STATSTG s,int f) {s=new STATSTG();}
        public void Clone(out IStream s) {s=null;throw new NotSupportedException();}
    }
}
