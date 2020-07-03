using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace SharpStick.Win32
{
    class ole32
    {
        [DllImport("ole32.dll")]
        public static extern int StgIsStorageFile([MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

        [DllImport("ole32.dll")]
        public static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)]string pwcsName,
            IStorage pstgPriority, STGM grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstgOpen);

        [Flags]
        public enum STGM : int
        {
            DIRECT = 0x00000000,
            TRANSACTED = 0x00010000,
            SIMPLE = 0x08000000,
            READ = 0x00000000,
            WRITE = 0x00000001,
            READWRITE = 0x00000002,
            SHARE_DENY_NONE = 0x00000040,
            SHARE_DENY_READ = 0x00000030,
            SHARE_DENY_WRITE = 0x00000020,
            SHARE_EXCLUSIVE = 0x00000010,
            PRIORITY = 0x00040000,
            DELETEONRELEASE = 0x04000000,
            NOSCRATCH = 0x00100000,
            CREATE = 0x00001000,
            CONVERT = 0x00020000,
            FAILIFTHERE = 0x00000000,
            NOSNAPSHOT = 0x00200000,
            DIRECT_SWMR = 0x00400000,
        }

        [ComImport]
        [Guid("0000000b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IStorage
        {
            void CreateStream(
                /* [string][in] */ string pwcsName,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void OpenStream(
                /* [string][in] */ string pwcsName,
                /* [unique][in] */ IntPtr reserved1,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void CreateStorage(
                /* [string][in] */ string pwcsName,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStorage ppstg);

            void OpenStorage(
                /* [string][unique][in] */ string pwcsName,
                /* [unique][in] */ IStorage pstgPriority,
                /* [in] */ uint grfMode,
                /* [unique][in] */ IntPtr snbExclude,
                /* [in] */ uint reserved,
                /* [out] */ out IStorage ppstg);

            void CopyTo(
                /* [in] */ uint ciidExclude,
                /* [size_is][unique][in] */ [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] Guid[] rgiidExclude,
                /* [unique][in] */ IntPtr snbExclude,
                /* [unique][in] */ IStorage pstgDest);

            void MoveElementTo(
                /* [string][in] */ string pwcsName,
                /* [unique][in] */ IStorage pstgDest,
                /* [string][in] */ string pwcsNewName,
                /* [in] */ uint grfFlags);

            void Commit(
                /* [in] */ uint grfCommitFlags);

            void Revert();

            void EnumElements(
                /* [in] */ uint reserved1,
                /* [size_is][unique][in] */ IntPtr reserved2,
                /* [in] */ uint reserved3,
                /* [out] */ out IEnumSTATSTG ppenum);

            void DestroyElement(
                /* [string][in] */ string pwcsName);

            void RenameElement(
                /* [string][in] */ string pwcsOldName,
                /* [string][in] */ string pwcsNewName);

            void SetElementTimes(
                /* [string][unique][in] */ string pwcsName,
                /* [unique][in] */
                System.Runtime.InteropServices.ComTypes.FILETIME pctime,
                /* [unique][in] */
                System.Runtime.InteropServices.ComTypes.FILETIME patime,
                /* [unique][in] */
                System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

            void SetClass(
                /* [in] */ Guid clsid);

            void SetStateBits(
                /* [in] */ uint grfStateBits,
                /* [in] */ uint grfMask);

            void Stat(
                /* [out] */ out
                System.Runtime.InteropServices.
                ComTypes.STATSTG pstatstg,
                /* [in] */ uint grfStatFlag);

        }

        [ComImport]
        [Guid("0000000d-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumSTATSTG
        {
            // The user needs to allocate an STATSTG array whose size is celt.
            [PreserveSig]
            uint Next(
                uint celt,
                [MarshalAs(UnmanagedType.LPArray),
                Out]
                System.Runtime.InteropServices.
                ComTypes.STATSTG[] rgelt,
            out uint pceltFetched
        );

            void Skip(uint celt);

            void Reset();

            [return: MarshalAs(UnmanagedType.Interface)]
            IEnumSTATSTG Clone();
        }
    }
}
