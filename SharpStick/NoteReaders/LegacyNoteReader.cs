using SharpStick.Interfaces;
using System;
using System.Collections.Generic;
using SharpStick.Win32;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace SharpStick.NoteReaders
{
    class LegacyNoteReader : IStickyNoteReader
    {
        public IEnumerable<string> GetNotes(string path)
        {
            var data = new List<string>();
            var isOLE = ole32.StgIsStorageFile(path);
            if (isOLE == 0)
            {
                //open the storage 
                ole32.IStorage Is;
                int result = ole32.StgOpenStorage(path, null, ole32.STGM.READ | ole32.STGM.SHARE_DENY_WRITE, IntPtr.Zero, 0, out Is);
                //set up to fetch one item on each call to next
                ole32.IEnumSTATSTG SSenum;
                Is.EnumElements(0, IntPtr.Zero, 0, out SSenum);
                var SSstruct = new System.Runtime.InteropServices.ComTypes.STATSTG[1];

                //do the loop until not more items
                uint NumReturned;
                do
                {
                    SSenum.Next(1, SSstruct, out NumReturned);
                    if (NumReturned != 0)
                    {
                        if (SSstruct[0].type == 1)
                        {
                            OpenSubStorage(Is, SSstruct[0].pwcsName, data);
                        }

                    }
                } while (NumReturned > 0);
            }
            return data;
        }

        //No problem cant be made worse with recursion! 
        private List<string> OpenSubStorage(ole32.IStorage Is, string pwcsName, List<string> data)
        {
            ole32.IStorage ppstg;
            Is.OpenStorage(pwcsName, null, (uint)(ole32.STGM.READ | ole32.STGM.SHARE_EXCLUSIVE), IntPtr.Zero, 0, out ppstg);

            //set up to fetch one item on each call to next
            ole32.IEnumSTATSTG SSenum;
            ppstg.EnumElements(0, IntPtr.Zero, 0, out SSenum);
            var SSstruct = new System.Runtime.InteropServices.ComTypes.STATSTG[1];

            //do the loop until not more items
            uint NumReturned;
            do
            {
                SSenum.Next(1, SSstruct, out NumReturned);
                if (NumReturned != 0)
                {
                    if (SSstruct[0].type == 1)
                    {
                        OpenSubStorage(ppstg, SSstruct[0].pwcsName, data);
                    }
                    else if (SSstruct[0].type == 2 && SSstruct[0].pwcsName == "3")
                    {
                        data.Add(readStream(ref ppstg, SSstruct[0].pwcsName));
                    }
                }
            } while (NumReturned > 0);

            return data;
        }

        private string readStream(ref ole32.IStorage Is, string pwcsName)
        {
            IStream stream;
            byte[] buf = new byte[1000];
            IntPtr readBuffer = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));

            Is.OpenStream(pwcsName, IntPtr.Zero, (uint)(ole32.STGM.READ | ole32.STGM.SHARE_EXCLUSIVE), 0, out stream);
            stream.Read(buf, 1000, readBuffer);
            int intValue = Marshal.ReadInt32(readBuffer);
            Marshal.FreeCoTaskMem(readBuffer);

            //only print the number of bytes we actually read. Lets just assume we never read more than 1000
            return System.Text.Encoding.Unicode.GetString(buf.Take(intValue - 2).ToArray());
        }
    }
}
