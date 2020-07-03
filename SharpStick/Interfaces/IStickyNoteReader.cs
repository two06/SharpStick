using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharpStick.Interfaces
{
    public interface IStickyNoteReader
    {
        IEnumerable<string> GetNotes(string path);
    }
}
