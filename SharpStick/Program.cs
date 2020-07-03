using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using SharpStick.Interfaces;
using SharpStick.NoteReaders;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharpStick
{
    class Program
    {
        static void Main(string[] args)
        {
            //Gotta have that ASCII art!
            PrintHeader();

            //Check which version of Windows we are running under
            string releaseId = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId", "").ToString();
            int releaseID;
            if(!int.TryParse(releaseId, out releaseID))
            {
                Console.WriteLine("[*] Could not get Windows Release ID!");
                Console.WriteLine("[*] Defaulting to COM Structured Storage Reader...");
                releaseID = 1;
            }
            //if its > 1607, we need to use the SQLite reader, otherwise its the legacy COM Structured Storage way
            IStickyNoteReader reader;
            string path;

            if(releaseID > 1607)
            {
                Console.WriteLine("[*] Windows release is later than 1607 - Using SQLite reader");
                reader = new SQLiteNoteReader();
                path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) 
                    + @"\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite";
            }
            else
            {
                Console.WriteLine("[*] Windows release is prior to Win 10 1607 - Using COM Structured Storate reader");
                reader = new LegacyNoteReader();
                path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) 
                    + @"\Microsoft\Sticky Notes\StickyNotes.snt";
            }
            
            if (!File.Exists(path))
            {
                Console.WriteLine("[*] StickNotes DB not found!");
                return;
            }
            try
            {
                var results = reader.GetNotes(path);
                if (! results.Any())
                {
                    Console.WriteLine("[*] No notes found!");
                    return;
                }
                Console.WriteLine("[*] Printing notes...");
                foreach (var result in results)
                {
                    Console.WriteLine("\t" + result);
                }
                Console.WriteLine("[*] Completed!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("[*] Exception occured reading StickyNotes DB!");
                Console.WriteLine(ex.ToString());
            }
            
        }

        private static void PrintHeader()
        {
            Console.WriteLine(@" _____ _                      _____ _   _      _    ");
            Console.WriteLine(@"/  ___| |                    /  ___| | (_)    | |   ");
            Console.WriteLine(@"\ `--.| |__   __ _ _ __ _ __ \ `--.| |_ _  ___| | __");
            Console.WriteLine(@" `--. \ '_ \ / _` | '__| '_ \ `--. \ __| |/ __| |/ /");
            Console.WriteLine(@"/\__/ / | | | (_| | |  | |_) /\__/ / |_| | (__|   < ");
            Console.WriteLine(@"\____/|_| |_|\__,_|_|  | .__/\____/ \__|_|\___|_|\_\");
            Console.WriteLine(@"                       | |                          ");
            Console.WriteLine(@"                       |_|                          ");
            Console.WriteLine(@"StickyNote Reader by @two06");
            Console.WriteLine("");
        }
    }
}
