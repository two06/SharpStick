using Microsoft.Data.Sqlite;
using Microsoft.Win32;
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
            var path = @"C:\Users\two06\AppData\Roaming\Microsoft\StickyNotes\StickyNotes.snt";
            var notes = RunQueryLegacy(path);
            return;

            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var dbRelativePath = @"\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite";
            if(!File.Exists(appData + dbRelativePath))
            {
                Console.WriteLine("[*] StickNotes SQLite DB not found!");
            }
            try
            {
                var results = RunQuery(appData + dbRelativePath);
                foreach (var result in results)
                {
                    Console.WriteLine(result);
                }
            }
            catch (Exception)
            {
                Console.WriteLine("[*] Exception occured reading StickyNotes DB!");
            }
            
            Console.ReadLine();
        }

        private static List<string> RunQuery(string dbPath)
        {
            var list = new List<string>();
            using (var connection = new SqliteConnection("Data Source=" + dbPath))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText =
                @"
                    SELECT text
                    FROM note
                ";
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var text = reader.GetString(0);
                        list.Add(text);
                    }
                }
            }
            return list;
        }

        private static List<string> RunQueryLegacy(string path)
        {
            var sStorage = new StructuredStorage();
            var notes = sStorage.readFile(path);
            return notes;
        }
        private static void PrintHeader()
        {

        }
    }
}
