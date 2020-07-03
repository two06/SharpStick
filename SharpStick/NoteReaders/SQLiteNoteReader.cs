using Microsoft.Data.Sqlite;
using SharpStick.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharpStick.NoteReaders
{
    class SQLiteNoteReader : IStickyNoteReader
    {
        public IEnumerable<string> GetNotes(string path)
        {
            var list = new List<string>();
            using (var connection = new SqliteConnection("Data Source=" + path))
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
                        //the notes have and ID string followed by a space, we want everything after that space
                        //we could do this with Regex, but Regex is aweful
                        list.Add(text.Substring(text.IndexOf(' ')));
                    }
                }
            }
            return list;
        }
    }
}
