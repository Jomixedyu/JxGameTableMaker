using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CSVReaderP
{
    public class CSVReader
    {
        public static DataTable CSV2DataTable(string data)
        {
            Encode(data);
            return null;
            DataTable dt = new DataTable();
            string[] contexts = data.Split('\r');
            if (contexts.Length < 2) return null;
            string[] fields = contexts[0].Split('\r');

            return null;
        }

        private static string Encode(string str)
        {
            string st = str.Replace("\"\"", "%Sym-34%");
            Regex reg = new Regex("\".*\"");
            foreach (var item in reg.Matches(st))
            {
                Console.WriteLine(item.ToString());
            } 
            return st;
        }
        private static string DeCode(string str)
        {
            string st = str.Replace("%Sym-34%", "\"\"");
            return st;
        }
    }
}
