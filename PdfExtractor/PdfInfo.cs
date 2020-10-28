using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace PdfExtractor
{
    class PdfInfo
    {
        string pp;
        public PdfInfo(string path)
        {
            pp = path;
        }

        public int ExtractInfo()
        {
            using (StreamReader sr = new StreamReader(File.OpenRead(pp)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
            }
        }
    }
}
