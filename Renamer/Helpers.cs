using System.Text.RegularExpressions;

namespace Renamer
{
    public class Helpers
    {
        public bool Comparator(string text, string text_search)
        {
            string regexPattern = Regex.Escape(text_search).Replace(@"\?", ".");
            return Regex.IsMatch(text, regexPattern, RegexOptions.IgnoreCase);
        }

        public string Replacer(string text, string text_search, string text_replace)
        {
            string regexPattern = Regex.Escape(text_search).Replace(@"\?", ".");
            return Regex.Replace(text, regexPattern, text_replace, RegexOptions.IgnoreCase);
        }

        public bool Contains(string text, string text_search)
        {
            return Regex.IsMatch(text, text_search, RegexOptions.IgnoreCase);
        }

        public string Replace(string text, string text_search, string text_replace)
        {
            return Regex.Replace(text, text_search, text_replace, RegexOptions.IgnoreCase);
        }

        /*string exePath = AppDomain.CurrentDomain.BaseDirectory;
        if (exePath.Contains(@"tnn\pir") || exePath.Contains("project"))
        {
            MessageBox.Show("Нельзя запускать программу на сетевом диске");
            Environment.Exit(1);
        }*/
    }
}
