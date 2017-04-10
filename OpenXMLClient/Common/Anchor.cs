namespace OpenXMLClient.Common
{
    using System.Collections.Generic;
    using System.Text.RegularExpressions;

    public class Anchor
    {
        public string Value;
        public string Href;
        public string Text;

        public static List<Anchor> Find(string value)
        {
            List<Anchor> returnList = new List<Anchor>();

            MatchCollection m1 = Regex.Matches(value, @"(<a.*?>.*?</a>)", RegexOptions.Singleline);
            
            foreach (Match m in m1)
            {
                string matchValue = m.Groups[1].Value;

                Anchor i = new Anchor { Value = matchValue };
                                
                Match m2 = Regex.Match(matchValue, @"href=\""(.*?)\""", RegexOptions.Singleline);

                if (m2.Success)
                {
                    i.Href = m2.Groups[1].Value;
                }
                
                string t = Regex.Replace(matchValue, @"\s*<.*?>\s*", "", RegexOptions.Singleline);

                i.Text = t;

                returnList.Add(i);
            }

            return returnList;
        }
    }
}
