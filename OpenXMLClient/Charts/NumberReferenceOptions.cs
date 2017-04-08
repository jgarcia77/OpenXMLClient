using System.Collections.Generic;

namespace OpenXMLClient.Charts
{
    public class NumberReferenceOptions
    {
        public NumberReferenceOptions()
        {
            this.Values = new List<string>();
        }

        public uint Id { get; set; }
        public string Letter { get; set; }
        public int RowStart { get; set; }
        public int RowEnd { get { return this.RowStart + this.Values.Count - 1; } }
        public List<string> Values { get; set; }
        public string FormatCode { get; set; }
    }
}
