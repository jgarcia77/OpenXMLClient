using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.Generic;

namespace OpenXMLClient.Charts
{
    public class NumberReferenceBuilder
    {
        public NumberReferenceBuilder(NumberReferenceOptions options)
        {
            this.Options = options;
        }

        public NumberReferenceOptions Options { get; private set; }

        public void Add(string value)
        {
            this.Options.Values.Add(value);
        }

        public NumberReference Build()
        {
            var numberReference = new NumberReference();

            var formulaString = string.Concat(this.Options.Letter, "$", this.Options.RowStart, ":", this.Options.Letter, this.Options.RowEnd);

            numberReference.Formula = new Formula(formulaString);
            
            numberReference.NumberingCache = new NumberingCache();

            numberReference.NumberingCache.FormatCode = new FormatCode(this.Options.FormatCode);

            numberReference.NumberingCache.PointCount = new PointCount { Val = (uint)this.Options.Values.Count };

            for (var index = 0; index < this.Options.Values.Count; index++)
            {
                var numericPoint = new NumericPoint();

                numericPoint.Index = (uint)index;

                numericPoint.NumericValue = new NumericValue(this.Options.Values[index]);

                numberReference.NumberingCache.Append(numericPoint);
            }

            return numberReference;
        }
    }
}
