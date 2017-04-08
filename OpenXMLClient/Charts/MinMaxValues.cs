namespace OpenXMLClient.Charts
{
    public struct MinMaxValues
    {
        public MinMaxValues(double minValue, double maxValue)
        {
            this.MinValue = minValue;

            this.MaxValue = maxValue;
        }

        public double MinValue { get; set; }
        public double MaxValue { get; set; }
    }
}
