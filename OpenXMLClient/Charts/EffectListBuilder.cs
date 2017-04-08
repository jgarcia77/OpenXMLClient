using DocumentFormat.OpenXml.Drawing;

namespace OpenXMLClient.Charts
{
    public class EffectListBuilder
    {
        public EffectListBuilder(EffectListOptions effectListOptions)
        {
            this.EffectListOptions = effectListOptions;
        }

        public EffectListOptions EffectListOptions { get; private set; }

        public EffectList Build()
        {
            var effectList = new EffectList();
            
            return effectList;
        }
    }
}
