using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class HIIN : LocalLanguage
    {
        public override string Tag { get; } = "hi-IN";
        public override string Name { get; } = "Hindi";
        public override string Location { get; } = "India";
        public override int Id { get; } = 0x0439;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDHindi;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdHindi;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "₹";
        public override string CurrencySymbolLocal2 { get; } = "[$₹-hi-IN]";
        public override string CurrencySymbolLocal3 { get; } = "[$₹-439]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            //TODO
        };

        public override string[] DateFormatLocal { get; } =
        {
            //TODO
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_ \"₹\" * #,##0.00_ ;_ \"₹\" * -#,##0.00_ ;_ \"₹\" * \"-\"??_ ;_ @_ ";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}