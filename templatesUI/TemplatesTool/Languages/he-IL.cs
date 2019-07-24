using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class HEIL : LocalLanguage
    {
        public override string Tag { get; } = "he-IL";
        public override string Name { get; } = "Hebrew";
        public override string Location { get; } = "Israel";
        public override int Id { get; } = 0x040D;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDHebrew;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdHebrew;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = true;
        public override bool IsSpecialFont { get; } = true;
        public override string SpecialFont { get; } = "Tahoma";
        public override string CurrencySymbolLocal1 { get; } = "₪";
        public override string CurrencySymbolLocal2 { get; } = "[$₪-he-IL]";
        public override string CurrencySymbolLocal3 { get; } = "[$₪-40D]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"₪\" #,##0.00",
            "\"₪\" #,##0.00;[Red]\"₪\" #,##0.00",
            "\"₪\" #,##0.00;\"₪\" -#,##0.00",
            "\"₪\" #,##0.00;[Red]\"₪\" -#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "[$-,101]d/m/yy;@",
            "[$-,101]d/m/yy;@",
            "[$-,101]d/m/yy;@",
            "[$-he-IL,101]d mmm yy;@",
            "[$-he-IL,101]d mmm yy;@",
            "[$-he-IL,101]d mmm yy;@",
            "[$-he-IL,101]d mmm yy;@",
            "[$-he-IL,101]d mmm yy;@",
            "[$-he-IL,101]d mmm yyyy;@",
            "[$-en-US,101]d/m/yyyy h:mm;@",
            "[$-en-US,101]d/m/yyyy h:mm;@",
            "[$-he-IL,101]d mmm yy;@",
            "[$-he-IL,101]d mmm yy;@",
            "[$-,101]d/m/yyyy;@",
            "[$-he-IL,101]d mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_ \"₪\" * #,##0.00_ ;_ \"₪\" * -#,##0.00_ ;_ \"₪\" * \"-\"??_ ;_ @_ ";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}