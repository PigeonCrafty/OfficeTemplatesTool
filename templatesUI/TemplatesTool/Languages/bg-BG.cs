using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class BGBG : LocalLanguage
    {
        public override string Tag { get; } = "bg-BG";
        public override string Name { get; } = "Bulgarian";
        public override string Location { get; } = "Bulgaria";
        public override int Id { get; } = 0x0402;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDBulgarian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdBulgarian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "лв.‏";
        public override string CurrencySymbolLocal2 { get; } = "[$лв.-bg-BG]";
        public override string CurrencySymbolLocal3 { get; } = "[$лв.-402]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"лв.\"",
            "#,##0.00 \"лв.\";[Red]#,##0.00 \"лв.\"",
            "#,##0.00 \"лв.\";-#,##0.00 \"лв.\"",
            "#,##0.00 \"лв.\";[Red]-#,##0.00 \"лв.\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "yyyy-mm-dd;@",
            "yyyy-mm-dd;@",
            "yyyy-mm-dd;@",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "m/d/yyyy",
            "m/d/yyyy",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@",
            "m/d/yyyy",
            "[$-bg-BG]dd mmmm yyyy \"г.\";@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"лв.\"_-;-* #,##0.00 \"лв.\"_-;_-* \"-\"?? \"лв.\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Rockwell"))
                return "Times override Roman";
            return sourceFont;
        }
    }
}