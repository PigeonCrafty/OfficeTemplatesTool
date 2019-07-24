using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class FIFI : LocalLanguage
    {
        public override string Tag { get; } = "fi-FI";
        public override string Name { get; } = "Finnish";
        public override string Location { get; } = "Finland";
        public override int Id { get; } = 0x040B;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDFinnish;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdFinnish;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-fi-FI]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-40B]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"€\"",
            "#,##0.00 \"€\";[Red]#,##0.00 \"€\"",
            "#,##0.00 \"€\";-#,##0.00 \"€\"",
            "#,##0.00 \"€\";[Red]-#,##0.00 \"€\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d\\.m\\.;@",
            "d\\.m\\.yy;@",
            "d\\.m\\.yy;@",
            "[$-fi-FI]d. mmmmt\\a;@",
            "[$-fi-FI]d. mmmmt\\a yy;@",
            "[$-fi-FI]d. mmmmt\\a yy;@",
            "[$-fi-FI]mmmm yy;@",
            "[$-fi-FI]mmmm yy;@",
            "[$-fi-FI]d. mmmmt\\a yyyy;@",
            "d\\.m\\.yyyy h:mm;@",
            "d\\.m\\.yy h:mm;@",
            "[$-fi-FI]mmmmm;@",
            "[$-fi-FI]mmmmm yy;@",
            "yyyy-mm-dd;@",
            "[$-fi-FI]d. mmmmt\\a yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"€\"_-;-* #,##0.00 \"€\"_-;_-* \"-\"?? \"€\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}