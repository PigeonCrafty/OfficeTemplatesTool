using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class DADK : LocalLanguage
    {
        public override string Tag { get; } = "da-DK";
        public override string Name { get; } = "Danish";
        public override string Location { get; } = "Denmark";
        public override int Id { get; } = 0x0406;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDDanish;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdDanish;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "kr.";
        public override string CurrencySymbolLocal2 { get; } = "[$kr.-da-DK]";
        public override string CurrencySymbolLocal3 { get; } = "[$kr.-406]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"Kč\"",
            "#,##0.00 \"Kč\";[Red]#,##0.00 \"Kč\"",
            "#,##0.00 \"Kč\";-#,##0.00 \"Kč\"",
            "#,##0.00 \"Kč\";[Red]-#,##0.00 \"Kč\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "dd-mm-yy;@",
            "dd-mm-yy;@",
            "dd-mm-yy;@",
            "yyyy.mm.dd;@",
            "yy-mm-dd;@",
            "dd-mm-yy;@",
            "[$-da-DK]mmmm yy;@",
            "d-m yyyy;@",
            "[$-da-DK]d. mmmm yyyy;@",
            "dd-mm-yy hh:mm:ss;@",
            "dd-mm-yy hh:mm:ss;@",
            "[$-da-DK]mmmm yy;@",
            "[$-da-DK]mmmm yy;@",
            "dd-mm-yyyy",
            "[$-da-DK]d. mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"kr.\"_-;-* #,##0.00 \"kr.\"_-;_-* \"-\"?? \"kr.\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}