using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class SRRS : LocalLanguage
    {
        public override string Tag { get; } = "sr-Latn-RS";
        public override string Name { get; } = "Serbian (Latin)";
        public override string Location { get; } = "Serbia";
        public override int Id { get; } = 0x241A;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDSerbianLatin;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdSerbianLatin;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "RSD";
        public override string CurrencySymbolLocal2 { get; } = "[$RSD]";
        public override string CurrencySymbolLocal3 { get; } = "[$RSD-241A]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"RSD\"",
            "#,##0.00 \"RSD\";[Red]#,##0.00 \"RSD\"",
            "#,##0.00 \"RSD\";-#,##0.00 \"RSD\"",
            "#,##0.00 \"RSD\";[Red]-#,##0.00 \"RSD\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d.m.yy;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-sr-Latn-RS]d. mmmm yyyy;@",
            "[$-sr-Latn-RS]d. mmmm yyyy;@",
            "[$-sr-Latn-RS]dd. mmmm yyyy;@",
            "[$-sr-Latn-RS]d. mmmm yyyy;@",
            "[$-sr-Latn-RS]d. mmmm yyyy;@",
            "[$-sr-Latn-RS]d. mmmm yyyy;@",
            "d.m.yy;@",
            "d.m.yy;@",
            "[$-sr-Latn-RS]d. mmmm yyyy;@",
            "[$-sr-Latn-RS]d. mmmm yyyy;@",
            "m/d/yyyy",
            "[$-sr-Latn-RS]d/ mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"RSD\"_-;-* #,##0.00 \"RSD\"_-;_-* \"-\"?? \"RSD\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Euphemia"))
                return "Calibri";
            if (sourceFont.Contains("Freestyle Script"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            return sourceFont;
        }
    }
}