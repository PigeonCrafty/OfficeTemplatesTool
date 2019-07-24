using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class DEDE : LocalLanguage
    {
        public override string Tag { get; } = "de-DE";
        public override string Name { get; } = "German";
        public override string Location { get; } = "Germany";
        public override int Id { get; } = 0x0407;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDGerman;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdGerman;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-de-DE]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-407]";

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
            "d.m;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-de-DE]d/ mmm/;@",
            "[$-de-DE]d/ mmm/ yy;@",
            "[$-de-DE]d/ mmm yy;@",
            "[$-de-DE]mmm/ yy;@",
            "[$-de-DE]mmmm yy;@",
            "[$-de-DE]d/ mmmm yyyy;@",
            "d/m/yy h:mm;@",
            "d.m.yy h:mm;@",
            "[$-de-DE]mmmmm;@",
            "[$-de-DE]mmmmm yy;@",
            "d.m.yyyy;@",
            "[$-de-DE]d. mmm. yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"€\"_-;-* #,##0.00 \"€\"_-;_-* \"-\"?? \"€\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            return sourceFont;
        }
    }
}