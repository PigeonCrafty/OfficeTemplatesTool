using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class LVLV : LocalLanguage
    {
        public override string Tag { get; } = "lv-LV";
        public override string Name { get; } = "Latvian";
        public override string Location { get; } = "Latvia";
        public override int Id { get; } = 0x0426;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDLatvian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdLatvian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "EUR";
        public override string CurrencySymbolLocal2 { get; } = "[$EUR]";
        public override string CurrencySymbolLocal3 { get; } = "[$EUR-427]";

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
            "yy\\.mm\\.dd\\.;@\"",
            "yy\\.mm\\.dd\\.;@",
            "yy\\.mm\\.dd\\.;@\"",
            "[$-lv-LV]dddd, yyyy\". gada \"d\\. mmmm;@",
            "m/d/yyyy",
            "m/d/yyyy",
            "m/d/yyyy",
            "m/d/yyyy",
            "m/d/yyyy",
            "yy\\.mm\\.dd\\.;@\"",
            "yy\\.mm\\.dd\\.;@\"",
            "m/d/yyyy",
            "m/d/yyyy",
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"€\"_-;-* #,##0.00 \"€\"_-;_-* \"-\"?? \"€\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Euphemia"))
                return "Calibri";
            return sourceFont;
        }
    }
}