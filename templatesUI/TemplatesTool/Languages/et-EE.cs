using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ETEE : LocalLanguage
    {
        public override string Tag { get; } = "et-EE";
        public override string Name { get; } = "Estonian";
        public override string Location { get; } = "Estonia";
        public override int Id { get; } = 0x0425;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDEstonian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdEstonian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-et-EE]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-425]";

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
            "dd\\.mm\\.yy;@",
            "dd\\.mm\\.yy;@",
            "dd\\.mm\\.yy;@",
            "[$-et-EE]d\\. mmmm yyyy\". a.\";@",
            "[$-et-EE]d\\. mmmm yyyy\". a.\";@",
            "[$-et-EE]dd\\. mmmm yyyy\". a.\";@",
            "[$-et-EE]d\\. mmmm yyyy\". a.\";@",
            "[$-et-EE]d\\. mmmm yyyy\". a.\";@",
            "[$-et-EE]d\\. mmmm yyyy\". a.\";@",
            "dd\\.mm\\.yy;@",
            "m/d/yyyy",
            "[$-et-EE]d\\. mmmm yyyy\". a.\";@",
            "m/d/yyyy",
            "m/d/yyyy",
            "[$-et-EE]dddd, d\\. mmmm yyyy;@"
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