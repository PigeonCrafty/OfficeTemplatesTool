using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class UKUA : LocalLanguage
    {
        public override string Tag { get; } = "uk-UA";
        public override string Name { get; } = "Ukrainian";
        public override string Location { get; } = "Ukraine";
        public override int Id { get; } = 0x0422;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDUkrainian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdUkrainian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "₴";
        public override string CurrencySymbolLocal2 { get; } = "[$₴-uk-UA]";
        public override string CurrencySymbolLocal3 { get; } = "[$₴-422]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"₴\"",
            "#,##0.00 \"₴\";[Red]#,##0.00 \"₴\"",
            "#,##0.00 \"₴\";-#,##0.00 \"₴\"",
            "#,##0.00 \"₴\";[Red]-#,##0.00 \"₴\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "dd\\.mm\\.yy;@",
            "dd\\.mm\\.yy;@",
            "dd\\.mm\\.yy;@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "dd\\.mm\\.yy;@",
            "dd\\.mm\\.yy;@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@",
            "dd\\.mm\\.yyyy;@",
            "[$-uk-UA-x-genlower]d mmmm yyyy\" р.\";@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"₴\"_-;-* #,##0.00 \"₴\"_-;_-* \"-\"?? \"₴\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            if (sourceFont.Contains("Gill Sans MT"))
                return "Calibri";
            return sourceFont;
        }
    }
}