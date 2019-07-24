using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class VIVN : LocalLanguage
    {
        public override string Tag { get; } = "vi-VN";
        public override string Name { get; } = "Vietnamese";
        public override string Location { get; } = "Vietnam";
        public override int Id { get; } = 0x042A;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDVietnamese;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdVietnamese;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "₫";
        public override string CurrencySymbolLocal2 { get; } = "[$₫-vi-VN]";
        public override string CurrencySymbolLocal3 { get; } = "[$₫-42A]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"₫\"",
            "#,##0.00 \"₫\";[Red]#,##0.00 \"₫\"",
            "#,##0.00 \"₫\";-#,##0.00 \"₫\"",
            "#,##0.00 \"₫\";[Red]-#,##0.00 \"₫\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "[$-,101]d/m/yy;@",
            "[$-,101]d/m/yy;@",
            "[$-,101]d/m/yy;@",
            "[$-vi-VN,101]d mmm yy;@",
            "[$-vi-VN,101]d mmm yy;@",
            "[$-vi-VN,101]d mmm yy;@",
            "[$-vi-VN,101]d mmm yy;@",
            "[$-vi-VN,101]d mmmm yyyy;@",
            "[$-vi-VN,101]d mmmm yyyy;@",
            "[$-vi-VN,101]d/m/yyyy h:mm AM/PM;@",
            "[$-vi-VN,101]d/m/yyyy h:mm AM/PM;@",
            "[$-,101]d/m/yyyy;@",
            "[$-,101]d/m/yyyy;@",
            "[$-,101]d/m/yyyy;@",
            "[$-vi-VN,101]d mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"₫\"_-;-* #,##0.00 \"₫\"_-;_-* \"-\"?? \"₫\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Arial Bold"))
                return "Segoe UI Bold";
            if (sourceFont.Contains("Arial Black") ||
                sourceFont.Contains("Impact"))
                return "Segoe UI Black";
            if (sourceFont.Contains("Cambria") ||
                sourceFont.Contains("Garamond") ||
                sourceFont.Contains("Georgia") ||
                sourceFont.Contains("Sylfaen") ||
                sourceFont.Contains("Palatino Linotype") ||
                sourceFont.Contains("Century School") ||
                sourceFont.Contains("Book Antiqua"))
                return "Times New Roman";
            if (sourceFont.Contains("Century Gothic") ||
                sourceFont.Contains("Euphemia") ||
                sourceFont.Contains("Freestyle Script") ||
                sourceFont.Contains("Rockwell") ||
                sourceFont.Contains("Gill Sans MT") ||
                sourceFont.Contains("French Script MT") ||
                sourceFont.Contains("Lucida Handwriting"))
                return "Calibri";
            if (sourceFont.Contains("Franklin Gothic"))
                return "Tahoma";
            if (sourceFont.Contains("Trebuchet MS") ||
                sourceFont.Contains("Bookman Old Style"))
                return "Arial";
            return sourceFont;
        }
    }
}