using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ENGB : LocalLanguage
    {
        public override string Tag { get; } = "en-GB";
        public override string Name { get; } = "English";
        public override string Location { get; } = "United Kingdom";
        public override int Id { get; } = 0x0809;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDEnglishUK;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdEnglishUK;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "£";
        public override string CurrencySymbolLocal2 { get; } = "[$£-en-GB]";
        public override string CurrencySymbolLocal3 { get; } = "[$£-809]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"£\"#,##0.00",
            "\"£\"#,##0.00;[Red]\"£\"#,##0.00",
            "\"£\"#,##0.00;-\"£\"#,##0.00",
            "\"£\"#,##0.00;[Red]-\"£\"#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d/m/yy;@",
            "d/m/yy;@",
            "dd/mm/yy;@",
            "[$-en-GB]d mmmm yyyy;@",
            "[$-en-GB]d mmmm yyyy;@",
            "[$-en-GB]dd mmmm yyyy;@",
            "[$-en-GB]d mmmm yyyy;@",
            "[$-en-GB]d mmmm yyyy;@",
            "[$-en-GB]d mmmm yyyy;@",
            "d/m/yy;@",
            "d/m/yy;@",
            "[$-en-GB]d mmmm yyyy;@",
            "[$-en-GB]d mmmm yyyy;@",
            "dd/mm/yyyy",
            "[$-en-GB]d mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"£\"* #,##0.00_-;-\"£\"* #,##0.00_-;_-\"£\"* \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}