using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class THTH : LocalLanguage
    {
        public override string Tag { get; } = "th-TH";
        public override string Name { get; } = "Thai";
        public override string Location { get; } = "Thailand";
        public override int Id { get; } = 0x041E;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDThai;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdThai;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = true; // Special
        public override bool IsSpecialFont { get; } = true;
        public override string SpecialFont { get; } = "Leelawadee";
        public override string CurrencySymbolLocal1 { get; } = "฿";
        public override string CurrencySymbolLocal2 { get; } = "[$฿-th-TH]";
        public override string CurrencySymbolLocal3 { get; } = "[$฿-41E]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"฿\"#,##0.00",
            "\"฿\"#,##0.00;[Red]\"฿\"#,##0.00",
            "\"฿\"#,##0.00;-\"฿\"#,##0.00",
            "\"฿\"#,##0.00;[Red]-\"฿\"#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "[$-,107]d/mm/yyyy;@",
            "[$-th-TH,107]d mmmm yyyy;@",
            "[$-,107]d/m/yy;@",
            "[$-,107]d/m/yy;@",
            "[$-,D07]d/m/yy;@",
            "[$-th-TH,D07]d mmm yy;@",
            "[$-th-TH,D07]d mmm yy;@",
            "[$-th-TH,D07]d mmm yy;@",
            "[$-th-TH,D07]d mmm yy;@",
            "[$-th-TH,107]d mmm yy;@",
            "[$-th-TH,107]d mmmm yyyy;@",
            "[$-,D07]d/mm/yyyy h:mm \"น.\";@",
            "[$-,107]d/mm/yyyy h:mm \"น.\";@",
            "[$-,107]d/m/yy;@",
            "[$-,107]d/m/yy;@",
            "[$-,107]d/mm/yyyy;@",
            "[$-th-TH,107]d mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"฿\"* #,##0.00_-;-\"฿\"* #,##0.00_-;_-\"฿\"* \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}