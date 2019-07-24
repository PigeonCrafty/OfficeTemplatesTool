using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class IDID : LocalLanguage
    {
        public override string Tag { get; } = "id-ID";
        public override string Name { get; } = "Indonesian";
        public override string Location { get; } = "Indonesia";
        public override int Id { get; } = 0x0421;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDIndonesian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdEnglishIndonesia;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "Rp";
        public override string CurrencySymbolLocal2 { get; } = "[$Rp-id-ID]";
        public override string CurrencySymbolLocal3 { get; } = "[$Rp-421]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"Rp\"#,##0.00",
            "\"Rp\"#,##0.00;[Red]\"Rp\"#,##0.00",
            "\"Rp\"#,##0.00;-\"Rp\"#,##0.00",
            "\"Rp\"#,##0.00;[Red]-\"Rp\"#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "dd/mm/yy;@",
            "dd/mm/yy;@",
            "dd/mm/yy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "dd/mm/yy;@",
            "dd/mm/yy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "[$-id-ID]dd mmmm yyyy;@",
            "dd/mm/yyyy",
            "[$-id-ID]dd mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"Rp\"* #,##0.00_-;-\"Rp\"* #,##0.00_-;_-\"Rp\"* \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            return sourceFont;
        }
    }
}