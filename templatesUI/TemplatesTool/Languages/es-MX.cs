using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ESMX : LocalLanguage
    {
        public override string Tag { get; } = "es-MX";
        public override string Name { get; } = "Spanish";
        public override string Location { get; } = "Mexico";
        public override int Id { get; } = 0x080A;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDMexicanSpanish;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdMexicanSpanish;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "$";
        public override string CurrencySymbolLocal2 { get; } = "[$$-es-MX]";
        public override string CurrencySymbolLocal3 { get; } = "[$$-80A]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"$\"#,##0.00",
            "\"$\"#,##0.00;[Red]\"$\"#,##0.00",
            "\"$\"#,##0.00;-\"$\"#,##0.00",
            "\"$\"#,##0.00;[Red]-\"$\"#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d/m/yy;@",
            "d/m/yy;@",
            "d/mm/yy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "d/m/yy;@",
            "d/m/yy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@",
            "dd/mm/yyyy;@",
            "[$-es-MX]d\" de \"mmmm\" de \"yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"$\"* #,##0.00_-;-\"$\"* #,##0.00_-;_-\"$\"* \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}