using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ITIT : LocalLanguage
    {
        public override string Tag { get; } = "it-IT";
        public override string Name { get; } = "Italian";
        public override string Location { get; } = "Italy";
        public override int Id { get; } = 0x0410;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDItalian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdItalian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-it-IT]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-410]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"€\" #,##0.00",
            "\"€\" #,##0.00;[Red]#,##0.00",
            "\"€\" #,##0.00;-\"€\" #,##0.00",
            "\"€\" #,##0.00;[Red]-\"€\" #,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d/m;@",
            "d/m/yy;@",
            "dd/mm/yy;@",
            "[$-it-IT]d-mmm;@",
            "[$-it-IT]d-mmm-yy;@",
            "[$-it-IT]dd-mmm-yy;@",
            "[$-it-IT]mmm-yy;@",
            "[$-it-IT]mmmm-yy;@",
            "[$-it-IT]d mmmm yyyy;@",
            "d/m/yy h.mm;@",
            "d/m/yy h.mm;@",
            "[$-it-IT]mmmmm;@",
            "[$-it-IT]mmmmm-yy;@",
            "d/m/yyyy;@",
            "[$-it-IT]d-mmm-yyyy;@"
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