using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class CSCZ : LocalLanguage
    {
        public override string Tag { get; } = "cs-CZ";
        public override string Name { get; } = "Czech";
        public override string Location { get; } = "Czech Republic";
        public override int Id { get; } = 0x0405;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDCzech;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdCzech;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "Kč‏";
        public override string CurrencySymbolLocal2 { get; } = "[$Kč-cs-CZ]";
        public override string CurrencySymbolLocal3 { get; } = "[$Kč-405]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"Kč\" #,##0.00",
            "\"Kč\" #,##0.00;[Red]\"Kč\" #,##0.00",
            "\"Kč\" #,##0.00;\"Kč\" -#,##0.00",
            "\"Kč\" #,##0.00;[Red]\"Kč\" -#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d.m;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-cs-CZ]d-mmm\\.;@",
            "[$-cs-CZ]d/mmm/yy;@",
            "[$-cs-CZ]dd-mmm-yy;@",
            "[$-cs-CZ]mmm-yy;@",
            "[$-cs-CZ]mmmm yy;@",
            "[$-cs-CZ]d\\. mmmm yyyy;@",
            "d.m.yy h:mm;@",
            "d.m.yy h:mm;@",
            "[$-cs-CZ]mmmmm;@",
            "[$-cs-CZ]mmmmm-yy;@",
            "d.m.yyyy;@",
            "[$-cs-CZ]d-mmm-yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"Kč\"_-;-* #,##0.00 \"Kč\"_-;_-* \"-\"?? \"Kč\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Euphemia") || sourceFont.Contains("Gill Sans MT"))
                return "Calibri";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            return sourceFont;
        }
    }
}