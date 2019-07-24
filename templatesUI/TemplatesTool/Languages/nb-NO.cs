using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class NBNO : LocalLanguage
    {
        public override string Tag { get; } = "nb-NO";
        public override string Name { get; } = "Norwegian (Bokmal)";
        public override string Location { get; } = "Norway";
        public override int Id { get; } = 0x0414;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDNorwegianBokmol;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdNorwegianBokmol;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "kr";
        public override string CurrencySymbolLocal2 { get; } = "[$kr-nb-NO]";
        public override string CurrencySymbolLocal3 { get; } = "[$kr-414]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"kr\" #,##0.00",
            "\"kr\" #,##0.00;[Red]\"kr\" #,##0.00",
            "\"kr\" #,##0.00;-\"kr\" #,##0.00",
            "\"kr\" #,##0.00;[Red]-\"kr\" #,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d.m.;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-nb-NO]d/ mmm.;@",
            "[$-nb-NO]d/ mmm. yyyy;@",
            "[$-nb-NO]d/ mmm. yyyy;@",
            "[$-nb-NO]mmm\\. yy;@",
            "[$-nb-NO]mmm\\. yy;@",
            "[$-nb-NO]d/ mmmm yyyy;@",
            "dd/mm/yy h:mm;@",
            "dd/mm/yy h:mm;@",
            "[$-nb-NO]mmm\\. yy;@",
            "[$-nb-NO]mmm\\. yy;@",
            "d.m.yyyy;@",
            "[$-nb-NO]d/ mmm. yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"kr\" * #,##0.00_-;-\"kr\" * #,##0.00_-;_-\"kr\" * \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}