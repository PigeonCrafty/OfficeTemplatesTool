using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class NLNL : LocalLanguage
    {
        public override string Tag { get; } = "nl-NL";
        public override string Name { get; } = "Dutch";
        public override string Location { get; } = "Netherlands";
        public override int Id { get; } = 0x0413;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDFrisianNetherlands;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdFrisianNetherlands;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-nl-NL]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-413]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"kr\"#,##0.00",
            "\"kr\"#,##0.00;[Red]\"kr\"#,##0.00",
            "\"kr\" #,##0.00;-\"kr\" #,##0.00",
            "\"kr\" #,##0.00;[Red]-\"kr\" #,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "d-m-yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d-m;@",
            "d-mm-yy;@",
            "dd-mm-yy;@",
            "[$-nl-NL]d-mmm;@",
            "[$-nl-NL]d-mmm-yy;@",
            "[$-nl-NL]dd-mmm-yy;@",
            "[$-nl-NL]mmm-yy;@",
            "[$-nl-NL]mmmm-yy;@",
            "[$-nl-NL]d mmmm yyyy;@",
            "d-mm-yy h:mm;@",
            "d-mm-yy h:mm;@",
            "[$-nl-NL]mmmmm;@",
            "[$-nl-NL]mmmmm-yy;@",
            "m-d-yyyy;@",
            "[$-nl-NL]d-mmm-yyyy;@"
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