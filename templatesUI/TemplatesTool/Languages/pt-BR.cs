using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class PTBR : LocalLanguage
    {
        public override string Tag { get; } = "pt-BR";
        public override string Name { get; } = "Portuguese";
        public override string Location { get; } = "Brazil";
        public override int Id { get; } = 0x0416;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDBrazilianPortuguese;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdPortugueseBrazil;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "R$";
        public override string CurrencySymbolLocal2 { get; } = "[$R$-pt-BR]";
        public override string CurrencySymbolLocal3 { get; } = "[$R$-416]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"R$\" #,##0.00",
            "\"R$\" #,##0.00;[Red]\"R$\" #,##0.00",
            "\"R$\" #,##0.00;-\"R$\" #,##0.00",
            "\"R$\" #,##0.00;[Red]-\"R$\" #,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d/m;@",
            "d/m/yy;@",
            "dd/mm/yy;@",
            "[$-pt-BR]d-mmm;@",
            "[$-pt-BR]d-mmm-yy;@",
            "[$-pt-BR]dd-mmm-yy;@",
            "[$-pt-BR]mmm-yy;@",
            "[$-pt-BR]mmmm-yy;@",
            "[$-pt-BR]d-mmm-yy;@",
            "d/m/yy h:mm;@",
            "d/m/yy h:mm;@",
            "[$-pt-BR]mmm-yy;@",
            "[$-pt-BR]mmm-yy;@",
            "dd/mm/yyyy",
            "[$-pt-BR]d-mmm-yy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"R$\" * #,##0.00_-;-\"R$\" * #,##0.00_-;_-\"R$\" * \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}