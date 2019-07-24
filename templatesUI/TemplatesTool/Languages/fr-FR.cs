using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class FRFR : LocalLanguage
    {
        public override string Tag { get; } = "fr-FR";
        public override string Name { get; } = "French";
        public override string Location { get; } = "France";
        public override int Id { get; } = 0x040C;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDFrench;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdFrench;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-fr-FR]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-40C]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"€\"",
            "#,##0.00 \"€\";[Red]#,##0.00 \"€\"",
            "#,##0.00 \"€\";-#,##0.00 \"€\"",
            "#,##0.00 \"€\";[Red]-#,##0.00 \"€\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d/m;@",
            "d/m/yy;@",
            "dd/mm/yy;@",
            "[$-fr-FR]d-mmm;@",
            "[$-fr-FR]d-mmm-yy;@",
            "[$-fr-FR]dd-mmm-yy;@",
            "[$-fr-FR]mmm-yy;@",
            "[$-fr-FR]mmmm-yy;@",
            "[$-fr-FR]d mmmm yyyy;@",
            "d/m/yy h:mm;@",
            "d/m/yy h:mm;@",
            "[$-fr-FR]mmmmm;@",
            "[$-fr-FR]mmmmm-yy;@",
            "m/d/yyyy;@",
            "[$-fr-FR]d-mmm-yyyy;@"
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