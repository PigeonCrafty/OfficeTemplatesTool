using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ESES : LocalLanguage
    {
        public override string Tag { get; } = "es-ES";
        public override string Name { get; } = "Spanish";
        public override string Location { get; } = "Spain";
        public override int Id { get; } = 0x0c0A;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDSpanish;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdSpanish;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-es-ES]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-408]";

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
            "d-m;@",
            "d-m-yy;@",
            "dd-mm-yy;@",
            "[$-es-ES]d-mmm;@",
            "[$-es-ES]d-mmm-yy;@",
            "[$-es-ES]dd-mmm-yy;@",
            "[$-es-ES]mmm-yy;@",
            "[$-es-ES]mmmm-yy;@",
            "[$-es-ES]d \"de\" mmmm \"de\" yyyy;@",
            "d-m-yy h:mm;@",
            "d-m-yy h:mm;@",
            "[$-es-ES]mmmmm;@",
            "[$-es-ES]mmmmm-yy;@",
            "d-m-yyyy;@",
            "[$-es-ES]d-mmm-yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"€\"_-;-* #,##0.00 \"€\"_-;_-* \"-\"?? \"€\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}