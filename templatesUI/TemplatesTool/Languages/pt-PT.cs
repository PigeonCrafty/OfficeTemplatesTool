using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class PTPT : LocalLanguage
    {
        public override string Tag { get; } = "pt-PT";
        public override string Name { get; } = "Portuguese";
        public override string Location { get; } = "Portugal";
        public override int Id { get; } = 0x0816;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDPortuguese;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdPortuguese;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-pt-PT]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-816]";

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
            "d/m/yy;@",
            "d/m/yy;@",
            "dd/mm/yy;@",
            "[$-pt-PT]d/mmm;@",
            "[$-pt-PT]d-mmm-yy;@",
            "[$-pt-PT]dd-mmm-yy;@",
            "[$-pt-PT]mmm/yy;@",
            "[$-pt-PT]mmmm yy;@",
            "[$-pt-PT]d \"de\" mmmm \"de\" yyyy;@",
            "d/m/yy h:mm;@",
            "d/m/yy h:mm;@",
            "[$-pt-PT]mmmmm;@",
            "[$-pt-PT]mmmmm-yy;@",
            "d/m/yyyy;@",
            "[$-pt-PT]d-mmm-yyyy;@"
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