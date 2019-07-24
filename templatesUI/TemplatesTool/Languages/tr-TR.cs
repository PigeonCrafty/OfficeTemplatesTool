using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class TRTR : LocalLanguage
    {
        public override string Tag { get; } = "tr-TR";
        public override string Name { get; } = "Turkish";
        public override string Location { get; } = "Turkey";
        public override int Id { get; } = 0x041F;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDTurkish;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdTurkish;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "₺";
        public override string CurrencySymbolLocal2 { get; } = "[$₺-tr-TR]";
        public override string CurrencySymbolLocal3 { get; } = "[$₺-41F]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"₺\"",
            "#,##0.00 \"₺\";[Red]#,##0.00 \"₺\"",
            "#,##0.00 \"₺\";-#,##0.00 \"₺\"",
            "#,##0.00 \"₺\";[Red]-#,##0.00 \"₺\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d.m;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-tr-TR]d mmmm;@",
            "[$-tr-TR]d mmmm yy;@",
            "[$-tr-TR]dd mmmm yy;@",
            "[$-tr-TR]mmmm yy;@",
            "[$-tr-TR]mmmm yy;@",
            "[$-tr-TR]d mmmm yyyy;@",
            "d.m.yy h:mm;@",
            "d.m.yy h:mm;@",
            "[$-tr-TR]mmmmm;@",
            "[$-tr-TR]mmmmm yy;@",
            "m.d.yyyy;@",
            "[$-tr-TR]d mmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"₺\"_-;-* #,##0.00 \"₺\"_-;_-* \"-\"?? \"₺\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Euphemia"))
                return "Calibri";
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            if (sourceFont.Contains("Gill Sans MT"))
                return "Arial";
            return sourceFont;
        }
    }
}