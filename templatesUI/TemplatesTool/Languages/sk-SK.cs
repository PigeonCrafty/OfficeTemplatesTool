using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class SKSK : LocalLanguage
    {
        public override string Tag { get; } = "sk-SK";
        public override string Name { get; } = "Slovak";
        public override string Location { get; } = "Slovakia";
        public override int Id { get; } = 0x041B;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDSlovak;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdSlovak;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "EUR";
        public override string CurrencySymbolLocal2 { get; } = "[$EUR]";
        public override string CurrencySymbolLocal3 { get; } = "[$EUR-41B]";

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
            "[$-sk-SK]d-mmm\\.;@",
            "[$-sk-SK]d.mmm.yy;@",
            "[$-sk-SK]dd-mmm-yy;@",
            "[$-sk-SK]mmm-yy;@",
            "[$-sk-SK]mmmm yy;@",
            "[$-sk-SK]d\\. mmmm yyyy;@",
            "d/m/yy h:mm;@",
            "d/m/yy h:mm;@",
            "[$-sk-SK]mmmmm;@",
            "[$-sk-SK]mmmmm-yy;@",
            "d/m/yyyy;@",
            "[$-sk-SK]d/mmm/yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"€\"_-;-* #,##0.00 \"€\"_-;_-* \"-\"?? \"€\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Gill Sans MT"))
                return "Calibri";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            return sourceFont;
        }
    }
}