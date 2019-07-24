using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class RORO : LocalLanguage
    {
        public override string Tag { get; } = "ro-RO";
        public override string Name { get; } = "Romanian";
        public override string Location { get; } = "Romania";
        public override int Id { get; } = 0x0418;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDRomanian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdRomanian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "lei";
        public override string CurrencySymbolLocal2 { get; } = "[$lei-ro-RO]";
        public override string CurrencySymbolLocal3 { get; } = "[$lei-418]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"lei\"",
            "#,##0.00 \"lei\";[Red]#,##0.00 \"lei\"",
            "#,##0.00 \"lei\";-#,##0.00 \"lei\"",
            "#,##0.00 \"lei\";[Red]-#,##0.00 \"lei\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d.m;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-ro-RO]d-mmm;@",
            "[$-ro-RO]d-mmm-yy;@",
            "[$-ro-RO]dd-mmm-yy;@",
            "[$-ro-RO]mmm-yy;@",
            "[$-ro-RO]mmmm-yy;@",
            "[$-ro-RO]d mmmm yyyy;@",
            "d.m.yy h:mm;@",
            "d.m.yy h:mm;@",
            "[$-ro-RO]mmmmm;@",
            "[$-ro-RO]mmmmm-yy;@",
            "d.m.yyyy;@",
            "[$-ro-RO]d-mmm-yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"lei\"_-;-* #,##0.00 \"lei\"_-;_-* \"-\"?? \"lei\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Century Gothic") || sourceFont.Contains("Franklin Gothic") ||
                sourceFont.Contains("Freestyle Script") || sourceFont.Contains("Gill Sans MT"))
                return "Calibri";
            if (sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Garamond") || sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            return sourceFont;
        }
    }
}