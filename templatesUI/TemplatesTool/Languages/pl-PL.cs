using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class PLPL : LocalLanguage
    {
        public override string Tag { get; } = "pl-PL";
        public override string Name { get; } = "Polish";
        public override string Location { get; } = "Poland";
        public override int Id { get; } = 0x0415;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDPolish;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdPolish;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "zł";
        public override string CurrencySymbolLocal2 { get; } = "[$zł-pl-PL]";
        public override string CurrencySymbolLocal3 { get; } = "[$zł-415]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"zł\"",
            "#,##0.00 \"zł\";[Red]#,##0.00 \"zł\"",
            "#,##0.00 \"zł\";-#,##0.00 \"zł\"",
            "#,##0.00 \"zł\";[Red]-#,##0.00 \"zł\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d-mm;@",
            "yy-mm-dd;@",
            "yy-mm-dd;@",
            "[$-pl-PL]d mmm;@",
            "[$-pl-PL]d mmm yy;@",
            "[$-pl-PL]dd mmm yy;@",
            "[$-pl-PL]mmm yy;@",
            "[$-pl-PL]mmmm yy;@",
            "[$-pl-PL]d mmmm yyyy;@",
            "dd-mm-yy h:mm;@",
            "dd-mm-yy h:mm;@",
            "[$-pl-PL]mmmmm;@",
            "[$-pl-PL]mmmmm\\.yy;@",
            "d-m-yyyy;@",
            "[$-pl-PL]d-mmm-yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"zł\"_-;-* #,##0.00 \"zł\"_-;_-* \"-\"?? \"zł\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Euphemia"))
                return "Calibri";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            return sourceFont;
        }
    }
}