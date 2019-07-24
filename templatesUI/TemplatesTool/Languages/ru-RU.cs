using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class RURU : LocalLanguage
    {
        public override string Tag { get; } = "ru-RU";
        public override string Name { get; } = "Russian";
        public override string Location { get; } = "Russia";
        public override int Id { get; } = 0x0419;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDRussian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdRussian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "₽";
        public override string CurrencySymbolLocal2 { get; } = "[$₽-ru-RU]";
        public override string CurrencySymbolLocal3 { get; } = "[$₽-419]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"₽\"",
            "#,##0.00 \"₽\";[Red]#,##0.00 \"₽\"",
            "#,##0.00 \"₽\";-#,##0.00 \"₽\"",
            "#,##0.00 \"₽\";[Red]-#,##0.00 \"₽\""
        };

        public override string[] DateFormatLocal { get; } = //TODO: Russian's VMware is not available
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d.m;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-ru-RU]d mmm;@",
            "[$-ru-RU]d mmm yy;@",
            "[$-ru-RU]dd mmm yy;@",
            "[$-ru-RU]mmmm yyyy;@",
            "[$-ru-RU]mmmm yyyy;@",
            "[$-ru-RU-x-genlower]dd mmmm yyyy г.;@",
            "dd/mm/yy h:mm;@",
            "dd/mm/yy h:mm;@",
            "[$-ru-RU]mmmm;@",
            "[$-ru-RU]mmmm yyyy;@",
            "d.m.yyyy;@",
            "[$-ru-RU]d-mmm-yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"₽\"_-;-* #,##0.00 \"₽\"_-;_-* \"-\"?? \"₽\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Gill Sans MT"))
                return "Calibri";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            return sourceFont;
        }
    }
}