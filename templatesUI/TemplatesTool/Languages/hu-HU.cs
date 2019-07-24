using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class HUHU : LocalLanguage
    {
        public override string Tag { get; } = "hu-HU";
        public override string Name { get; } = "Hungary";
        public override string Location { get; } = "Hungarian";
        public override int Id { get; } = 0x040E;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDHungarian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdHungarian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "Ft";
        public override string CurrencySymbolLocal2 { get; } = "[$Ft-hu-HU]";
        public override string CurrencySymbolLocal3 { get; } = "[$Ft-40E]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"Ft\"",
            "#,##0.00 \"Ft\";[Red]#,##0.00 \"Ft\"",
            "#,##0.00 \"Ft\";-#,##0.00 \"Ft\"",
            "#,##0.00 \"Ft\";[Red]-#,##0.00 \"Ft\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "m. d.;@",
            "yyyy. m. d.;@",
            "yyyy.mm.dd;@",
            "[$-hu-HU]yy/ mmmm;@",
            "[$-hu-HU]yy/ mmmm d.;@",
            "[$-hu-HU]yy/ mmmm d.;@",
            "[$-hu-HU]yy/ mmmm;@",
            "[$-hu-HU]yy/ mmmm;@",
            "[$-hu-HU]yyyy/ mmmm d.;@",
            "[$-hu-HU]yyyy/ m/ d. h:mm AM/PM;@",
            "yyyy/ m/ d. h:mm;@",
            "[$-hu-HU]mmmmm\\.;@",
            "[$-hu-HU]mmm/ d.;@",
            "m/d/yyyy",
            "[$-hu-HU]yyyy. mmm. d\\.;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"Ft\"_-;-* #,##0.00 \"Ft\"_-;_-* \"-\"?? \"Ft\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            return sourceFont;
        }
    }
}