using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class HRHR : LocalLanguage
    {
        public override string Tag { get; } = "hr-HR";
        public override string Name { get; } = "Croatian";
        public override string Location { get; } = "Croatia";
        public override int Id { get; } = 0x041A;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDCroatian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdCroatian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "kn";
        public override string CurrencySymbolLocal2 { get; } = "[$kn-hr-HR]";
        public override string CurrencySymbolLocal3 { get; } = "[$kn-41A]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"kn\"",
            "#,##0.00 \"kn\";[Red]#,##0.00 \"kn\"",
            "#,##0.00 \"kn\";-#,##0.00 \"kn\"",
            "#,##0.00 \"kn\";[Red]-#,##0.00 \"kn\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d.m.;@",
            "d.m.yy.;@",
            "dd.mm.yy.;@",
            "[$-hr-HR]d-mmm;@",
            "[$-hr-HR]d-mmm-yy;@",
            "[$-hr-HR]dd-mmm-yy;@",
            "[$-hr-HR]mmm-yy;@",
            "[$-hr-HR]mmmm-yy;@",
            "[$-hr-HR]d. mmmm yyyy.;@",
            "d.m.yy. h:mm;@",
            "d.m.yy. h:mm;@",
            "[$-hr-HR]mmmmm;@",
            "[$-hr-HR]mmmmm-yy.;@",
            "d.m.yyyy.;@",
            "[$-hr-HR]d-mmm-yyyy.;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"kn\"_-;-* #,##0.00 \"kn\"_-;_-* \"-\"?? \"kn\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Rockwell"))
                return "Times override Roman";
            return sourceFont;
        }
    }
}