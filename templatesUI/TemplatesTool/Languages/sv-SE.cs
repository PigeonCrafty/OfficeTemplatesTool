using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class SVSE : LocalLanguage
    {
        public override string Tag { get; } = "sv-SE";
        public override string Name { get; } = "Swedish";
        public override string Location { get; } = "Sweden";
        public override int Id { get; } = 0x041D;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDSwedish;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdSwedish;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "kr";
        public override string CurrencySymbolLocal2 { get; } = "[$kr-sv-SE]";
        public override string CurrencySymbolLocal3 { get; } = "[$kr-41D]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "#,##0.00 \"kr\"",
            "#,##0.00 \"kr\";[Red]#,##0.00 \"kr\"",
            "#,##0.00 \"kr\";-#,##0.00 \"kr\"",
            "#,##0.00 \"kr\";[Red]-#,##0.00 \"kr\""
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "d-m-yy;@",
            "d-m-yy;@",
            "yy-mm-dd;@",
            "[$-sv-SE]dd-mmm;@",
            "[$-sv-SE]d mmmm -yy;@",
            "[$-sv-SE]d mmmm -yy;@",
            "[$-sv-SE]mmm-yy;@",
            "[$-sv-SE]mmmm -yy;@",
            "[$-sv-SE]d mmmm yyyy;@",
            "yy-mm-dd hh:mm;@",
            "yy-mm-dd hh:mm;@",
            "[$-sv-SE]mmmmm;@",
            "[$-sv-SE]mmmmm-yy;@",
            "d-m yyyy;@",
            "[$-sv-SE]d mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"kr\"_-;-* #,##0.00 \"kr\"_-;_-* \"-\"?? \"kr\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}