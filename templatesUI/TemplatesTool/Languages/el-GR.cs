using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ELGR : LocalLanguage
    {
        public override string Tag { get; } = "el-GR";
        public override string Name { get; } = "Greek";
        public override string Location { get; } = "Greece";
        public override int Id { get; } = 0x0408;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDGreek;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdGreek;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-el-GR]";
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
            "d/m;@",
            "d/m/yy;@",
            "dd/mm/yy;@",
            "[$-el-GR]d-mmm;@",
            "[$-el-GR]d-mmm-yy;@",
            "[$-el-GR]d-mmm-yy;@",
            "[$-el-GR]mmm-yy;@",
            "[$-el-GR]mmm-yy;@",
            "[$-el-GR]d mmmm yyyy;@",
            "[$-el-GR]d/m/yy h:mm AM/PM;@",
            "d/m/yy h:mm;@",
            "[$-el-GR]mmmmm;@",
            "[$-el-GR]mmmmm-yy;@",
            "[$-el-GR]d-mmm-yyyy;@",
            "[$-el-GR]d-mmm-yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"€\"_-;-* #,##0.00 \"€\"_-;_-* \"-\"?? \"€\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("Century Gothic"))
                return "Tahoma";
            if (sourceFont.Contains("Freestyle Script") || sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Euphemia") || sourceFont.Contains("Gill Sans MT"))
                return "Calibri";
            if (sourceFont.Contains("Rockwell"))
                return "Times New Roman";
            return sourceFont;
        }
    }
}