using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class SLSI : LocalLanguage
    {
        public override string Tag { get; } = "sl-SI";
        public override string Name { get; } = "Slovenian";
        public override string Location { get; } = "Slovenia";
        public override int Id { get; } = 0x0424;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDSlovenian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdSlovenian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "€";
        public override string CurrencySymbolLocal2 { get; } = "[$€-sl-SI]";
        public override string CurrencySymbolLocal3 { get; } = "[$€-424]";

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
            "d.m.yy;@",
            "d.m.yy;@",
            "dd.mm.yy;@",
            "[$-sl-SI]d. mmmm yyyy;@",
            "[$-sl-SI]d. mmmm yyyy;@",
            "[$-sl-SI]dd. mmmm yyyy;@",
            "[$-sl-SI]d. mmmm yyyy;@",
            "[$-sl-SI]d. mmmm yyyy;@",
            "[$-sl-SI]d. mmmm yyyy;@",
            "dd.mm.yy;@",
            "dd.mm.yy;@",
            "[$-sl-SI]d. mmmm yyyy;@",
            "[$-sl-SI]d. mmmm yyyy;@",
            "m/d/yyyy",
            "[$-sl-SI]d. mmmm yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-* #,##0.00 \"€\"_-;-* #,##0.00 \"€\"_-;_-* \"-\"?? \"€\"_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (sourceFont.Contains("French Script MT"))
                return "Monotype Corsiva";
            if (sourceFont.Contains("Euphemia"))
                return "Calibri";
            return sourceFont;
        }
    }
}