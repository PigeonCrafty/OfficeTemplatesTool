using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class LTLT : LocalLanguage
    {
        public override string Tag { get; } = "lt-LT";
        public override string Name { get; } = "Lithuanian";
        public override string Location { get; } = "Lithuania";
        public override int Id { get; } = 0x0427;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDLithuanian;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdLithuanian;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "EUR";
        public override string CurrencySymbolLocal2 { get; } = "[$EUR]";
        public override string CurrencySymbolLocal3 { get; } = "[$EUR-427]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "[$EUR] #,##0.00",
            "[$EUR] #,##0.00;[Red][$EUR] #,##0.00",
            "[$EUR] #,##0.00_);([$EUR] #,##0.00)",
            "[$EUR] #,##0.00_);[Red]([$EUR] #,##0.00)"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "yyyy-mm-dd;@",
            "yyyy.mm.dd;@",
            "yyyy-mm-dd;@",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "m/d/yyyy",
            "m/d/yyyy",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@",
            "m/d/yyyy",
            "[$-lt-LT]yyyy \"m.\" mmmm d \"d.\";@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_([$EUR] * #,##0.00_);_([$EUR] * (#,##0.00);_([$EUR] * \"-\"??_);_(@_)";
        }

        public override string GetLocFont(string sourceFont)
        {
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