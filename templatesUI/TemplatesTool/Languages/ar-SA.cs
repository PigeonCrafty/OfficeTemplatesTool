using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ARSA : LocalLanguage
    {
        public override string Tag { get; } = "ar-SA";
        public override string Name { get; } = "Arabic";
        public override string Location { get; } = "Saudi Arabia";
        public override int Id { get; } = 0x0401;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDArabic;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdArabic;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = true;
        public override bool IsSpecialFont { get; } = true;
        public override string SpecialFont { get; } = "Tahoma";
        public override string CurrencySymbolLocal1 { get; } = "ر.س.‏";
        public override string CurrencySymbolLocal2 { get; } = "[$ر.س.‏-ar-SA]";
        public override string CurrencySymbolLocal3 { get; } = "[$ر.س.‏-401]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"ر.س.‏\" #,##0.00_-",
            "\"ر.س.‏\" #,##0.00;[Red]\"ر.س.‏\" #,##0.00",
            "\"ر.س.‏\" #,##0.00_-;\"ر.س.‏\" #,##0.00-",
            "\"ر.س.‏\" #,##0.00_-;[Red]\"ر.س.‏\" #,##0.00-"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "dd/mm/yy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "[$-,101]d/m/yyyy;@",
            "[$-,101]yyyy/mm/dd;@",
            "[$-,201]d/mm/yyyy;@",
            "[$-,201]yyyy/mm/dd;@",
            "dd/mm/yy",
            "dd/mm/yy",
            "dd/mm/yy",
            "dd/mm/yy",
            "dd/mm/yy",
            "[$-ar-SA,101]d/m/yyyy h:mm AM/PM;@",
            "[$-ar-SA,201]d/mm/yyyy h:mm AM/PM;@",
            "dd/mm/yy",
            "dd/mm/yy",
            "dd/mm/yy",
            "[$-,101]d/m/yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"ر.س.‏\" * #,##0.00_-;_-\"ر.س.‏\" * #,##0.00-;_-\"ر.س.‏\" * \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}