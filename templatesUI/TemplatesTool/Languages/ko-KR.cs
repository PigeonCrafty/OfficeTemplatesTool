using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class KOKR : LocalLanguage
    {
        public override string Tag { get; } = "ko-KR";
        public override string Name { get; } = "Korean";
        public override string Location { get; } = "Korea";
        public override int Id { get; } = 0x0412;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDKorean;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdKorean;

        public override bool IsFarEast { get; } = true;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = true;
        public override string SpecialFont { get; } = "Malgun Gothic";
        public override string CurrencySymbolLocal1 { get; } = "₩";
        public override string CurrencySymbolLocal2 { get; } = "[$₩-ko-KR]";
        public override string CurrencySymbolLocal3 { get; } = "[$₩-412]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"₩\"#,##0.00",
            "\"₩\"#,##0.00;[Red]\"₩\"#,##0.00",
            "\"₩\"#,##0.00;-\"₩\"#,##0.00",
            "\"₩\"#,##0.00;[Red]-\"₩\"#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "m\"/\"d;@",
            "yy\"-\"m\"-\"d;@",
            "yy\"-\"m\"-\"d;@",
            "m\"월\" d\"일\";@",
            "yyyy\"년\" m\"월\" d\"일\";@",
            "yyyy\"년\" m\"월\" d\"일\";@",
            "m\"월\" d\"일\";@",
            "m\"월\" d\"일\";@",
            "yyyy\"년\" m\"월\" d\"일\";@",
            "[$-ko-KR]yy\"-\"m\"-\"d AM/PM h:mm;@",
            "yy\"-\"m\"-\"d h:mm;@",
            "yyyy\"년\" m\"월\";@",
            "yyyy\"년\" m\"월\";@",
            "yyyy\"-\"m\"-\"d;@",
            "yyyy\"년\" m\"월\" d\"일\";@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"₩\"* #,##0.00_-;-\"₩\"* #,##0.00_-;_-\"₩\"* \"-\"??_-;_-@_-";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}