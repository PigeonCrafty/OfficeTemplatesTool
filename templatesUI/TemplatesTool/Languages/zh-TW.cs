using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ZHTW : LocalLanguage
    {
        public override string Tag { get; } = "zh-TW";
        public override string Name { get; } = "Chinese (Traditional)";
        public override string Location { get; } = "Taiwan";
        public override int Id { get; } = 0x0404;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDTraditionalChinese;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdTraditionalChinese;

        public override bool IsFarEast { get; } = true;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = true;
        public override string SpecialFont { get; } = "Microsoft JhengHei UI";
        public override string CurrencySymbolLocal1 { get; } = "NT$";
        public override string CurrencySymbolLocal2 { get; } = "[$NT$-zh-TW]";
        public override string CurrencySymbolLocal3 { get; } = "[$NT$-404]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"NT$\"#,##0.00",
            "\"NT$\"#,##0.00;[Red]\"NT$\"#,##0.00",
            "\"NT$\"#,##0.00_);(\"NT$\"#,##0.00)",
            "\"NT$\"#,##0.00_);[Red](\"NT$\"#,##0.00)",
            "\"NT$\"#,##0.00;[Red]-\"NT$\"#,##0.00"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "yyyy/m/d",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "m/d;@",
            "m/d/yy;@",
            "mm/dd/yy;@",
            "[DBNum1][$-zh-TW]m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy/m/d h:mm;@",
            "yyyy/m/d h:mm;@",
            "[DBNum1][$-zh-TW]m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy/m/d",
            "yyyy\"年\"m\"月\"d\"日\";@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_-\"NT$\"* #,##0.00_ ;_-\"NT$\"* -#,##0.00 ;_-\"NT$\"* \"-\"??_ ;_-@_ ";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}