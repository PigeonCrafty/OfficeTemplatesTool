using Microsoft.Office.Core;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Languages
{
    public sealed class ZHCN : LocalLanguage
    {
        public override string Tag { get; } = "zh-CN";
        public override string Name { get; } = "Chinese (Simplified)";
        public override string Location { get; } = "People's Republic of China";
        public override int Id { get; } = 0x0804;
        public override MsoLanguageID pptID { get; } = MsoLanguageID.msoLanguageIDSimplifiedChinese;
        public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdSimplifiedChinese;

        public override bool IsFarEast { get; } = true;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = true;
        public override string SpecialFont { get; } = "Microsoft YaHei UI";
        public override string CurrencySymbolLocal1 { get; } = "¥";
        public override string CurrencySymbolLocal2 { get; } = "[$¥-zh-CN]";
        public override string CurrencySymbolLocal3 { get; } = "[$¥-804]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"¥\"#,##0.00;\"¥\"-#,##0.00",
            "\"¥\"#,##0.00;[Red]\"¥\"#,##0.00",
            "\"¥\"#,##0.00_);(\"¥\"#,##0.00)",
            "\"¥\"#,##0.00_);[Red](\"¥\"#,##0.00)"
            //"\"¥\"#,##0.00;[Red]\"¥\"-#,##0.00" // Additional 1: Index[4]
        };

        public override string[] DateFormatLocal { get; } =
        {
            "yyyy/m/d",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "m\"月\"d\"日\";@",
            "m/d/yy;@",
            "mm/dd/yy;@",
            "[DBNum1][$-zh-CN]m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\";@",
            "yyyy\"年\"m\"月\";@",
            "[DBNum1][$-zh-CN]yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy/m/d h:mm;@",
            "yyyy/m/d h:mm;@",
            "[DBNum1][$-zh-CN]yyyy\"年\"m\"月\";@",
            "[DBNum1][$-zh-CN]yyyy\"年\"m\"月\";@",
            "yyyy/m/d;@",
            "yyyy\"年\"m\"月\"d\"日\";@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_ \"¥\"* #,##0.00_ ;_ \"¥\"* -#,##0.00_ ;_ \"¥\"* \"-\"??_ ;_ @_ ";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}