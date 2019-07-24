using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Models
{
    public partial class WordHandler
    {
        #region Public Methods

        public void ProcPageSize()
        {
            if (doc.PageSetup.Orientation == WORD.WdOrientation.wdOrientPortrait)
                doc.PageSetup.PaperSize = WORD.WdPaperSize.wdPaperA4;
            else
                Common.WriteLine("<!> Please change Page Size manually!");
        }

        public void ProcStyles(WORD.Style style)
        {
            var targetFont = LocalLang.GetLocFont(style.Font.Name);
            
            //_SetFontProperty(targetFont, style.Font.Name, style.Font.NameAscii, style.Font.NameOther);

            style.Font.Name = targetFont;
            style.Font.NameAscii = targetFont;
            style.Font.NameOther = targetFont;

            if (LocalLang.IsFarEast) style.Font.NameFarEast = targetFont;

            if (LocalLang.IsRightToLeft) style.Font.NameBi = targetFont;

        }

        public void ProcContentControls(WORD.ContentControl contentControl)
        {
            contentControl.Range.LanguageID = LocalLang.wdID;

            var targetFont = LocalLang.GetLocFont(contentControl.Range.Font.Name);
            
            //_SetFontProperty(targetFont, contentControl.Range.Font.Name, contentControl.Range.Font.NameAscii,
            //contentControl.Range.Font.NameOther);

            contentControl.Range.Font.Name = targetFont;
            contentControl.Range.Font.NameAscii = targetFont;
            contentControl.Range.Font.NameOther = targetFont;

            if (LocalLang.IsFarEast)
            {
                contentControl.Range.LanguageIDFarEast = LocalLang.wdID;
                contentControl.Range.Font.NameFarEast = targetFont;
            }

            if (LocalLang.IsRightToLeft)
            {
                contentControl.Range.LanguageIDOther = LocalLang.wdID;
                contentControl.Range.Font.NameBi = targetFont;
            }
        }

        public void ProcHeadersFooters(WORD.HeaderFooter header, WORD.HeaderFooter footer)
        {
            header.Range.LanguageID = LocalLang.wdID;
            footer.Range.LanguageID = LocalLang.wdID;

            var targetFontHeader = LocalLang.GetLocFont(header.Range.Font.Name);

            //_SetFontProperty(targetFontHeader, header.Range.Font.Name, header.Range.Font.NameAscii,
            //    header.Range.Font.NameOther);
            header.Range.Font.Name = targetFontHeader;
            header.Range.Font.NameAscii = targetFontHeader;
            header.Range.Font.NameOther = targetFontHeader;

            var targetFontFooter = LocalLang.GetLocFont(footer.Range.Font.Name);

            //_SetFontProperty(targetFontFooter, footer.Range.Font.Name, footer.Range.Font.NameAscii,
            //    footer.Range.Font.NameOther);
            footer.Range.Font.Name = targetFontFooter;
            footer.Range.Font.NameAscii = targetFontFooter;
            footer.Range.Font.NameOther = targetFontFooter;

            if (LocalLang.IsFarEast)
            {
                header.Range.LanguageIDFarEast = LocalLang.wdID;
                header.Range.Font.NameFarEast = targetFontHeader;

                footer.Range.LanguageIDFarEast = LocalLang.wdID;
                footer.Range.Font.NameFarEast = targetFontFooter;
            }

            if (LocalLang.IsRightToLeft)
            {
                header.Range.LanguageIDOther = LocalLang.wdID;
                header.Range.Font.NameBi = targetFontHeader;

                footer.Range.LanguageIDOther = LocalLang.wdID;
                footer.Range.Font.NameBi = targetFontFooter;
            }
        }

        public void ProcShapes(WORD.Shape shape)
        {
            shape.TextFrame.TextRange.LanguageID = LocalLang.wdID;

            var sourceFont = !string.IsNullOrEmpty(shape.TextEffect.Text)
                ? shape.TextEffect.FontName
                : shape.TextFrame2.TextRange.Font.Name;
            var targetFont = LocalLang.GetLocFont(sourceFont);

            //_SetFontProperty(targetFont, shape.TextEffect.FontName, shape.TextFrame.TextRange.Font.Name,
            //    shape.TextFrame.TextRange.Font.NameAscii, shape.TextFrame.TextRange.Font.NameOther);
            shape.TextEffect.FontName = targetFont;
            shape.TextFrame.TextRange.Font.NameAscii = targetFont;
            shape.TextFrame.TextRange.Font.NameOther = targetFont;

            if (LocalLang.IsFarEast)
            {
                shape.TextFrame.TextRange.LanguageIDFarEast = LocalLang.wdID;
                shape.TextFrame.TextRange.Font.NameFarEast = targetFont;
            }

            if (LocalLang.IsRightToLeft)
            {
                shape.TextFrame.TextRange.LanguageIDOther = LocalLang.wdID;
                shape.TextFrame.TextRange.Font.NameBi = targetFont;
            }
        }

        public void ProcTables(WORD.Cell cell)
        {
            cell.Range.LanguageID = LocalLang.wdID;

            var targetFont = LocalLang.GetLocFont(cell.Range.Font.Name);

            //_SetFontProperty(targetFont, cell.Range.Font.Name, cell.Range.Font.NameAscii,
            //    cell.Range.Font.NameOther);
            cell.Range.Font.Name = targetFont;
            cell.Range.Font.NameAscii = targetFont;
            cell.Range.Font.NameOther = targetFont;

            if (LocalLang.IsFarEast)
            {
                cell.Range.LanguageIDFarEast = LocalLang.wdID;
                cell.Range.Font.NameFarEast = targetFont;
            }

            if (LocalLang.IsRightToLeft)
            {
                cell.Range.LanguageIDOther = LocalLang.wdID;
                cell.Range.Font.NameBi = targetFont;
            }
        }

        public void ProcTableContent(WORD.Paragraph paragraph)
        {
            paragraph.Range.LanguageID = LocalLang.wdID;

            var targetFont = LocalLang.GetLocFont(paragraph.Range.Font.Name);

            //_SetFontProperty(targetFont, paragraph.Range.Font.Name, paragraph.Range.Font.NameAscii,
            //    paragraph.Range.Font.NameOther);
            paragraph.Range.Font.Name = targetFont;
            paragraph.Range.Font.NameAscii = targetFont;
            paragraph.Range.Font.NameOther = targetFont;

            if (LocalLang.IsFarEast)
            {
                paragraph.Range.LanguageIDFarEast = LocalLang.wdID;
                paragraph.Range.Font.NameFarEast = targetFont;
            }

            if (LocalLang.IsRightToLeft)
            {
                paragraph.Range.LanguageIDOther = LocalLang.wdID;
                paragraph.Range.Font.NameBi = targetFont;
            }
        }

        public void ProcWholeDoc()
        {
            doc.Content.LanguageID = LocalLang.wdID;

            var targetFont = LocalLang.GetLocFont(doc.Content.Font.Name);

            //_SetFontProperty(targetFont, doc.Content.Font.Name, doc.Content.Font.NameAscii,
            //    doc.Content.Font.NameOther);
            doc.Content.Font.Name = targetFont;
            doc.Content.Font.NameAscii = targetFont;
            doc.Content.Font.NameOther = targetFont;

            if (LocalLang.IsFarEast)
            {
                doc.Content.LanguageIDFarEast = LocalLang.wdID;
                doc.Content.Font.NameFarEast = targetFont;
            }

            if (LocalLang.IsRightToLeft)
            {
                doc.Content.LanguageIDOther = LocalLang.wdID;
                doc.Content.Font.NameBi = targetFont;
            }
        }

        public void Dispose()
        {
            if (doc != null)
            {
                doc.Save();
                doc.Close();
                Marshal.ReleaseComObject(doc);
                GC.WaitForPendingFinalizers();
            }

            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }

            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        #endregion

        #region Private Methods

        private static void _SetFontProperty(string targetFont, params object[] propertiesObjects)
        {
            for (var i = 0; i < propertiesObjects.Length; i++)
            {
                propertiesObjects[i] = targetFont;
            }
        }

        #endregion
    }
}