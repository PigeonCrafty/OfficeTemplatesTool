using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;

namespace TemplatesTool.Models
{
    public partial class PowerPointHandler
    {
        #region Public Methods

        public void ProcText(POWERPOINT.Shape shape)
        {
            shape.TextFrame2.TextRange.LanguageID = LocalLang.pptID;

            var sourceFont = !string.IsNullOrEmpty(shape.TextEffect.FontName)
                ? shape.TextEffect.FontName
                : shape.TextFrame2.TextRange.Font.Name;
            var targetFont = LocalLang.GetLocFont(sourceFont);

            //_SetFontProperty(targetFont, shape.TextEffect.FontName, shape.TextFrame2.TextRange.Font.Name,
            //    shape.TextFrame2.TextRange.Font.NameAscii);
            shape.TextEffect.FontName = targetFont;
            shape.TextFrame2.TextRange.Font.Name = targetFont;
            shape.TextFrame2.TextRange.Font.NameAscii = targetFont;

            if (LocalLang.IsFarEast) shape.TextFrame2.TextRange.Font.NameFarEast = targetFont;
            if (LocalLang.IsRightToLeft)
            {
                shape.TextFrame2.TextRange.Font.NameComplexScript = targetFont;
                shape.TextFrame2.TextRange.Font.NameOther = targetFont;
            }
        }

        public void ProcTable(POWERPOINT.Shape shape)
        {
            for (var i = 1; i <= shape.Table.Rows.Count; i++)
            for (var j = 1; j <= shape.Table.Columns.Count; j++)
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    shape.Table.Cell(i,j).Shape.TextFrame2.TextRange.LanguageID = LocalLang.pptID;

                    var sourceFont = shape.Table.Cell(i, j).Shape.TextFrame2.TextRange.Font.Name;
                    var targetFont = LocalLang.GetLocFont(sourceFont);

                    //_SetFontProperty(targetFont, shape.Table.Cell(i, j).Shape.TextFrame2.TextRange.Font.Name,
                    //    shape.Table.Rows[i].Cells[j].Shape.TextFrame2.TextRange.Font.NameAscii);
                    shape.Table.Cell(i, j).Shape.TextFrame2.TextRange.Font.Name = targetFont;
                    shape.Table.Rows[i].Cells[j].Shape.TextFrame2.TextRange.Font.NameAscii = targetFont;

                    if (LocalLang.IsFarEast)
                        shape.Table.Cell(i, j).Shape.TextFrame2.TextRange.Font.NameFarEast = targetFont;
                    if (LocalLang.IsRightToLeft)
                    {
                        shape.Table.Cell(i, j).Shape.TextFrame2.TextRange.Font.NameComplexScript = targetFont;
                        shape.Table.Cell(i, j).Shape.TextFrame2.TextRange.Font.NameOther = targetFont;
                    }
                }
        }

        public void ProcChart(POWERPOINT.Shape shape)
        {
            if (shape.Chart.HasTitle)
            {
                string sourceFont = shape.Chart.ChartTitle.Font.Name;
                var targetFont = LocalLang.GetLocFont(sourceFont);

                shape.Chart.ChartTitle.Font.Name = targetFont;
            }

            if (shape.Chart.HasLegend)
            {
                string sourceFont = shape.Chart.ChartArea.Font.Name;
                var targetFont = LocalLang.GetLocFont(sourceFont);

                shape.Chart.ChartArea.Font.Name = targetFont;
            }

            if (shape.Chart.HasLegend)
            {
                string sourceFont = shape.Chart.ChartArea.Font.Name;
                var targetFont = LocalLang.GetLocFont(sourceFont);

                shape.Chart.ChartArea.Font.Name = targetFont;
            }

            if (shape.Chart.HasDataTable)
            {
                string sourceFont = shape.Chart.DataTable.Font.Name;
                var targetFont = LocalLang.GetLocFont(sourceFont);

                shape.Chart.DataTable.Font.Name = targetFont;
            }
        }

        public void ProcSmartArt(POWERPOINT.Shape shape)
        {
            for (var i = 1; i <= shape.SmartArt.Nodes.Count; i++)
                if (shape.GroupItems[i].HasTextFrame == MsoTriState.msoTrue)
                {
                    shape.GroupItems[i].TextFrame2.TextRange.LanguageID = LocalLang.pptID;

                    var sourceFont = shape.GroupItems[i].TextFrame2.TextRange.Font.Name;
                    var targetFont = LocalLang.GetLocFont(sourceFont);

                    //_SetFontProperty(targetFont, shape.GroupItems[i].TextFrame2.TextRange.Font.Name,
                    //    shape.GroupItems[i].TextFrame2.TextRange.Font.NameAscii);
                    shape.GroupItems[i].TextFrame2.TextRange.Font.Name = targetFont;
                    shape.GroupItems[i].TextFrame2.TextRange.Font.NameAscii = targetFont;

                    if (LocalLang.IsFarEast) shape.GroupItems[i].TextFrame2.TextRange.Font.NameFarEast = targetFont;

                    if (LocalLang.IsRightToLeft)
                    {
                        //_SetFontProperty(targetFont, shape.GroupItems[i].TextFrame2.TextRange.Font.NameComplexScript,
                        //    shape.GroupItems[i].TextFrame2.TextRange.Font.NameOther);
                        shape.GroupItems[i].TextFrame2.TextRange.Font.NameComplexScript = targetFont;
                        shape.GroupItems[i].TextFrame2.TextRange.Font.NameOther = targetFont;
                    }
                }
        }

        public void ProcGroups(POWERPOINT.Shape shape)
        {
            for (var i = 1; i <= shape.GroupItems.Count; i++)
                if (shape.GroupItems[i].HasTextFrame == MsoTriState.msoTrue)
                    if (shape.GroupItems[i].HasTextFrame == MsoTriState.msoTrue)
                    {
                        shape.GroupItems[i].TextFrame2.TextRange.LanguageID = LocalLang.pptID;

                        var sourceFont = shape.GroupItems[i].TextFrame2.TextRange.Font.Name;
                        var targetFont = LocalLang.GetLocFont(sourceFont);

                        //_SetFontProperty(targetFont, shape.TextEffect.FontName,
                        //    shape.GroupItems[i].TextFrame2.TextRange.Font.Name,
                        //    shape.GroupItems[i].TextFrame2.TextRange.Font.NameAscii);
                        shape.TextEffect.FontName = targetFont;
                        shape.GroupItems[i].TextFrame2.TextRange.Font.Name = targetFont;
                        shape.GroupItems[i].TextFrame2.TextRange.Font.NameAscii = targetFont;

                        if (LocalLang.IsFarEast) shape.GroupItems[i].TextFrame2.TextRange.Font.NameFarEast = targetFont;

                        if (LocalLang.IsRightToLeft)
                        {
                            shape.GroupItems[i].TextFrame2.TextRange.Font.NameComplexScript = targetFont;
                            shape.GroupItems[i].TextFrame2.TextRange.Font.NameOther = targetFont;

                            //_SetFontProperty(targetFont,
                            //    shape.GroupItems[i].TextFrame2.TextRange.Font.NameComplexScript,
                            //    shape.GroupItems[i].TextFrame2.TextRange.Font.NameOther);
                            shape.GroupItems[i].TextFrame2.TextRange.Font.NameComplexScript = targetFont;
                            shape.GroupItems[i].TextFrame2.TextRange.Font.NameOther = targetFont;
                        }
                    }
        }

        public void ProcPlaceHolder(POWERPOINT.Shape shape)
        {
            Common.WriteLine("Cannot update the Placeholder of " + shape.PlaceholderFormat.Type);
            // TODO:
            //if (shape.PlaceholderFormat.Type == POWERPOINT.PpPlaceholderType.ppPlaceholderPicture)
        }

        public void Dispose()
        {
            if (pre != null)
            {
                pre.Save();
                pre.Close();
                Marshal.ReleaseComObject(pre);
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