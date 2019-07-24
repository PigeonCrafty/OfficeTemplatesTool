using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace TemplatesTool.Models
{
    public partial class ExcelHandler
    {
        #region Public Methods

        public void UnprotectedAndUnhide()
        {
            if (wb == null) throw new Exception("No active workbook.");

            foreach (EXCEL.Worksheet sheet in wb.Worksheets)
                if (sheet.Type != EXCEL.XlSheetType.xlWorksheet)
                {
                    Common.WriteLine("[Manual Needed Sheet] " + sheet.Name);
                }
                else
                {
                    if (sheet.Visible != EXCEL.XlSheetVisibility.xlSheetVisible)
                        try
                        {
                            _listHiddenWorkbook.Add(sheet);
                            sheet.Visible = EXCEL.XlSheetVisibility.xlSheetVisible;
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("[Failed to Unhide sheet] " + sheet.Name);
                            continue;
                        }

                    if (sheet.ProtectScenarios)
                        try
                        {
                            _listProtectedWorkbook.Add(sheet);
                            sheet.Unprotect();
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("[Failed to Unprotect sheet] " + sheet.Name);
                        }
                }

            Common.WriteLine("<Hidden Sheet>: " + _listHiddenWorkbook.Count);
            Common.WriteLine("<Protected Sheet>: " + _listProtectedWorkbook.Count);
        }

        public void GetSheetStatusBack()
        {
            if (_listHiddenWorkbook.Count > 0)
                foreach (var sheet in _listHiddenWorkbook)
                    try
                    {
                        sheet.Visible = EXCEL.XlSheetVisibility.xlSheetHidden;
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("[Failed to hide sheet: ] " + sheet.Name);
                    }

            if (_listProtectedWorkbook.Count > 0)
                foreach (var sheet in _listProtectedWorkbook)
                    try
                    {
                        sheet.Protect();
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("[Failed to protect sheet: ] " + sheet.Name);
                    }
        }

        public void ProcStyles(EXCEL.Style style)
        {
            try
            {
                style.Font.Name = LocalLang.GetLocFont(style.Font.Name);
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed] Style Font: " + style.Name + " " + style.Font.Name);
            }


            if (style.NumberFormat.Contains("$") && style.NumberFormat.Contains("#"))
            {
                var numberFormat = style.NumberFormat;
                var styleName = style.Name;
                try
                {
                    style.NumberFormat = ProcCurrency(numberFormat, styleName);
                }
                catch (Exception)
                {
                    Common.WriteLine("[Failed] Style Currency: " + style.Name + " " + style.NumberFormat);
                }
            }
        }

        public void SetPaperSize(EXCEL.Worksheet worksheet)
        {
            try
            {
                worksheet.PageSetup.PaperSize = EXCEL.XlPaperSize.xlPaperA4;
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to Resize]" + worksheet.Name);
            }
        }

        public void ProcTextShape(EXCEL.Shape shape)
        {
            try
            {
                if (!string.IsNullOrEmpty(shape.TextEffect.Text) ||
                    !string.IsNullOrEmpty(shape.TextFrame2.TextRange.Text))
                {
                    var sourceFont = !string.IsNullOrEmpty(shape.TextEffect.Text)
                        ? shape.TextEffect.FontName
                        : shape.TextFrame2.TextRange.Font.Name;

                    _ChangeFont(sourceFont, shape);
                }
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to Update Textbox] " + shape.Name);
            }
        }

        public void ProcTextGroup(EXCEL.GroupShapes group)
        {
            try
            {
                foreach (EXCEL.Shape shape in group)
                    if (shape.Type == MsoShapeType.msoAutoShape || shape.Type == MsoShapeType.msoTextBox)
                        ProcTextShape(shape);
            }
            catch (Exception)
            {
                Common.WriteLine("[Fail to Update Group]");
            }
        }

        public void ProcChart(EXCEL.Shape shape)
        {
            try
            {
                string chartTitle = shape.Chart.ChartTitle.Font.Name;
                shape.Chart.ChartTitle.Font.Name = LocalLang.GetLocFont(chartTitle);
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to update Chart Title]" + shape.Chart.Name);
            }

            try
            {
                if (shape.Chart.HasLegend)
                {
                    string chartLegend = shape.Chart.Legend.Font.Name;
                    shape.Chart.Legend.Font.Name = LocalLang.GetLocFont(chartLegend);
                }
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to update Chart Legend]" + shape.Chart.Name);
            }

            try
            {
                if (shape.Chart.HasDataTable)
                {
                    string chartDataTable = shape.Chart.DataTable.Font.Name;
                    shape.Chart.Legend.Font.Name = LocalLang.GetLocFont(chartDataTable);
                }
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to update Chart Datatable]" + shape.Chart.Name);
            }
        }

        public void UpdateCellFontFromStyle(EXCEL.Range cell)
        {
            if (!LocalLang.IsSpecialFont) return;
            if (cell.Font.Name != cell.Style.Font.Name) cell.Font.Name = cell.Style.Font.Name;
        }

        public string ProcCurrency(string numberFormat, string msg)
        {
            try
            {
                var digitNum = GetDigitNum(numberFormat);
                bool isAccount;

                numberFormat = numberFormat.Contains("*")
                    ? LocalLang.GetLocAccounting(numberFormat)
                    : LocalLang.GetLocCurrency(numberFormat);

                if (numberFormat.Contains("*"))
                {
                    numberFormat = LocalLang.GetLocAccounting(numberFormat);
                    isAccount = true;
                }
                else
                {
                    numberFormat = LocalLang.GetLocCurrency(numberFormat);
                    isAccount = false;
                }

                numberFormat = ProcDigitNum(numberFormat, msg, digitNum, isAccount);
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed] Number Format: " + msg);
            }

            return numberFormat;
        }

        public string ProcDigitNum(string numFormat, string msg, int digitNum, bool isAccount)
        {
            if (digitNum >= 0 && digitNum <= 2) // most common types
                switch (digitNum)
                {
                    case 2: // default
                        return numFormat;

                    case 1:
                        numFormat = isAccount
                            ? numFormat.Replace("##0.00", "##0.0").Replace("??", "?")
                            : numFormat.Replace("##0.00", "##0.0");
                        return numFormat;

                    case 0:
                        numFormat = isAccount
                            ? numFormat.Replace("##0.00", "##0").Replace("??", "")
                            : numFormat.Replace("##0.00", "##0");
                        return numFormat;
                }
            else
                Common.WriteLine("[Manual] Digit Number:" + msg);

            return numFormat;
        }

        public int GetDigitNum(string numFormat)
        {
            if (numFormat.Contains(";"))
            {
                var str = numFormat.Remove(numFormat.IndexOfAny(new[] {';'}));
                if (str.Contains("."))
                {
                    var digitNum = str.LastIndexOf('0') - str.IndexOf('.');
                    return digitNum;
                }

                return 0;
            }
            else
            {
                var str = numFormat;
                if (str.Contains("."))
                {
                    var digitNum = str.LastIndexOf('0') - str.IndexOf('.');
                    return digitNum;
                }

                return 0;
            }
        }

        public string ProcDateTime(string numberFormat)
        {
            return LocalLang.GetLocDate(numberFormat);
        }

        public void SelectAll()
        {
            try
            {
                rng = ws.UsedRange;
            }
            catch (Exception)
            {
                Common.WriteLine("[No available cell selected!]");
            }
        }

        public void SaveAndClose()
        {
            wb.Save();
            wb.Close(true);
            //wb = null;
        }

        public void Dispose()
        {
            if (rng != null)
            {
                Marshal.ReleaseComObject(rng);
                GC.WaitForPendingFinalizers();
            }

            if (ws != null)
            {
                Marshal.ReleaseComObject(ws);
                GC.WaitForPendingFinalizers();
            }

            if (wb != null)
            {
                SaveAndClose();
                Marshal.ReleaseComObject(wb);
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

        private void _ChangeFont(string sourceFont, EXCEL.Shape shape)
        {
            var targetFont = LocalLang.GetLocFont(sourceFont);

            if (shape.TextEffect.FontName != null) shape.TextEffect.FontName = targetFont;

            if (shape.TextFrame2.TextRange.Font.Name != null) shape.TextFrame2.TextRange.Font.Name = targetFont;

            if (LocalLang.IsFarEast) shape.TextFrame2.TextRange.Font.NameFarEast = targetFont;

            if (LocalLang.IsRightToLeft) //TODO: check Thailand
                shape.TextFrame2.TextRange.Font.NameComplexScript = targetFont;
        }

        private void _PrintShapeType(EXCEL._Worksheet worksheet)
        {
            if (worksheet != null)
                if (worksheet.Shapes.Count > 0)
                    foreach (EXCEL.Shape shape in worksheet.Shapes)
                        Common.WriteLine(shape.Type.ToString());
        }

        private void _PrintWorksheetType(EXCEL._Workbook workbook)
        {
            if (workbook != null)
                if (workbook.Worksheets.Count > 0)
                    foreach (EXCEL.Worksheet sheet in workbook.Worksheets)
                        Common.WriteLine(sheet.Type.ToString());
        }

        #endregion
    }
}