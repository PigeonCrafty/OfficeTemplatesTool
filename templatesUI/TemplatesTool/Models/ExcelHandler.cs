using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Office.Core;
using TemplatesTool.Languages;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace TemplatesTool.Models
{
    public partial class ExcelHandler
    {
        #region Constructor

        public ExcelHandler(string filePath)
        {
            app = new EXCEL.Application {DisplayAlerts = false};
            LangName = Common.GetLangName(filePath);
            LocalLang = Common.GetLangObj(LangName.ToLower());
        }

        #endregion

        public void ExcelMain(string filePath)
        {
            try
            {
                wb = app.Workbooks.Open(filePath, misValue, false, misValue, misValue, misValue, true,
                    misValue, misValue, true, misValue, misValue, misValue);
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to Open]: " + filePath);
                app.Quit();
                return;
            }

            //_PrintWorksheetType(wb);

            UnprotectedAndUnhide();

            // 1. Process Styles of whole workbook
            foreach (EXCEL.Style style in wb.Styles) ProcStyles(style);

            // 2. Page size, Shapes, Textboxs and so on..
            foreach (EXCEL.Worksheet sheet in wb.Worksheets)
            {
                ws = sheet;
                Common.WriteLine("Sheet => " + ws.Name);

                // 2.1 set paper size to A4
                if (ws.PageSetup.PaperSize != EXCEL.XlPaperSize.xlPaperA4) SetPaperSize(ws);

                // 2.2 change text of shapes
                if (ws.Shapes.Count > 0)
                    foreach (EXCEL.Shape shape in ws.Shapes)
                        switch (shape.Type)
                        {
                            case MsoShapeType.msoAutoShape:
                            case MsoShapeType.msoTextBox:
                                ProcTextShape(shape);
                                break;

                            case MsoShapeType.msoGroup:
                                var group = shape.GroupItems;
                                ProcTextGroup(group);
                                break;

                            case MsoShapeType.msoChart:
                                ProcChart(shape);
                                break;
                        }

                SelectAll(); // set range for all activated cells

                // 3. Font of singel cell, currency, and so on..
                foreach (EXCEL.Range cell in rng)
                {
                    string numberFormat = cell.NumberFormat;
                    var cellAddress = cell.Address;
                    if (LocalLang.Id.Equals(0x0401)) cell.ReadingOrder = (int) EXCEL.Constants.xlRTL;

                    // Update cell font different with Styles used for CCJKAHT
                    try
                    {
                        UpdateCellFontFromStyle(cell);
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("[Failed] cell font " + cell.Address);
                    }

                    // 3.1 Change all style currency format
                    if (numberFormat.Contains("$") && numberFormat.Contains("#") &&
                        !numberFormat.Contains(LocalLang.CurrencySymbolLocal2) &&
                        !numberFormat.Contains(LocalLang.CurrencySymbolLocal3))
                    {
                        cell.NumberFormat = ProcCurrency(numberFormat, cellAddress);
                        continue;
                    }

                    // 3.2 Change date format
                    if ((numberFormat.Contains("d") || numberFormat.Contains("m") || numberFormat.Contains("y")) &&
                        !numberFormat.Contains("Red"))
                    {
                        if (numberFormat.StartsWith("*") || numberFormat.StartsWith("[$-x-sysdate]"))
                        {
                            Common.WriteLine("<System Setting Date> " + cellAddress + ": " + numberFormat);
                        }
                        else
                        {
                            var updatedFormatCode = ProcDateTime(numberFormat);

                            if (!updatedFormatCode.Equals(numberFormat)
                            ) // double validate in case of sometimes fail to correct
                                try
                                {
                                    cell.NumberFormat = updatedFormatCode;
                                }
                                catch (Exception)
                                {
                                    Common.WriteLine("[Failed] " + cell.Address);
                                }
                        }
                    }
                }
            }

            GetSheetStatusBack();

            // Save and Close
            try
            {
                //Console.WriteLine(LangName + "\\" + wb.Name + " Complete!");
                Dispose();
                LangName = "";
                _listHiddenWorkbook.Clear();
                _listProtectedWorkbook.Clear();
            }
            catch (Exception)
            {
                Common.WriteLine("Failed to save workbook!");
            }
        }

        #region Private Fields & Properties

        private EXCEL.Application app { get; }
        private EXCEL.Workbook wb { get; set; }
        private EXCEL.Worksheet ws { get; set; }
        private EXCEL.Range rng { get; set; }

        private string LangName { get; set; }
        private LocalLanguage LocalLang { get; }
        private readonly Missing misValue = Missing.Value;

        private readonly List<EXCEL.Worksheet> _listProtectedWorkbook = new List<EXCEL.Worksheet>();
        private readonly List<EXCEL.Worksheet> _listHiddenWorkbook = new List<EXCEL.Worksheet>();

        #endregion
    }
}