using System;
using System.Reflection;
using Microsoft.Office.Core;
using TemplatesTool.Languages;
using WORD = Microsoft.Office.Interop.Word;

namespace TemplatesTool.Models
{
    public partial class WordHandler
    {
        #region Constructor

        public WordHandler(string filePath)
        {
            app = new WORD.Application() {DisplayAlerts = WORD.WdAlertLevel.wdAlertsNone};
            LangName = Common.GetLangName(filePath);
            LocalLang = Common.GetLangObj(LangName.ToLower());
        }

        #endregion

        public void WordMain(string filePath)
        {
            try
            {
                //doc = app.Documents.Open(filePath);
                doc = app.Documents.Open(filePath, misValue, false,
                    misValue, misValue, misValue, misValue,
                    misValue, misValue, misValue, misValue,
                    false, false, //OpenAndRepair: True to repair the document to prevent document corruption.
                    misValue, true, misValue);
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to Open]: " + filePath);
                app.Quit();
                return;
            }

            // 1. Update Page Size to 'A4' if original size is 'Letter'
            //if (doc.PageSetup.PaperSize.Equals(WORD.WdPaperSize.wdPaperLetter))
            //if (doc.PageSetup.PaperSize == WORD.WdPaperSize.wdPaperLetter)
            //    try
            //    {
            //        ProcPageSize();
            //    }
            //    catch (Exception)
            //    {
            //        Common.WriteLine("<!> Error in changing Paper Size to A4");
            //        throw;
            //    }

            // 2. Update Styles Items font and language ID
            for (var i = 1; i <= doc.Styles.Count; i++)
                //Common.WriteLine("Styles " + i);
                //Common.WriteLine(doc.Styles[i].Type.ToString());
                if (doc.Styles[i].QuickStyle || doc.Styles[i].InUse)
                    try
                    {
                        ProcStyles(doc.Styles[i]);
                        doc.UpdateStyles();
                        //Common.WriteLine("Succeed");
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in changing Styles");
                    }

            // 3. Update all Control box font and language ID
            if (doc.ContentControls.Count > 0)
                for (var i = 1; i <= doc.ContentControls.Count; i++)
                    try
                    {
                        ProcContentControls(doc.ContentControls[i]);
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in changing Content Controls");
                    }

            // 4. Update headers and footers font and langauge ID 
            for (var i = 1; i <= doc.Sections.Count; i++)
                try
                {
                    var docAllHeaderFooterIndex = WORD.WdHeaderFooterIndex.wdHeaderFooterPrimary;
                    ProcHeadersFooters(doc.Sections[i].Headers[docAllHeaderFooterIndex],
                        doc.Sections[i].Footers[docAllHeaderFooterIndex]);
                }
                catch (Exception)
                {
                    Common.WriteLine("<!> Error in changing Content Controls");
                }

            // 5. Update shapes exist in document 
            if (doc.Shapes.Count > 0)

                for (var i = 1; i <= doc.Shapes.Count; i++)
                {
                    try
                    {
                        ProcShapes(doc.Shapes[i]);
                           
                        if (doc.Shapes[i].Type.Equals(MsoShapeType.msoGroup))
                            for (int j = 1; j <= doc.Shapes[i].GroupItems.Count; j++)
                            {
                                ProcShapes(doc.Shapes[i].GroupItems[j]);
                            }
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in changing Shapes");
                    }
                }

            // 6. Update tables in document
            if (doc.Tables.Count > 0)
            {
                for (int i = 1; i <= doc.Tables.Count; i++)
                for (int j = 1; j <= doc.Tables[i].Rows.Count; j++)
                for (int k = 1; k <= doc.Tables[i].Columns.Count; k++)
                    try
                    {
                        ProcTables(doc.Tables[i].Cell(j, k));
                    }
                    catch (Exception ex)
                    {
                        Common.WriteLine("<!> Error in changing Tables: " + ex.Message);
                    }
            }

            // 7. Update table of content
            if (doc.TablesOfContents.Count != 0)
                for (var i = 1; i <= doc.TablesOfContents.Count; i++)
                for (var j = 1; j <= doc.TablesOfContents[i].Range.Paragraphs.Count; j++)
                    try
                    {
                        ProcTableContent(doc.TablesOfContents[i].Range.Paragraphs[j]);
                        // Update new Table of Content
                        doc.TablesOfContents[1].UpdatePageNumbers();
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in changing Tables Content paragraphs");
                    }

            // 8. Update whole document content
            try
            {
                ProcWholeDoc();
            }
            catch (Exception)
            {
                Common.WriteLine("<!> Error in changing Whole Content");
            }

            // Save and close
            try
            {
                // Prompt complete info when each file processed
                //Console.WriteLine(LangName + "\\" + doc.Name + " Complete!" + "\n");
                LangName = "";
                Dispose();
            }
            catch (Exception ex)
            {
                Common.WriteLine("Error occurs during saving: " + ex.Message);
                throw;
            }
            finally
            {
                GC.Collect();
            }
        }

        #region Private Fields

        private WORD.Application app { get; }
        private WORD.Document doc { get; set; }

        private string LangName { get; set; }
        private LocalLanguage LocalLang { get; }

        private readonly Missing misValue = Missing.Value;

        #endregion
    }
}