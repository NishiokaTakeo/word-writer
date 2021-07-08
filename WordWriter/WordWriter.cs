using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;
using System.IO;
using WordWriter.Extensions;
//using CET.Controllers;
using NLog;

namespace WordWriter
{

    public class WordWriter : IDisposable
    {
        static Logger logger = LogManager.GetCurrentClassLogger();

        protected Word.Application App = null;
        protected Word.Document WordDoc = null;
        public string FilePath { get; set; }
        public bool Visible { set { App.Visible = value; } }
        public enum FormatType
        {
            None,
            Number
        }

        public WordWriter(string path)
        {
            FilePath = path;

            // Create a new instance of Word & make it visible
            App = new Word.Application();
            App.Visible = false;
        }

        public void open()
        {
            // file is now read only
            change_word_attr(true);

            WordDoc = App.Documents.Open(FilePath);
            WordDoc.Select();
        }

        protected void change_word_attr(bool writable = false)
        {
            // file is now read only
            FileAttributes attributes = File.GetAttributes(FilePath);
            if (writable) attributes &= ~FileAttributes.Normal; else attributes &= ~FileAttributes.ReadOnly;
            File.SetAttributes(FilePath, attributes);
        }
        public void close()
        {
            WordDoc.Save();
            WordDoc.Close();
            App.Quit();

        }
        /// <summary>
        /// type text at current selection
        /// </summary>
        /// <param name="text"></param>
        public void write(string text = "", string where = "", FormatType format = FormatType.Number)
        {
            where = where.TrimWhiteSpace();
            /*
            if (!string.IsNullOrWhiteSpace(where) && !WordDoc.Bookmarks.Exists(where))
                return;
            */

            // if bookmark is specifid then goto and write. otherwise just write at current cursol
            if (string.IsNullOrEmpty(where))
                App.Selection.TypeText(text);
            else
            {
                string t = Format(text, format);
                if (WordDoc.Bookmarks.Exists(where))
                {
                    App.Selection.GoTo(                    
                        What: Word.WdGoToItem.wdGoToBookmark,
                        Name: where);

                    App.Selection.TypeText(t);
                }

                // ALso Override
                //App.Selection.Range.Find.Execute("(" + where.Trim() + "}", false, true, false, false, false, true, 1, false, t, 2, false, false, false, false);

                Replace(where, t);

                //footer.Range.Find.Execute("{{FULL_NAME}}", false, true, false, false, false, true, 1, false, Name4Footer, 2, false, false, false, false);                
            }
        }

        public void Replace(string from, string to)
        {
            to = to.TrimWhiteSpace();
            from = from.Trim();

            string[] list = GetSameBookmark(from);

            foreach (string where in list)
            {
                logger.Trace("Starting replace doc with {0}:{1}", where,to);

                if (to.Length < 255)
                {
                    App.Selection.Range.Find.Execute(where, false, true, false, false, false, true, 1, false, to, 2, false, false, false, false);
                }
                else
                {
                    var textArray = to.SplitByLength(200);
                    var text = string.Empty;

                    for (var i = 0; i < (textArray.Count()-1);i ++)
                    {
                        text = textArray.ElementAt(i);
                        WordDoc.Content.Find.Execute(where, false, true, false, false, false, true, 1, false, text + where, 2, false, false, false, false);
                    }

                    text = textArray.Last();
                    WordDoc.Content.Find.Execute(where, false, true, false, false, false, true, 1, false, text, 2, false, false, false, false);
                }


                //App.Selection.Range.Find.Execute(where, false, true, false, false, false, true, 1, false, to, 2, false, false, false, false);


            }
        }

        public string[] GetSameBookmark(string name)
        {
            string[] list = new string[] {
                "{" + name.Trim() + "}",
                "{ " + name.Trim() + "}",
                "{" + name.Trim() + " }",
                "{ " + name.Trim() + " }"
            };

            return list;
        }

        //---------------------------------------------------------------------------------
        public void write(int tableIndex = 0, int row = 0, string[] texts = null, FormatType format = FormatType.Number)
        {
            var table = WordDoc.Tables[tableIndex];

            // Insert the data into the specific cell.
            for (int i = 1; i <= texts.Count(); i++)
            {
                string text = texts[i - 1].TrimWhiteSpace();
                text = Format(text, format);

                if (row > GetLastRowAtTable(tableIndex))
                    addRow(tableIndex);

                table.Cell(row, i).Range.InsertAfter(text);
            }
        }

        public void WriteHeaderAndFooter(BookmarkDB certDb, int spiderDocIdVersion)
        {
            var wordDocument = WordDoc;
            wordDocument.TrackRevisions = false; //Disable Tracking for the Field replacement operation

            List<Tuple<string, string>> list = new List<Tuple<string, string>>();

            foreach (KeyValuePair<string, string> pair in certDb.GetPrimitive().AsDictionary())
            {
                string[] listAux = GetSameBookmark(pair.Key);
                foreach (string at in listAux)
                    list.Add(Tuple.Create(at, pair.Value.TrimWhiteSpace()));
            }

            int headerDidNotFindCount = 0;
            int headerTotalCount = 0;
            int footerDidNotFindCount = 0;
            int footerTotalCount = 0;
            //bool bSkipHeader = LetterController.GetSkipHeaderAndFooterWordDoc(spiderDocIdVersion, "header");
            //bool bSkipFooter = LetterController.GetSkipHeaderAndFooterWordDoc(spiderDocIdVersion, "footer");

            foreach (Word.Section section in wordDocument.Sections)
            {
                //if (!bSkipHeader)
                    foreach (Word.HeaderFooter header in section.Headers)
                        foreach (Tuple<string, string> at in list)
                        {
                            headerTotalCount++;
                            if (!header.Range.Find.Execute(at.Item1, false, true, false, false, false, true, 1, false, at.Item2, 2, false, false, false, false))
                                headerDidNotFindCount++;
                        }

                //if (!bSkipFooter)
                    foreach (Word.HeaderFooter footer in section.Footers)
                        foreach (Tuple<string, string> at in list)
                        {
                            footerTotalCount++;
                            if (!footer.Range.Find.Execute(at.Item1, false, true, false, false, false, true, 1, false, at.Item2, 2, false, false, false, false))
                                footerDidNotFindCount++;
                        }
            }

            //if (!bSkipHeader && headerTotalCount == headerDidNotFindCount)
            //    LetterController.InsertSkipHeaderAndFooterWordDoc(spiderDocIdVersion, "header");

            //if (!bSkipFooter && footerTotalCount == footerDidNotFindCount)
            //    LetterController.InsertSkipHeaderAndFooterWordDoc(spiderDocIdVersion, "footer");
        }


        public int GetLastRowAtTable(int index)
        {
            return WordDoc.Tables[index].Rows.Count;
        }

        protected string Format(object any, FormatType format = FormatType.None)
        {

            //int tmp;
            //if (format == FormatType.None) return any.ToString();

            //if (format == FormatType.Number && int.TryParse(any.ToString().Trim(), out tmp))
            //    return String.Format("{0:0,0}", tmp);
            //else
            //    return any.ToString();

            return any.ToString();
        }

        public WordWriter addRow(int tableIndex = 1)
        {
            object m = System.Reflection.Missing.Value;

            var table = WordDoc.Tables[tableIndex];
            table.Rows.Add(ref m);
            table.Rows[table.Rows.Count].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;

            return this;
        }

        //---------------------------------------------------------------------------------

        protected void add_table(int table_index = 0, int rows = 0, params int[] columns)
        {
            //pTable.Format.SpaceAfter = 10f;
            WordDoc.Tables.Add(App.Selection.Range, NumRows: rows, NumColumns: columns.Length);

            for (int i = 1; i <= columns.Length; i++)
                WordDoc.Tables[table_index].Columns[i].SetWidth(columns[i - 1], Word.WdRulerStyle.wdAdjustNone);

            WordDoc.Tables[table_index].Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            WordDoc.Tables[table_index].Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;

            App.Selection.Delete();

        }
        public int next_table_index()
        {
            return WordDoc.Tables.Count + 1;
        }

        public string SaveAsDoc()
        {

            var sourceFile = new FileInfo(FilePath);

            string newFileName = sourceFile.FullName.Replace(".docm", ".docx").Replace(".dotx", ".docx").Replace(".dot", ".doc");
            newFileName = newFileName.Replace(".DOCM", ".docx").Replace(".DOTX", ".docx").Replace(".DOT", ".doc");

            WordDoc.SaveAs2(newFileName, Word.WdSaveFormat.wdFormatXMLDocument,
                             CompatibilityMode: Word.WdCompatibilityMode.wdWord2010);

            return newFileName;
        }

        public string SaveAsPDF()
        {

            var sourceFile = new FileInfo(FilePath);

            string newFileName = sourceFile.FullName.Replace(".docm", ".pdf").Replace(".dotx", ".pdf").Replace(".dot", ".pdf");
            newFileName = newFileName.Replace(".DOCM", ".pdf").Replace(".DOTX", ".pdf").Replace(".DOT", ".pdf");

            WordDoc.SaveAs2(newFileName, Word.WdSaveFormat.wdFormatPDF,
                             CompatibilityMode: Word.WdCompatibilityMode.wdWord2010);

            return newFileName;
        }


        //------------------------------------------

        public List<string> GetBookmarks()
        {
            var bookmarks = new List<string>();
            foreach (Word.Bookmark bookmark in WordDoc.Bookmarks)
            {
                bookmarks.Add(bookmark.Name.ToString());
            }

            var text = WordDoc.Content.Text;
            System.Text.RegularExpressions.MatchCollection matchList = System.Text.RegularExpressions.Regex.Matches(text, "{(.*?)}");
            var braceBookmarks = matchList.Cast<System.Text.RegularExpressions.Match>().Select(match => match.Value).ToList();
            foreach (var bookmark in braceBookmarks)
            {
                bookmarks.Add(bookmark.Replace("{", "").Replace("}", "").Trim());
            }


            foreach (Microsoft.Office.Interop.Word.Section section in WordDoc.Sections)
            {
                WordDoc.TrackRevisions = false; //Disable Tracking for the Field replacement operation

                Microsoft.Office.Interop.Word.HeadersFooters headers = section.Headers;
                foreach (Microsoft.Office.Interop.Word.HeaderFooter header in headers)
                {
                    var headerText = header.Range.Text;
                    System.Text.RegularExpressions.MatchCollection headerMatchList = System.Text.RegularExpressions.Regex.Matches(headerText, "{(.*?)}");
                    var headerBraceBookmarks = headerMatchList.Cast<System.Text.RegularExpressions.Match>().Select(match => match.Value).ToList();
                    foreach (var bookmark in headerBraceBookmarks)
                    {
                        bookmarks.Add(bookmark.Replace("{", "").Replace("}", "").Trim());
                    }
                }

                Microsoft.Office.Interop.Word.HeadersFooters footers = section.Footers;
                foreach (Microsoft.Office.Interop.Word.HeaderFooter footer in footers)
                {
                    var footerText = footer.Range.Text;
                    System.Text.RegularExpressions.MatchCollection footerMatchList = System.Text.RegularExpressions.Regex.Matches(footerText, "{(.*?)}");
                    var footerBraceBookmarks = footerMatchList.Cast<System.Text.RegularExpressions.Match>().Select(match => match.Value).ToList();
                    foreach (var bookmark in footerBraceBookmarks)
                    {
                        bookmarks.Add(bookmark.Replace("{", "").Replace("}", "").Trim());
                    }
                }
            }



            return bookmarks.Distinct().OrderBy(q => q).ToList();
        }

        private static bool FindAndReplaceFoods(Microsoft.Office.Interop.Word.Application wordApp, List<string> findTags, List<string> replaceTexts)
        {
            bool result = false;

            for (int i = 0; i < findTags.Count; i++)
            {
                object findText = findTags[i];
                object replaceWith = replaceTexts[i];
                result = FindAndReplace2(wordApp, findText, replaceWith);
            }

            return result;
        }

        private static bool FindAndReplace2(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            bool result = false;

            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllwordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object readOnly = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            string strNewReplaceWith = string.Empty;
            string[] strReplacements = replaceWithText.ToString().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < strReplacements.Length; i++)
            {
                if (i != strReplacements.Length - 1)
                {
                    strNewReplaceWith = string.Format("{0}\r\n{1}", strReplacements[i], findText);
                }
                else
                {
                    strNewReplaceWith = strReplacements[i];
                }
                object newReplaceWith = strNewReplaceWith;
                result = wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike, ref matchAllwordForms,
                ref forward, ref wrap, ref format, ref newReplaceWith, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza,
                ref matchControl);
            }

            return result;
        }

        public void Dispose()
        {

            WordDoc = null;
            App = null;


        }
    }
}
