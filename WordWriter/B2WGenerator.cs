using CET.Controllers;
using CET.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NLog;
using WebSpiderDocs.Helpers;
using CET.Contact;

namespace WordWriter
{
    public class WordWriter
    {
        protected Logger logger = LogManager.GetCurrentClassLogger();

        protected SpiderDocHelper SDHelper;
        protected bool _incompletedBookmark = false;

        public WordWriter(SpiderDocHelper sdHelper = null)
        {
            SDHelper = sdHelper;
        }

        protected virtual void Setup()
        {
            //CertCode = certcode;
            if (SDHelper == null)
                SDHelper = new SpiderDocHelper(GlobalFactory.Instance4WebSpiderDocs(), new ConfigurationFactory.SpiderDocsConf());
        }

        /// <summary>
        /// Get Template for CompletionLetterCert.dot
        /// Original Location : "E:\Course Certificates - Fee For Service\INS11 Course\CompletionLetterCert.dot"
        /// SpiderDocs Location : Course Certificates - Fee For Service\INS11 Course\CompletionLetterCert.dot
        /// </summary>
        protected void RetriveTemplate(SpiderDocsModule.SearchCriteria criteria)
        {

            // Get a template file from SD
            SDHelper.GetByCriteria(criteria);

            // Put it onto temp folder 
            SDHelper.GetDownloadURL();

            //remain the path so that I can bind data later on.
            SDHelper.DownloadAll();
        }

        protected string BindDataToTemplate(BookmarkDB db)
        {
            //Dictionary<string, string> dic = db.AsDictionary();

            var doc = SDHelper.Docs.FirstOrDefault();    // Should be one
            // string dest = System.IO.Path.GetDirectoryName(doc.DownloadedPath) + "\\" + System.IO.Path.GetFileNameWithoutExtension(doc.DownloadedPath) + ".doc";

            // System.IO.File.Move(doc.DownloadedPath, dest);
            // doc.DownloadedPath = dest;
            // doc.DownloadURL = string.Join(".",doc.DownloadURL.Split('.').Take(doc.DownloadURL.Split('.').Length -1 )) + ".doc";

            WordWriter word = new WordWriter(doc.DownloadedPath);

            try
            {
                word.open();
                BookmarkDB certDb = new BookmarkDB(db).Create();

                foreach (KeyValuePair<string, string> pair in certDb.GetPrimitive().AsDictionary())
                {
                    logger.Debug("write to {0}:{1}", pair.Value, pair.Key);
                    word.write(text: pair.Value, where: pair.Key);
                }

                foreach (KeyValuePair<string, string> pair in certDb.GetTable().AsDictionary())
                {
                    int index = 0;
                    try
                    {
                        index = BookmarkDB.GetTableIndex(pair.Key);
                        word.write(tableIndex: index, row: BookmarkDB.GetTableRowIndex(pair.Key), texts: BookmarkDB.GetTableRow(pair.Value));
                    }
                    catch
                    {
                    }
                }
                
                throw new Exception("");
                // Footer                 
                try { word.WriteHeaderAndFooter(certDb, doc.Document.id_version); } catch { }

            }
            catch (Exception ex)
            {
                logger.Error(ex);

                throw ex;
            }
            finally
            {
                logger.Trace("Before SaveAsPDF");

                string url = word.SaveAsPDF();
                doc.DownloadURL = Fn.ToWebPath(Fn.ToWebURL(url));
                doc.DownloadedPath = url;
                doc.FileName = System.IO.Path.GetFileName(url);

                word.close();
            }

            return doc.DownloadedPath;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="criteria"></param>
        /// <param name="db"></param>
        /// <returns>System path</returns>
        public string Run(SpiderDocsModule.SearchCriteria criteria, BookmarkDB db)
        {

            RetriveTemplate(criteria);

            string systemPath = BindDataToTemplate(db);

            return systemPath;
        }

        protected virtual void Completed()
        {
        }
        public override string ToString()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public interface IGen
    {

        bool Setup(string suffix = "");

        void Gen();

        void Completed();
        string AsURL();

        bool HasInvalidBookmark();

        string GetNumber();
    }
}