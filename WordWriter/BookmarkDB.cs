using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using System.Text.RegularExpressions;

namespace WordWriter
{

    #region Base Bookmark Db
    public class BookmarkDB : IDisposable
    {
        public class BDic : Dictionary<string, string>{}

        static protected Logger Logger = LogManager.GetCurrentClassLogger();

        #region Private Property
        BDic DB = new BDic();
        static readonly string SPLIT = "-";

        #endregion

        public BookmarkDB()
        {
            
        }
        public BookmarkDB(BookmarkDB db)
        {
            DB = db.GetDB();
        }

        #region Public Methods

        public void Clear()
        {
            DB = new BDic();
        }

        /// <summary>
        /// Create dictonary. Notice replacement is case sensitive, Be Aware.
        /// </summary>
        /// <returns></returns>
        public virtual BookmarkDB Create() 
        { 
            Add("bmdate", DateTime.Now.ToString("d MMMM yyyy"));

            return this; 
        }

        public BookmarkDB Merge(params BookmarkDB[] dbs)
        {
            foreach(var db in dbs)
            {
                foreach (var v in db.AsDictionary()) Add(v.Key, v.Value);
            }

            return this;
        }

        public BDic AsDictionary(BDic dic = null)
        {
            // protect original one
            BDic db = new BDic();

            // Additional. This call fist. so additional key is more high priority.
            if (dic != null) foreach (KeyValuePair<string, string> pair in dic) if (!db.ContainsKey(GetKey(pair.Key))) db.Add(GetKey(pair.Key), pair.Value);
            

            // Main. If key exists in additonal table, then no insert.
            foreach (KeyValuePair<string, string> pair in DB) if (!db.ContainsKey(GetKey(pair.Key))) db.Add(GetKey(pair.Key), pair.Value);
            
            return db;
        }


        public void Update(string key, object val)
        {
            if (val == null) val = string.Empty; 

            key = GetKey(key);

            if (!DB.ContainsKey(key))
                DB.Add(key, val.ToString().Trim());
            else
                DB[key] = val.ToString().Trim();
        }

        public void Add(string key, object val)
        {

            key = GetKey(key);

            if (val == null) val = string.Empty;

            if (!DB.ContainsKey(key)) DB.Add(key, val.ToString().Trim());
        }
        
        public BookmarkDB GetPrimitive()
        {
            BookmarkDB db = new BookmarkDB();

            foreach (var row in this.GetDB())
            {
                if (!isTableRow(row.Key)) db.Add(row.Key, row.Value);
            }

            return db;
        }

        public BookmarkDB GetTable()
        {
            BookmarkDB db = new BookmarkDB();

            foreach (var row in this.GetDB())
            {
                if (isTableRow(row.Key)) db.Add(row.Key, row.Value);
            }

            return db;
        }
        
        static public int GetTableIndex(string key)
        {
            int ok = 0;
            if (!int.TryParse(key.Substring(3, 1), out ok))
                Logger.Error("Table Index Could not find: {0}", key);

            return ok;
        }

        static public int GetTableRowIndex(string key)
        {
            Regex rgx = new Regex(@"^bmt[0-9]+R([0-9]+)_", RegexOptions.IgnoreCase);

            var matches = rgx.Matches(key.ToLower());

            if (matches.Count == 0 || matches[0].Groups.Count == 0)
            {
                Logger.Error("Bookmark Name format is wrong.");
                return 0;
            }

            int ok = 1;

            if (!int.TryParse(matches[0].Groups[1].Value, out ok))
                Logger.Error("Table Index Could not find: {0}", key);

            return ok;
        }

        static public string[] GetTableRow(string value)
        {
            return value.Split(new string[] { SPLIT }, StringSplitOptions.None).Select(x => DecodeSplit(x)).ToArray();
        }

        static public string DecodeSplit(string val)
        {
            //return val.Replace(SPLIT + "0", "-");
            return val.Replace("[[__ENC_INTERNAL_SPLIT__]]", "-");
            
        }

        #endregion
        protected BDic GetDB()
        {
            return DB;
        }

        static bool isTableRow(string key)
        {
            //int ok = 0;

            Regex rgx = new Regex(@"^bmt[0-9]+R[0-9]+_", RegexOptions.IgnoreCase);

            return rgx.IsMatch(key.ToLower());
        }

        string GetKey(string key)
        {
            return key.ToUpper();
        }


        void IDisposable.Dispose()
        {
    
        }
    }
    #endregion
}
