//using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using NUnit.Framework;
using WordWriter;

namespace WordWriterTest
{
    [TestFixture]
    public class UnitTest1
    {
        WordWriter.WordWriter _word;
        BookmarkDB _db;
        [OneTimeSetUp]
		public void OneTimeSetup()
		{
			//

		}


		[SetUp]
		public void Setup()
		{
            //_word = new WordWriter.WordWriter();
            _db = new BookmarkDB();

            _db.Add("HeAder1", "Added! at Header1");
            _db.Add("FoOter1", "Added! at Footer1");
            _db.Add("BOdy1", "Added! at Body1");
            _db.Add("BodyTooLong1", $"Added! {TextOver300()} at Body1");
            _db.Add("BodyNotTooLong1", $"Added! {TextWithLen(80)} at Body1");
        }

        string TextBelow255()
        {
            return TextWithLen(254);
        }

        string TextOver300()
        {
            return TextWithLen(301);
        }
        string TextWithLen(int len)
        {
            string text = "";
            for (var i = 0; i < len; i++)
            {
                text += "A";
            }

            return text;
        }

        [Test(Description= "Bookmark and interpolation Bookmak")]
        [TestCase(@"C:\Users\takeo\source\repos\WordWriter\WordWriterTest\Source\BookMarkAsTestFull.dot", ExpectedResult = true)]
        public bool GivenTooLongStringThenWriteCorrectly(string path)
        {
            //_word = new WordWriter.WordWriter(path);

            WordWriter.WordWriter word = new WordWriter.WordWriter(path);

            word.open();
            word.Visible = true;
            BookmarkDB certDb = new BookmarkDB(_db).Create();

            foreach (KeyValuePair<string, string> pair in certDb.GetPrimitive().AsDictionary())
            {
                //logger.Debug("write to {0}:{1}", pair.Value, pair.Key);
                word.write(text: pair.Value, where: pair.Key);
            }

            //foreach (KeyValuePair<string, string> pair in certDb.GetTable().AsDictionary())
            //{
            //    int index = 0;
            //    try
            //    {
            //        index = BookmarkDB.GetTableIndex(pair.Key);
            //        word.write(tableIndex: index, row: BookmarkDB.GetTableRowIndex(pair.Key), texts: BookmarkDB.GetTableRow(pair.Value));
            //    }
            //    catch
            //    {
            //    }
            //}


            //logger.Trace("Before SaveAsPDF");
            
            string url = word.SaveAsPDF();

            word.close();

            return false;
        }


        [Test(Description = "Bookmark and interpolation Bookmak")]
        [TestCase(@"C:\Users\takeo\source\repos\WordWriter\WordWriterTest\Source\BookMarkAsTextOne.dot", ExpectedResult = true)]
        public bool GivenStringWithEOLThenWriteCorrectly(string path)
        {
            //_word = new WordWriter.WordWriter(path);
            // Arrange 
            BookmarkDB db = new BookmarkDB();
            db.Add("Test", @"
" + TextWithLen(250) + @"

OK
");

            WordWriter.WordWriter word = new WordWriter.WordWriter(path);

            word.open();
            word.Visible = true;


            BookmarkDB certDb = new BookmarkDB(db).Create();

            foreach (KeyValuePair<string, string> pair in certDb.GetPrimitive().AsDictionary())
            {
                //logger.Debug("write to {0}:{1}", pair.Value, pair.Key);
                word.write(text: pair.Value, where: pair.Key);
            }

            word.close();

            return false;
        }
        [Test(Description = "Bookmark and interpolation Bookmak")]
        [TestCase(@"C:\Users\takeo\source\repos\WordWriter\WordWriterTest\Source\BookMarkAsTextOne.dot", ExpectedResult = true)]
        public bool Given254StringThenWriteCorrectly(string path)
        {
            //_word = new WordWriter.WordWriter(path);
            // Arrange 
            BookmarkDB db = new BookmarkDB();
            db.Add("Test", $"{TextWithLen(254)}");

            WordWriter.WordWriter word = new WordWriter.WordWriter(path);

            word.open();
            word.Visible = true;


            BookmarkDB certDb = new BookmarkDB(db).Create();

            foreach (KeyValuePair<string, string> pair in certDb.GetPrimitive().AsDictionary())
            {
                //logger.Debug("write to {0}:{1}", pair.Value, pair.Key);
                word.write(text: pair.Value, where: pair.Key);
            }

            word.close();

            return false;
        }


        [Test(Description = "Bookmark and interpolation Bookmak")]
        [TestCase(@"C:\Users\takeo\source\repos\WordWriter\WordWriterTest\Source\BookMarkAsTextOne.dot", ExpectedResult = true)]
        public bool Given255StringThenWriteCorrectly(string path)
        {
            //_word = new WordWriter.WordWriter(path);
            // Arrange 
            _db = new BookmarkDB();
            _db.Add("Test", $"{TextWithLen(255)}");

            WordWriter.WordWriter word = new WordWriter.WordWriter(path);

            word.open();
            word.Visible = true;


            BookmarkDB certDb = new BookmarkDB(_db).Create();

            foreach (KeyValuePair<string, string> pair in certDb.GetPrimitive().AsDictionary())
            {
                //logger.Debug("write to {0}:{1}", pair.Value, pair.Key);
                word.write(text: pair.Value, where: pair.Key);
            }

            word.close();

            return false;
        }

        [Test(Description = "Bookmark and interpolation Bookmak")]
        [TestCase(@"C:\Users\takeo\source\repos\WordWriter\WordWriterTest\Source\BookMarkAsTextOne.dot", ExpectedResult = true)]
        public bool Given256StringThenWriteCorrectly(string path)
        {
            //_word = new WordWriter.WordWriter(path);
            // Arrange 
            _db = new BookmarkDB();
            _db.Add("Test", $"{TextWithLen(256)}");

            WordWriter.WordWriter word = new WordWriter.WordWriter(path);

            word.open();
            word.Visible = true;


            BookmarkDB certDb = new BookmarkDB(_db).Create();

            foreach (KeyValuePair<string, string> pair in certDb.GetPrimitive().AsDictionary())
            {
                //logger.Debug("write to {0}:{1}", pair.Value, pair.Key);
                word.write(text: pair.Value, where: pair.Key);
            }

            word.close();

            return false;
        }

    }
}
