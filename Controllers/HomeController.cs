using Microsoft.AspNetCore.Mvc;
using SplitDocument.Models;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Compression;
using Syncfusion.Compression.Zip;

namespace SplitDocument.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult SplitBySection()
        {
            using(FileStream stream = new FileStream(Path.GetFullPath("Data/InputWithMultipleSections.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using(WordDocument document = new WordDocument(stream, FormatType.Docx))
                {
                    int i = 1;
                    var zip = new ZipArchive();
                    foreach(WSection section in document.Sections)
                    {
                        using(WordDocument newDocument = new WordDocument())
                        {
                            newDocument.Sections.Add(section.Clone());
                            SaveIntoZip(zip, "Section_" + i + ".docx", newDocument);
                            i++;
                        }
                    }
                    return ReturnZipFile(zip, "SplitBySection.zip");
                }
            }
        }
        private void SaveIntoZip(ZipArchive zip, string filename, WordDocument newDocument)
        {
            MemoryStream stream = new MemoryStream();
            newDocument.Save(stream, FormatType.Docx);
            ZipArchiveItem item = new ZipArchiveItem(zip, filename, stream, true, Syncfusion.Compression.FileAttributes.Compressed);
            zip.AddItem(item);
        }
        private FileContentResult ReturnZipFile(ZipArchive zip, string filename)
        {
            MemoryStream stream = new MemoryStream();
            zip.Save(stream, true);
            return File(stream.ToArray(), "application/zip", filename);
        }
        public IActionResult SplitByHeading()
        {
            using(FileStream stream = new FileStream(Path.GetFullPath("Data/InputWithHeadings.docx"), FileMode.Open, FileAccess.Read))
            {
                using(WordDocument document = new WordDocument(stream, FormatType.Docx))
                {
                    WordDocument newDocument = null;
                    WSection newSection = null;
                    var zip = new ZipArchive();
                    int i = 0;
                    foreach(WSection section in document.Sections)
                    {
                        if (newDocument != null)
                            newSection = AddSection(newDocument, section);
                        foreach(TextBodyItem item in section.Body.ChildEntities)
                        {
                            if (item is WParagraph)
                            {
                                WParagraph paragraph = item as WParagraph;
                                if (paragraph.StyleName == "Heading 1")
                                {
                                    if (newDocument != null)
                                    {
                                        SaveIntoZip(zip, "Heading" + i + ".docx", newDocument);
                                        i++;
                                    }
                                    newDocument = new WordDocument();
                                    newSection = AddSection(newDocument, section);
                                    AddEntity(newSection, paragraph);
                                }
                                else if (newDocument != null)
                                    AddEntity(newSection, paragraph);
                            }
                            else
                                AddEntity(newSection, item);
                        }
                    }
                    if (newDocument != null)
                        SaveIntoZip(zip, "Heading" + i + ".docx", newDocument);

                    return ReturnZipFile(zip, "SplitByHeading.zip");
                }
            }
        }
        private void AddEntity(WSection newSection, Entity entity)
        {
            newSection.Body.ChildEntities.Add(entity.Clone());
        }
        private WSection AddSection(WordDocument newDocument, WSection section)
        {
            WSection newSection = section.Clone();
            newSection.Body.ChildEntities.Clear();
            newSection.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
            newSection.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
            newSection.HeadersFooters.OddFooter.ChildEntities.Clear();
            newSection.HeadersFooters.OddHeader.ChildEntities.Clear();
            newSection.HeadersFooters.EvenFooter.ChildEntities.Clear();
            newSection.HeadersFooters.EvenHeader.ChildEntities.Clear();
            newDocument.Sections.Add(newSection);
            return newSection;
        }

        public IActionResult SplitByBookmark()
        {
            FileStream stream = new FileStream(Path.GetFullPath("Data/InputWithBookmarks.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using(WordDocument document = new WordDocument(stream, FormatType.Docx))
            {
                BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                BookmarkCollection bookmarkCollection = document.Bookmarks;
                var zip = new ZipArchive();
                foreach(Bookmark bookmark in bookmarkCollection)
                {
                    bookmarksNavigator.MoveToBookmark(bookmark.Name);
                    WordDocumentPart documentPart = bookmarksNavigator.GetContent();
                    WordDocument newDocument = documentPart.GetAsWordDocument();
                    SaveIntoZip(zip, bookmark.Name + ".docx", newDocument);
                }
                return ReturnZipFile(zip, "SplitByBookmark.zip");
            }
        }
        public IActionResult SplitByPlaceholder()
        {
            FileStream stream = new FileStream(Path.GetFullPath(@"Data/InputWithPlaceHolder.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using(WordDocument document = new WordDocument(stream, FormatType.Docx))
            {
                String[] placeholders = new string[] { "Chapter 1  Introducing C# and .NET", "End of Chapter 1", "Chapter 2  Coding Expressions and Statements ",
                    "End of Chapter 2", "Chapter 3  Methods and Properties ", "End of Chapter 3" };
                var zip = new ZipArchive();
                int bkmkId = 1;
                string bookmarkName = "";
                List<string> bookmarks = new List<string>();
                for(int i=0; i<placeholders.Length-1; i++)
                {
                    WParagraph startParagraph = document.Find(placeholders[i], true, true).GetAsOneRange().OwnerParagraph;
                    i++;
                    bookmarkName = "Bookmark_" + bkmkId;
                    bookmarks.Add(bookmarkName);
                    BookmarkStart bkmkStart = new BookmarkStart(document, bookmarkName);
                    startParagraph.ChildEntities.Insert(0, bkmkStart);
                    WParagraph endParagraph = document.Find(placeholders[i], true, true).GetAsOneRange().OwnerParagraph;
                    BookmarkEnd bkmkEnd = new BookmarkEnd(document, bookmarkName);
                    endParagraph.ChildEntities.Add(bkmkEnd);
                    bkmkId++;
                }
                BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                int id = 1;
                foreach(string bookmark in bookmarks)
                {
                    bookmarksNavigator.MoveToBookmark(bookmark);
                    WordDocumentPart wordDocumentPart = bookmarksNavigator.GetContent();
                    WordDocument newDocument = wordDocumentPart.GetAsWordDocument();
                    SaveIntoZip(zip, "Placeholder_" + id + ".docx", newDocument);
                    id++;
                }
                return ReturnZipFile(zip, "SplitByPlaceholder.zip");
            }
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}