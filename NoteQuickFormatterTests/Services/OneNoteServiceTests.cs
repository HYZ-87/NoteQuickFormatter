using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoteQuickFormatter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NoteQuickFormatterTests.Services
{
    [TestClass()]
    public class OneNoteServiceTests
    {
        [TestMethod()]
        public void GetPageContentTest()
        {
            OneNoteService service = new OneNoteService();
            service.RefreshHierarchy();
            //service.GetPageContent()

            Assert.IsNotNull("");
        }
    }
}