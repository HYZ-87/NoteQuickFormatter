using NoteQuickFormatter;

namespace NoteQuickFormatterTests.Services;

[TestClass]
public class DateTimeHelperTests
{
    [TestMethod]
    [DataRow(2025, 5, new string[] { "5/5-5/9", "5/12-5/16", "5/19-5/23", "5/26-5/30" })]
    public void GetWeekdayRanges_ReturnDatetimeRangeStringArray(int year, int month, string[] expected)
    {
        List<string> ranges = DateTimeHelper.GetWeekdayRanges(year, month);
        
        CollectionAssert.AreEqual(expected, ranges);
    }
}
