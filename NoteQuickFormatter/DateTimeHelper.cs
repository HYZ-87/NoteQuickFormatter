using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NoteQuickFormatter
{
    public class DateTimeHelper
    {
        public static List<string> GetWeekdayRanges(int year, int month)
        {
            DateTime date = new DateTime(year, month, 1);
            do
            {
                if (date.DayOfWeek == DayOfWeek.Monday)
                {
                    break;
                }
                date = date.AddDays(1);
            } while (true);
            List<string> ranges = new List<string>();
            while (date.Month == month)
            {
                ranges.Add(string.Format("{0}-{1}", date.ToString("M/d"), date.AddDays(4).ToString("M/d")));
                date = date.AddDays(7);
            }
            return ranges;
        }

        public static DateTime ConvertStringToDateTime(string dateString)
        {
            int[] monthAndDay = dateString.Split('/').Select(s => Convert.ToInt32(s)).ToArray();
            return new DateTime(DateTime.Today.Year, monthAndDay[0], monthAndDay[1]);
        }
    }
}
