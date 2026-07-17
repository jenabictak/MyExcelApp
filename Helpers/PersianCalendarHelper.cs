using System.Globalization;

namespace MyExcelApp.Helpers
{
    public static class PersianCalendarHelper
    {
        private static readonly PersianCalendar _persianCalendar = new PersianCalendar();

        public static string ToPersianDateTime(this DateTime dateTime)
        {
            int year = _persianCalendar.GetYear(dateTime);
            int month = _persianCalendar.GetMonth(dateTime);
            int day = _persianCalendar.GetDayOfMonth(dateTime);
            int hour = _persianCalendar.GetHour(dateTime);
            int minute = _persianCalendar.GetMinute(dateTime);
            int second = _persianCalendar.GetSecond(dateTime);

            return $"{year:0000}/{month:00}/{day:00} {hour:00}:{minute:00}:{second:00}";
        }

        public static string ToPersianDate(this DateTime dateTime)
        {
            int year = _persianCalendar.GetYear(dateTime);
            int month = _persianCalendar.GetMonth(dateTime);
            int day = _persianCalendar.GetDayOfMonth(dateTime);

            return $"{year:0000}/{month:00}/{day:00}";
        }

        public static string ToPersianDateTimeShort(this DateTime dateTime)
        {
            int year = _persianCalendar.GetYear(dateTime);
            int month = _persianCalendar.GetMonth(dateTime);
            int day = _persianCalendar.GetDayOfMonth(dateTime);
            int hour = _persianCalendar.GetHour(dateTime);
            int minute = _persianCalendar.GetMinute(dateTime);

            return $"{year:0000}/{month:00}/{day:00} {hour:00}:{minute:00}";
        }
    }
}