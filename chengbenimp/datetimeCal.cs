using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace chengbenimp
{
    public class datetimeCal
    {
        #region 获取 本周、本月、本季度、本年 的开始时间或结束时间
        /// <summary>
        /// 获取结束时间
        /// </summary>
        /// <param name="TimeType">Week、Month、Season、Year</param>
        /// <param name="now"></param>
        /// <returns></returns>
        public static DateTime? GetDateStartByType(string TimeType, DateTime now)
        {
            switch (TimeType)
            {
                case "Week":
                    return now.AddDays(-(int)now.DayOfWeek + 1);
                case "Month":
                    return now.AddDays(-now.Day + 1); //new DateTime(now.Year, now.Month, 1);
                case "Season":
                    var time = now.AddMonths(0 - ((now.Month - 1) % 3));
                    return time.AddDays(-time.Day + 1);
                case "Year":
                    return now.AddDays(-now.DayOfYear + 1);
                default:
                    return null;
            }
        }
        public static DateTime? GetDateTimeStartByType(string TimeType, DateTime now)
        {
            switch (TimeType)
            {
                case "Week":
                    var weektime = now.AddDays(-(int)now.DayOfWeek + 1);
                    return new DateTime(weektime.Year,weektime.Month,weektime.Day);
                case "Month":
                    return new DateTime(now.Year, now.Month, 1);
                case "Season":
                    var seasontime = now.AddMonths(0 - ((now.Month - 1) % 3));
                    seasontime = seasontime.AddDays(-seasontime.Day + 1);
                    return new DateTime(seasontime.Year, seasontime.Month, seasontime.Day);
                case "Year":
                    return new DateTime(now.Year, 1, 1);
                default:
                    return null;
            }
        }
        /// <summary>
        /// 获取结束时间
        /// </summary>
        /// <param name="TimeType">Week、Month、Season、Year</param>
        /// <param name="now"></param>
        /// <returns></returns>
        public static DateTime? GetDateEndByType(string TimeType, DateTime now)
        {
            switch (TimeType)
            {
                case "Week":
                    return now.AddDays(7 - (int)now.DayOfWeek);
                case "Month":
                    return now.AddMonths(1).AddDays(-now.AddMonths(1).Day + 1).AddDays(-1); //d1.AddMonths(1).AddDays(-1);
                case "Season":
                    var time = now.AddMonths((3 - ((now.Month - 1) % 3) - 1));
                    return time.AddMonths(1).AddDays(-time.AddMonths(1).Day + 1).AddDays(-1);
                case "Year":
                    var time2 = now.AddYears(1);
                    return time2.AddDays(-time2.DayOfYear);
                default:
                    return null;
            }
        }
        public static DateTime? GetDateTimeEndByType(string TimeType, DateTime now)
        {
            switch (TimeType)
            {
                case "Week":
                    var weektime = now.AddDays(7 - (int)now.DayOfWeek);
                    return new DateTime(weektime.Year,weektime.Month,weektime.Day,23,59,59);
                case "Month":
                    var monthtime =  now.AddMonths(1).AddDays(-now.AddMonths(1).Day + 1).AddDays(-1);
                    return new DateTime(monthtime.Year, monthtime.Month, monthtime.Day, 23, 59, 59);
                case "Season":
                    var time = now.AddMonths((3 - ((now.Month - 1) % 3) - 1));
                    time = time.AddMonths(1).AddDays(-time.AddMonths(1).Day + 1).AddDays(-1);
                    return new DateTime(time.Year, time.Month, time.Day, 23, 59, 59);
                case "Year":
                    var time2 = now.AddYears(1);
                    time2 = time2.AddDays(-time2.DayOfYear);
                    return new DateTime(time2.Year, time2.Month, time2.Day, 23, 59, 59);
                default:
                    return null;
            }
        }
        #endregion
        //Convert.ToDateTime(string)
    }
}
