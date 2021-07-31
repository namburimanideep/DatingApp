using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace GlobalNetApps.Support.Common
{
    public class DateTimeMethods
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(typeof(DateTimeMethods));
        public DateTime GetFirstDayOfWeek(DateTime dayInWeek)
        {
            try
            {
                CultureInfo defaultCultureInfo = CultureInfo.CurrentCulture;
                DayOfWeek firstDay = defaultCultureInfo.DateTimeFormat.FirstDayOfWeek;
                DateTime firstDayInWeek = dayInWeek.Date;
                while (firstDayInWeek.DayOfWeek != firstDay)
                    firstDayInWeek = firstDayInWeek.AddDays(-1);
                return firstDayInWeek;
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
                throw(ex);
            }
        }
        public DateTime GetLastDayOfWeek(DateTime dayInWeek)
        {
            try
            {
                return GetFirstDayOfWeek(dayInWeek).AddDays(6);
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
                throw (ex);
            }
        }
    }
}