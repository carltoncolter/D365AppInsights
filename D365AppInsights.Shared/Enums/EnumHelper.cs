using System;

namespace JLattimer.D365AppInsights
{
    public static class EnumHelper
    {
        public static T ParseEnum<T>(this object o)
        {
            try
            {
                return (T)Enum.Parse(typeof(T), o.ToString());
            }
            catch (Exception)
            {
                return default(T); //Information
            }
        }
    }
}