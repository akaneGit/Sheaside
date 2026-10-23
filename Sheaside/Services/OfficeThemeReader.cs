using System;
using System.IO;
using System.Security;
using Microsoft.Win32;

namespace Sheaside.Services
{
    internal enum OfficeTheme { System, Colorful, White, DarkGray, Black, LightGray }
    internal static class OfficeThemeReader
    {
        // Office exposes no supported theme-change event. This preference is best effort.
        internal static OfficeTheme Decode(int major, int? value, bool windowsDark = false)
        {
            if (major == 15)
                return value == 0 ? OfficeTheme.White : value == 1 ? OfficeTheme.LightGray : value == 2 ? OfficeTheme.DarkGray : OfficeTheme.System;
            if (major >= 16)
                return value == 0 ? OfficeTheme.Colorful : value == 3 ? OfficeTheme.DarkGray : value == 4 ? OfficeTheme.Black :
                    value == 5 ? OfficeTheme.White : value == 7 ? (windowsDark ? OfficeTheme.Black : OfficeTheme.White) : OfficeTheme.System;
            return OfficeTheme.System;
        }
        internal static OfficeTheme Read(string version)
        {
            int major;
            if (version == null || !int.TryParse(version.Split('.')[0], out major)) return OfficeTheme.System;
            try
            {
                using (var office = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + major + @".0\Common"))
                using (var windows = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    object value = office?.GetValue("UI Theme");
                    return Decode(major, value is int ? (int?)value : null, Equals(windows?.GetValue("AppsUseLightTheme"), 0));
                }
            }
            catch (Exception e) when (e is IOException || e is UnauthorizedAccessException || e is SecurityException)
            { return OfficeTheme.System; }
        }
    }
}
