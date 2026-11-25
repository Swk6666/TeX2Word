using System;
using System.Windows;
using System.Windows.Media;

namespace TeX2Word {
    internal static class Theme {
        static ResourceDictionary _currentDict;
        public static bool IsDark { get; private set; }

        static SolidColorBrush Brush(string hex) {
            var color = (Color)ColorConverter.ConvertFromString(hex);
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        public static ResourceDictionary BuildDictionary(bool dark) {
            var res = new ResourceDictionary();
            if (dark) {
                res["CanvasBg"]          = Brush("#0B1120");
                res["CardBg"]            = Brush("#111827");
                res["CardBorder"]        = Brush("#1F2937");
                res["SubtleBg"]          = Brush("#162032");
                res["SubtleBorder"]      = Brush("#2A3952");
                res["RowBg"]             = Brush("#151F30");
                res["RowHoverBg"]        = Brush("#1C2940");
                res["RowBorder"]         = Brush("#24324A");
                res["DropOverlayBg"]     = Brush("#E60F172A");
                res["DividerBg"]         = Brush("#1F2937");

                res["TextPrimary"]       = Brush("#F8FAFC");
                res["TextSecondary"]     = Brush("#CBD5E1");
                res["TextMuted"]         = Brush("#64748B");

                res["AccentBg"]          = Brush("#3B82F6");
                res["AccentHoverBg"]     = Brush("#2563EB");
                res["AccentPressBg"]     = Brush("#1D4ED8");
                res["AccentDisabledBg"]  = Brush("#1E293B");
                res["AccentDisabledFg"]  = Brush("#475569");
                res["AccentSoftBg"]      = Brush("#172554");
                res["AccentSoftBorder"]  = Brush("#1E40AF");
                res["AccentText"]        = Brush("#93C5FD");

                res["ButtonBg"]          = Brush("#1E293B");
                res["ButtonHoverBg"]     = Brush("#273549");
                res["ButtonPressBg"]     = Brush("#334155");
                res["ButtonBorder"]      = Brush("#334155");
                res["GhostHoverBg"]      = Brush("#1E293B");

                res["BadgePdfBg"]        = Brush("#3B141A");
                res["BadgePdfFg"]        = Brush("#FCA5A5");
                res["BadgePdfBorder"]    = Brush("#7F1D1D");
                res["BadgeTexBg"]        = Brush("#1E1B4B");
                res["BadgeTexFg"]        = Brush("#A5B4FC");
                res["BadgeTexBorder"]    = Brush("#3730A3");
                res["BadgeMdBg"]         = Brush("#06392C");
                res["BadgeMdFg"]         = Brush("#6EE7B7");
                res["BadgeMdBorder"]     = Brush("#065F46");

                res["StatusPendingBg"]   = Brush("#1E293B");
                res["StatusPendingFg"]   = Brush("#94A3B8");
                res["StatusRunningBg"]   = Brush("#172554");
                res["StatusRunningFg"]   = Brush("#93C5FD");
                res["StatusDoneBg"]      = Brush("#052E1A");
                res["StatusDoneFg"]      = Brush("#6EE7B7");
                res["StatusFailedBg"]    = Brush("#3B141A");
                res["StatusFailedFg"]    = Brush("#FCA5A5");

                res["SuccessDot"]        = Brush("#10B981");
                res["WarnDot"]           = Brush("#F59E0B");
                res["LogBg"]             = Brush("#070B14");
                res["LogFg"]             = Brush("#CBD5E1");
                res["LogBorder"]         = Brush("#1E293B");
                res["ProgressTrack"]     = Brush("#1E293B");
                res["ScrollThumb"]       = Brush("#334155");
            } else {
                res["CanvasBg"]          = Brush("#F3F6FB");
                res["CardBg"]            = Brush("#FFFFFF");
                res["CardBorder"]        = Brush("#E2E8F0");
                res["SubtleBg"]          = Brush("#F8FAFC");
                res["SubtleBorder"]      = Brush("#E2E8F0");
                res["RowBg"]             = Brush("#FFFFFF");
                res["RowHoverBg"]        = Brush("#F8FAFC");
                res["RowBorder"]         = Brush("#E8EEF6");
                res["DropOverlayBg"]     = Brush("#EBEDF4FF");
                res["DividerBg"]         = Brush("#E2E8F0");

                res["TextPrimary"]       = Brush("#0F172A");
                res["TextSecondary"]     = Brush("#475569");
                res["TextMuted"]         = Brush("#94A3B8");

                res["AccentBg"]          = Brush("#2563EB");
                res["AccentHoverBg"]     = Brush("#1D4ED8");
                res["AccentPressBg"]     = Brush("#1E40AF");
                res["AccentDisabledBg"]  = Brush("#E2E8F0");
                res["AccentDisabledFg"]  = Brush("#94A3B8");
                res["AccentSoftBg"]      = Brush("#EFF6FF");
                res["AccentSoftBorder"]  = Brush("#BFDBFE");
                res["AccentText"]        = Brush("#2563EB");

                res["ButtonBg"]          = Brush("#FFFFFF");
                res["ButtonHoverBg"]     = Brush("#F1F5F9");
                res["ButtonPressBg"]     = Brush("#E2E8F0");
                res["ButtonBorder"]      = Brush("#CBD5E1");
                res["GhostHoverBg"]      = Brush("#F1F5F9");

                res["BadgePdfBg"]        = Brush("#FEF2F2");
                res["BadgePdfFg"]        = Brush("#DC2626");
                res["BadgePdfBorder"]    = Brush("#FECACA");
                res["BadgeTexBg"]        = Brush("#EEF2FF");
                res["BadgeTexFg"]        = Brush("#4F46E5");
                res["BadgeTexBorder"]    = Brush("#C7D2FE");
                res["BadgeMdBg"]         = Brush("#ECFDF5");
                res["BadgeMdFg"]         = Brush("#059669");
                res["BadgeMdBorder"]     = Brush("#A7F3D0");

                res["StatusPendingBg"]   = Brush("#F1F5F9");
                res["StatusPendingFg"]   = Brush("#475569");
                res["StatusRunningBg"]   = Brush("#EFF6FF");
                res["StatusRunningFg"]   = Brush("#2563EB");
                res["StatusDoneBg"]      = Brush("#ECFDF5");
                res["StatusDoneFg"]      = Brush("#047857");
                res["StatusFailedBg"]    = Brush("#FEF2F2");
                res["StatusFailedFg"]    = Brush("#DC2626");

                res["SuccessDot"]        = Brush("#10B981");
                res["WarnDot"]           = Brush("#D97706");
                res["LogBg"]             = Brush("#0F172A");
                res["LogFg"]             = Brush("#E2E8F0");
                res["LogBorder"]         = Brush("#1E293B");
                res["ProgressTrack"]     = Brush("#E2E8F0");
                res["ScrollThumb"]       = Brush("#CBD5E1");
            }
            return res;
        }

        public static void Apply(ResourceDictionary appResources, Window window, bool dark) {
            IsDark = dark;
            var next = BuildDictionary(dark);
            if (_currentDict != null) appResources.MergedDictionaries.Remove(_currentDict);
            appResources.MergedDictionaries.Add(next);
            _currentDict = next;
            ShellDialogs.ApplyTitleBarTheme(window, dark, dark ? 0x0020110B : 0x00FBF6F3);
        }
    }
}
