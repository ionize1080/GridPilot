using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;

namespace SheetPilot.Core
{
    [DataContract]
    public sealed class NavigationSettings : INotifyPropertyChanged
    {
        [DataMember] public int SchemaVersion { get; set; } = 1;
        public event PropertyChangedEventHandler PropertyChanged;
        private int dockSide = 0;
        [DataMember] public int DockSide { get => dockSide; set => Set(ref dockSide, Range(value, 0, 1)); }
        private int rowHeight = 36;
        [DataMember] public int RowHeight { get => rowHeight; set => Set(ref rowHeight, Range(value, 28, 80)); }
        private int rowGap = 2;
        [DataMember] public int RowGap { get => rowGap; set => Set(ref rowGap, Range(value, 0, 10)); }
        private int fontSize = 13;
        [DataMember] public int FontSize { get => fontSize; set => Set(ref fontSize, Range(value, 11, 18)); }
        private int workbookNameMode = 0;
        [DataMember] public int WorkbookNameMode { get => workbookNameMode; set => Set(ref workbookNameMode, Range(value, 0, 2)); }
        private int paneWidth = 320;
        [DataMember] public int PaneWidth { get => paneWidth; set => Set(ref paneWidth, Range(value, 260, 1600)); }
        private bool showTooltips = true;
        [DataMember] public bool ShowTooltips { get => showTooltips; set => Set(ref showTooltips, value); }
        private string backgroundColor = "#F7F9F7";
        [DataMember] public string BackgroundColor { get => backgroundColor; set { if (PreferenceInput.TryColor(value, out var color)) Set(ref backgroundColor, color); } }
        private string surfaceColor = "#FFFFFF";
        [DataMember] public string SurfaceColor { get => surfaceColor; set { if (PreferenceInput.TryColor(value, out var color)) Set(ref surfaceColor, color); } }
        private string textColor = "#294737";
        [DataMember] public string TextColor { get => textColor; set { if (PreferenceInput.TryColor(value, out var color)) Set(ref textColor, color); } }
        private string mutedTextColor = "#7C9582";
        [DataMember] public string MutedTextColor { get => mutedTextColor; set { if (PreferenceInput.TryColor(value, out var color)) Set(ref mutedTextColor, color); } }
        private string accentColor = "#23825A";
        [DataMember] public string AccentColor { get => accentColor; set { if (PreferenceInput.TryColor(value, out var color)) Set(ref accentColor, color); } }
        private string hoverColor = "#E2EEE7";
        [DataMember] public string HoverColor { get => hoverColor; set { if (PreferenceInput.TryColor(value, out var color)) Set(ref hoverColor, color); } }
        public void ApplyPalette(string palette)
        {
            HoverColor=palette=="dark"?"#31463F":palette=="neutral"?"#E4EAF4":"#E2EEE7";
            if (palette == "dark") { BackgroundColor = "#20262C"; SurfaceColor = "#2B333C"; TextColor = "#ECF0F3"; MutedTextColor = "#A3B2BE"; AccentColor = "#67CBA7"; }
            else if (palette == "neutral") { BackgroundColor = "#F5F6F8"; SurfaceColor = "#FFFFFF"; TextColor = "#293442"; MutedTextColor = "#758092"; AccentColor = "#537BC8"; }
            else { BackgroundColor = "#F7F9F7"; SurfaceColor = "#FFFFFF"; TextColor = "#294737"; MutedTextColor = "#7C9582"; AccentColor = "#23825A"; }
        }
        private bool showSheetNumbers = true;
        [DataMember] public bool ShowSheetNumbers { get => showSheetNumbers; set => Set(ref showSheetNumbers, value); }
        private bool showSheetType = false;
        [DataMember] public bool ShowSheetType { get => showSheetType; set => Set(ref showSheetType, value); }
        private bool showHeader = false;
        [DataMember] public bool ShowHeader { get => showHeader; set => Set(ref showHeader, value); }
        private bool showFooter = true;
        [DataMember] public bool ShowFooter { get => showFooter; set => Set(ref showFooter, value); }
        private bool showKeyboardHints = false;
        [DataMember] public bool ShowKeyboardHints { get => showKeyboardHints; set => Set(ref showKeyboardHints, value); }
        private bool autoRefresh = true;
        [DataMember] public bool AutoRefresh { get => autoRefresh; set => Set(ref autoRefresh, value); }
        private int refreshSeconds = 2;
        [DataMember] public int RefreshSeconds { get => refreshSeconds; set => Set(ref refreshSeconds, Range(value, 1, 10)); }
        private int scopeIndex = 0;
        [DataMember] public int ScopeIndex { get => scopeIndex; set => Set(ref scopeIndex, Range(value, 0, 1)); }
        private bool includeHidden = true;
        [DataMember] public bool IncludeHidden { get => includeHidden; set => Set(ref includeHidden, value); }
        private bool useRegex = false;
        [DataMember] public bool UseRegex { get => useRegex; set => Set(ref useRegex, value); }
        private bool matchCase = false;
        [DataMember] public bool MatchCase { get => matchCase; set => Set(ref matchCase, value); }
        private int quickSortMode;
        [DataMember] public int QuickSortMode { get => quickSortMode; set => Set(ref quickSortMode, Range(value,0,2)); }
        private bool quickSortDescending;
        [DataMember] public bool QuickSortDescending { get => quickSortDescending; set => Set(ref quickSortDescending,value); }
        [OnDeserializing] private void Initialize(StreamingContext context) { Reset(); }
        public void Reset()
        {
            SchemaVersion = 1;
            DockSide = 0;
            RowHeight = 36;
            RowGap = 2;
            FontSize = 13;
            WorkbookNameMode = 0;
            PaneWidth = 320;
            ShowTooltips = true;
            ApplyPalette("green");
            ShowSheetNumbers = true;
            ShowSheetType = false;
            ShowHeader = false;
            ShowFooter = true;
            ShowKeyboardHints = false;
            AutoRefresh = true;
            RefreshSeconds = 2;
            ScopeIndex = 0;
            IncludeHidden = true;
            UseRegex = false;
            MatchCase = false;
            QuickSortMode = 0;
            QuickSortDescending = false;
        }
        private static int Range(int value, int min, int max) => Math.Max(min, Math.Min(max, value));
        private void Set<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}

