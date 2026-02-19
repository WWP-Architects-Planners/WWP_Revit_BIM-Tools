using Microsoft.UI.Xaml;

namespace BEPDesigner.WinUI.Models;

public sealed class TopicOption
{
    public string Name { get; set; } = string.Empty;
    public bool Keep { get; set; } = true;
    public string GroupTitle { get; set; } = string.Empty;
    public bool ShowGroupHeader { get; set; }
    public Thickness IndentMargin { get; set; } = new(0);
    public Visibility GroupHeaderVisibility => ShowGroupHeader ? Visibility.Visible : Visibility.Collapsed;
}
