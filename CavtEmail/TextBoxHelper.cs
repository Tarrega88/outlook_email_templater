using System.Windows;

namespace CavtEmail;

/// <summary>
/// Attached property used by the TextBox style to render placeholder text when empty.
/// Usage: &lt;TextBox local:TextBoxHelper.Placeholder="name@example.com" /&gt;
/// </summary>
public static class TextBoxHelper
{
    public static readonly DependencyProperty PlaceholderProperty =
        DependencyProperty.RegisterAttached(
            "Placeholder", typeof(string), typeof(TextBoxHelper),
            new PropertyMetadata(string.Empty));

    public static string GetPlaceholder(DependencyObject d) => (string)d.GetValue(PlaceholderProperty);
    public static void SetPlaceholder(DependencyObject d, string value) => d.SetValue(PlaceholderProperty, value);
}
