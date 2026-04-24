using System.Globalization;
using System.Windows;

namespace CavtEmail;

public partial class AddContactDialog : Window
{
    public string Email { get; }
    public string FirstName => FirstBox.Text?.Trim() ?? "";
    public string LastName  => LastBox.Text?.Trim() ?? "";

    /// <summary>Combined display name — empty if both fields are blank.</summary>
    public string FullName
    {
        get
        {
            var f = FirstName; var l = LastName;
            if (f.Length == 0 && l.Length == 0) return "";
            if (f.Length == 0) return l;
            if (l.Length == 0) return f;
            return f + " " + l;
        }
    }

    public AddContactDialog(string email)
    {
        InitializeComponent();
        Email = email ?? "";
        EmailBox.Text = Email;
        (var first, var last) = GuessName(Email);
        FirstBox.Text = first;
        LastBox.Text = last;
        Loaded += (_, _) =>
        {
            // Focus the first empty name field so user can type immediately.
            if (string.IsNullOrEmpty(FirstBox.Text)) { FirstBox.Focus(); FirstBox.SelectAll(); }
            else { FirstBox.Focus(); FirstBox.SelectAll(); }
        };
    }

    /// <summary>
    /// Best-effort first/last guess from the email's local-part.
    /// "john.doe@x" -> ("John", "Doe"); "jdoe@x" or "sales@x" -> blanks.
    /// </summary>
    internal static (string first, string last) GuessName(string email)
    {
        if (string.IsNullOrWhiteSpace(email)) return ("", "");
        var at = email.IndexOf('@');
        var local = at > 0 ? email[..at] : email;
        var parts = local.Split(new[] { '.', '_', '-' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2) return ("", "");
        string Title(string s) => s.Length == 0
            ? s
            : CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLowerInvariant());
        return (Title(parts[0]), Title(parts[^1]));
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
