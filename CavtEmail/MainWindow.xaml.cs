using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CavtEmail.Models;
using CavtEmail.ViewModels;

namespace CavtEmail;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private MainViewModel VM => (MainViewModel)DataContext;

    // ------------------------------------------------------------------
    // Attachments drag-and-drop
    // ------------------------------------------------------------------

    private void Attachments_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private void Attachments_Drop(object sender, DragEventArgs e)
    {
        if (VM.SelectedGroup == null) return;
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] items) return;

        foreach (var p in items)
        {
            // Folders are kept as references (resolved at send time); files are added directly.
            if (Directory.Exists(p) || File.Exists(p))
                VM.AddAttachmentPath(p);
        }
        e.Handled = true;
    }

    // ------------------------------------------------------------------
    // Token / row insertion
    // ------------------------------------------------------------------

    /// <summary>Where a row's Subj/Body button inserts text.</summary>
    private enum Target { Body, Subject }

    private static Target ParseTarget(object? tag) =>
        string.Equals(tag as string, "Subject", StringComparison.OrdinalIgnoreCase)
            ? Target.Subject
            : Target.Body;

    private void InsertAtCursor(TextBox tb, string text)
    {
        if (string.IsNullOrEmpty(text)) return;

        var caret = tb.CaretIndex;
        if (tb.SelectionLength > 0)
        {
            tb.SelectedText = text;
            tb.CaretIndex = caret + text.Length;
        }
        else
        {
            tb.Text = (tb.Text ?? "").Insert(caret, text);
            tb.CaretIndex = caret + text.Length;
        }
        tb.Focus();
    }

    private void Insert(Target target, string text)
    {
        var tb = target == Target.Subject ? SubjectBox : BodyBox;
        InsertAtCursor(tb, text);
    }

    private static Target TargetFromSender(object sender)
    {
        if (sender is FrameworkElement fe)
            return ParseTarget(fe.Tag);
        return Target.Body;
    }

    private void InsertRecipient_MenuClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.DataContext is not Recipient r) return;

        var full  = (r.Name ?? "").Trim();
        var first = Services.NameUtil.FirstName(r.Name);
        var last  = Services.NameUtil.LastName(r.Name);

        var menu = new ContextMenu { PlacementTarget = btn, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };

        void AddItem(string label, string text, Target tgt)
        {
            var mi = new MenuItem { Header = label };
            mi.Click += (_, _) => Insert(tgt, text);
            menu.Items.Add(mi);
        }

        void AddGroup(string component, string? text)
        {
            if (string.IsNullOrEmpty(text)) return;
            AddItem($"{component} \u2192 Body",    text, Target.Body);
            AddItem($"{component} \u2192 Subject", text, Target.Subject);
        }

        // Full name (fallback to email when Name is empty)
        AddGroup(string.IsNullOrEmpty(full) ? "Email" : "Full name",
                 string.IsNullOrEmpty(full) ? r.Email : full);

        if (!string.IsNullOrEmpty(first) && first != full)
        {
            menu.Items.Add(new Separator());
            AddGroup("First name", first);
        }
        if (!string.IsNullOrEmpty(last))
        {
            if (menu.Items.Count > 0 && menu.Items[^1] is not Separator)
                menu.Items.Add(new Separator());
            AddGroup("Last name", last);
        }

        if (menu.Items.Count == 0)
        {
            var empty = new MenuItem { Header = "(fill in Name or Email first)", IsEnabled = false };
            menu.Items.Add(empty);
        }

        menu.IsOpen = true;
    }

    private void InsertAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is AttachmentItem a)
        {
            Insert(TargetFromSender(sender), a.FileName);
        }
    }

    /// <summary>Token pills on the right panel; honors the Body/Subject radio choice.</summary>
    private void Token_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button b && b.Content is string s)
        {
            var target = TokenTargetSubject.IsChecked == true ? Target.Subject : Target.Body;
            Insert(target, s);
        }
    }

    // ------------------------------------------------------------------
    // Contact picker popup
    // ------------------------------------------------------------------

    private Recipient? _pickingFor;

    private void PickContact_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.DataContext is not Recipient r) return;
        _pickingFor = r;
        ContactPicker.PlacementTarget = btn;
        ContactPickerSearch.Text = "";
        RefreshContactPickerList("");
        ContactPicker.IsOpen = true;
        ContactPickerSearch.Focus();
    }

    private void RefreshContactPickerList(string query)
    {
        var all = VM.Contacts.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(query))
        {
            all = all.Where(c =>
                (c.Name?.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (c.Email?.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0));
        }
        ContactPickerList.ItemsSource = all
            .OrderBy(c => string.IsNullOrEmpty(c.Name) ? c.Email : c.Name,
                     StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (ContactPickerList.Items.Count > 0)
            ContactPickerList.SelectedIndex = 0;
    }

    private void ContactPickerSearch_TextChanged(object sender, TextChangedEventArgs e)
        => RefreshContactPickerList(ContactPickerSearch.Text?.Trim() ?? "");

    private void ContactPickerSearch_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) { ContactPicker.IsOpen = false; e.Handled = true; }
        else if (e.Key == Key.Down && ContactPickerList.Items.Count > 0)
        {
            ContactPickerList.Focus();
            if (ContactPickerList.SelectedIndex < 0) ContactPickerList.SelectedIndex = 0;
            if (ContactPickerList.ItemContainerGenerator.ContainerFromIndex(ContactPickerList.SelectedIndex) is ListBoxItem lbi)
                lbi.Focus();
            e.Handled = true;
        }
        else if (e.Key == Key.Enter)
        {
            CommitPick();
            e.Handled = true;
        }
    }

    private void ContactPickerList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) { ContactPicker.IsOpen = false; e.Handled = true; }
        else if (e.Key == Key.Enter) { CommitPick(); e.Handled = true; }
    }

    private void ContactPickerCommit(object sender, MouseButtonEventArgs e) => CommitPick();

    private void CommitPick()
    {
        if (_pickingFor == null) { ContactPicker.IsOpen = false; return; }
        if (ContactPickerList.SelectedItem is not Contact c) { ContactPicker.IsOpen = false; return; }
        _pickingFor.Name = c.Name;
        _pickingFor.Email = c.Email;
        ContactPicker.IsOpen = false;
    }
}